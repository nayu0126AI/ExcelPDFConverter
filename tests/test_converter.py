from pathlib import Path
from types import ModuleType, SimpleNamespace

import pytest

import excel_pdf_converter.converter as converter
from excel_pdf_converter.converter import _find_local_pdf_printer, validate_inputs


def test_validate_inputs_accepts_excel_and_creates_output(tmp_path: Path) -> None:
    source = tmp_path / "sample.xlsx"
    source.touch()
    output = tmp_path / "pdf"
    assert validate_inputs([source], output) == [source.resolve()]
    assert output.is_dir()


def test_validate_inputs_rejects_non_excel(tmp_path: Path) -> None:
    source = tmp_path / "sample.txt"
    source.touch()
    with pytest.raises(ValueError, match="対応していない"):
        validate_inputs([source], tmp_path / "pdf")


def test_find_local_pdf_printer_prefers_microsoft_pdf() -> None:
    fake = SimpleNamespace(
        PRINTER_ENUM_LOCAL=2,
        EnumPrinters=lambda *_: [
            {"pPrinterName": "Microsoft XPS Document Writer"},
            {"pPrinterName": "Microsoft Print to PDF"},
        ],
    )
    assert _find_local_pdf_printer(fake) == "Microsoft Print to PDF"


def test_convert_exports_only_visible_sheets_and_closes_excel(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "sample.xlsx"
    source.touch()
    output = tmp_path / "pdf"
    exported: list[str] = []

    class FakeSheet:
        def __init__(self, name: str, visible: int) -> None:
            self.Name = name
            self.Visible = visible

        def ExportAsFixedFormat(self, **kwargs: object) -> None:
            exported.append(Path(str(kwargs["Filename"])).name)

    class FakeWorkbook:
        Worksheets = [FakeSheet("表示", -1), FakeSheet("非表示", 0)]
        closed = False

        def Close(self, SaveChanges: bool) -> None:
            assert SaveChanges is False
            self.closed = True

    workbook = FakeWorkbook()

    class FakeWorkbooks:
        def Open(self, filename: str, **kwargs: object) -> FakeWorkbook:
            assert filename == str(source.resolve())
            assert kwargs["ReadOnly"] is True
            return workbook

    class FakeExcel:
        def __init__(self) -> None:
            self.Workbooks = FakeWorkbooks()
            self.quit_called = False

        def Quit(self) -> None:
            self.quit_called = True

    excel = FakeExcel()
    pythoncom = ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: None  # type: ignore[attr-defined]
    pythoncom.CoUninitialize = lambda: None  # type: ignore[attr-defined]
    client = ModuleType("win32com.client")
    client.DispatchEx = lambda _: excel  # type: ignore[attr-defined]
    win32com = ModuleType("win32com")
    win32com.client = client  # type: ignore[attr-defined]
    printer_changes: list[str] = []
    win32print = ModuleType("win32print")
    win32print.PRINTER_ENUM_LOCAL = 2  # type: ignore[attr-defined]
    win32print.EnumPrinters = lambda *_: [  # type: ignore[attr-defined]
        {"pPrinterName": "Microsoft Print to PDF"}
    ]
    win32print.GetDefaultPrinter = lambda: "社内ネットワークプリンター"  # type: ignore[attr-defined]
    win32print.SetDefaultPrinter = printer_changes.append  # type: ignore[attr-defined]

    monkeypatch.setattr(converter.sys, "platform", "win32")
    monkeypatch.setitem(converter.sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(converter.sys.modules, "win32com", win32com)
    monkeypatch.setitem(converter.sys.modules, "win32com.client", client)
    monkeypatch.setitem(converter.sys.modules, "win32print", win32print)

    result = converter.convert_workbooks([source], output)

    assert exported == ["sample_表示.pdf"]
    assert result.success_count == 1
    assert result.elapsed_seconds >= 0
    assert not result.failures
    assert workbook.closed is True
    assert excel.quit_called is True
    assert printer_changes == [
        "Microsoft Print to PDF",
        "社内ネットワークプリンター",
        "社内ネットワークプリンター",
    ]
