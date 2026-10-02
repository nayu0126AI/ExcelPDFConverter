"""Microsoft Excel COM based PDF conversion service.

This module intentionally has no GUI dependencies.  It can be tested separately
and reused by a different interface later.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

from .naming import available_pdf_path

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb"}
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_SHEET_VISIBLE = -1

ProgressCallback = Callable[[int, int, str], None]

_LOCAL_PDF_PRINTERS = (
    "Microsoft Print to PDF",
    "Microsoft XPS Document Writer",
)


class ConversionUnavailableError(RuntimeError):
    """Raised when Microsoft Excel automation is unavailable."""


@dataclass(frozen=True)
class ConversionFailure:
    workbook: Path
    message: str


@dataclass
class ConversionResult:
    pdf_files: list[Path] = field(default_factory=list)
    failures: list[ConversionFailure] = field(default_factory=list)

    @property
    def success_count(self) -> int:
        return len(self.pdf_files)


def validate_inputs(files: Iterable[Path], output_dir: Path) -> list[Path]:
    """Validate and normalize user-selected paths before starting Excel."""
    normalized = [Path(path).resolve() for path in files]
    if not normalized:
        raise ValueError("Excelファイルを選んでください。")
    invalid = [path for path in normalized if path.suffix.lower() not in SUPPORTED_EXTENSIONS]
    if invalid:
        raise ValueError("対応していないファイルが含まれています。")
    missing = [path for path in normalized if not path.is_file()]
    if missing:
        raise ValueError("選択したファイルが見つかりません。もう一度選び直してください。")
    output_dir.mkdir(parents=True, exist_ok=True)
    return normalized


def _friendly_excel_error(error: Exception) -> str:
    text = str(error).lower()
    if "password" in text or "パスワード" in text:
        return "パスワードで保護されているため開けませんでした。"
    if "permission" in text or "access" in text or "アクセス" in text:
        return "ファイルを開けませんでした。Excelで開いている場合は閉じてください。"
    return "変換できませんでした。ファイルを閉じて、もう一度お試しください。"


def _find_local_pdf_printer(win32print: object) -> str | None:
    """Return a local Windows virtual printer suitable for Excel page layout."""
    try:
        flags = win32print.PRINTER_ENUM_LOCAL  # type: ignore[attr-defined]
        printers = win32print.EnumPrinters(flags, None, 2)  # type: ignore[attr-defined]
    except Exception:
        return None

    names = {
        str(printer.get("pPrinterName", "")).casefold(): str(
            printer.get("pPrinterName", "")
        )
        for printer in printers
        if isinstance(printer, dict) and printer.get("pPrinterName")
    }
    for preferred_name in _LOCAL_PDF_PRINTERS:
        if preferred_name.casefold() in names:
            return names[preferred_name.casefold()]
    return None


def _prepare_local_printer(win32print: object) -> tuple[str | None, bool]:
    """Temporarily make a local virtual printer the Windows default.

    Excel captures the default printer when it starts. The caller restores the
    user's original default immediately after Excel has started.
    """
    printer_name = _find_local_pdf_printer(win32print)
    if not printer_name:
        raise ConversionUnavailableError(
            "PDF変換に必要な「Microsoft Print to PDF」が見つかりません。"
            "Windowsの設定でこの機能を有効にしてから、もう一度お試しください。"
        )

    try:
        original = str(win32print.GetDefaultPrinter())  # type: ignore[attr-defined]
    except Exception:
        original = None

    changed = original != printer_name
    if changed:
        try:
            win32print.SetDefaultPrinter(printer_name)  # type: ignore[attr-defined]
        except Exception as error:
            raise ConversionUnavailableError(
                "プリンターの接続待ちを回避できませんでした。"
                "Windowsの既定のプリンターを「Microsoft Print to PDF」にしてから、"
                "もう一度お試しください。"
            ) from error
    return original, changed


def _restore_default_printer(
    win32print: object, original: str | None, changed: bool
) -> None:
    if not changed or not original:
        return
    try:
        win32print.SetDefaultPrinter(original)  # type: ignore[attr-defined]
    except Exception:
        # PDF conversion can continue. The app makes another restoration attempt
        # during final cleanup.
        pass


def convert_workbooks(
    files: Iterable[Path],
    output_dir: Path,
    progress: ProgressCallback | None = None,
) -> ConversionResult:
    """Convert every visible worksheet in each workbook to an individual PDF.

    The original workbook is opened read-only and is never saved. Existing PDFs
    are preserved; a numeric suffix is added to the newly generated file.
    """
    paths = validate_inputs(files, Path(output_dir).resolve())
    if sys.platform != "win32":
        raise ConversionUnavailableError(
            "この変換機能は、Microsoft Excelが入ったWindowsで使用できます。"
        )

    try:
        import pythoncom
        import win32print
        import win32com.client
    except ImportError as error:
        raise ConversionUnavailableError(
            "変換に必要な機能を読み込めませんでした。アプリを入れ直してください。"
        ) from error

    result = ConversionResult()
    excel = None
    original_printer: str | None = None
    printer_changed = False
    pythoncom.CoInitialize()
    try:
        original_printer, printer_changed = _prepare_local_printer(win32print)
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as error:
            raise ConversionUnavailableError(
                "Microsoft Excelを起動できませんでした。Excelが入っているか確認してください。"
            ) from error
        finally:
            _restore_default_printer(win32print, original_printer, printer_changed)

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        total = len(paths)

        for index, workbook_path in enumerate(paths, start=1):
            workbook = None
            if progress:
                progress(index - 1, total, f"変換中：{workbook_path.name}")
            try:
                workbook = excel.Workbooks.Open(
                    str(workbook_path),
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True,
                    Notify=False,
                )
                visible_count = 0
                for worksheet in workbook.Worksheets:
                    if worksheet.Visible != XL_SHEET_VISIBLE:
                        continue
                    visible_count += 1
                    output_path = available_pdf_path(
                        output_dir, workbook_path.stem, str(worksheet.Name)
                    )
                    worksheet.ExportAsFixedFormat(
                        Type=XL_TYPE_PDF,
                        Filename=str(output_path),
                        Quality=XL_QUALITY_STANDARD,
                        IncludeDocProperties=True,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False,
                    )
                    result.pdf_files.append(output_path)
                if visible_count == 0:
                    result.failures.append(
                        ConversionFailure(workbook_path, "表示されているシートがありません。")
                    )
            except Exception as error:
                result.failures.append(
                    ConversionFailure(workbook_path, _friendly_excel_error(error))
                )
            finally:
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
                    workbook = None
            if progress:
                progress(index, total, f"処理済み：{workbook_path.name}")
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None
        _restore_default_printer(win32print, original_printer, printer_changed)
        pythoncom.CoUninitialize()

    return result
