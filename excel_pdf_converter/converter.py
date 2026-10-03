"""Microsoft Excel COM based PDF conversion service.

This module intentionally has no GUI dependencies.  It can be tested separately
and reused by a different interface later.
"""

from __future__ import annotations

import os
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

from .naming import available_pdf_path

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb"}
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_SHEET_VISIBLE = -1
XL_CALCULATION_MANUAL = -4135
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

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
    elapsed_seconds: float = 0.0

    @property
    def success_count(self) -> int:
        return len(self.pdf_files)


@dataclass(frozen=True)
class LocalPrinter:
    name: str
    port: str


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


def _find_local_pdf_printer(win32print: object) -> LocalPrinter | None:
    """Return a local Windows virtual printer and its exact port."""
    try:
        flags = win32print.PRINTER_ENUM_LOCAL  # type: ignore[attr-defined]
        printers = win32print.EnumPrinters(flags, None, 2)  # type: ignore[attr-defined]
    except Exception:
        return None

    printers_by_name = {
        str(printer.get("pPrinterName", "")).casefold(): LocalPrinter(
            name=str(printer.get("pPrinterName", "")),
            port=str(printer.get("pPortName", "")).split(",")[0].strip(),
        )
        for printer in printers
        if isinstance(printer, dict) and printer.get("pPrinterName")
    }
    for preferred_name in _LOCAL_PDF_PRINTERS:
        if preferred_name.casefold() in printers_by_name:
            return printers_by_name[preferred_name.casefold()]
    return None


def _prepare_local_printer(
    win32print: object,
) -> tuple[LocalPrinter, str | None, bool]:
    """Temporarily make a local virtual printer the Windows default.

    Excel consults the active printer again while opening a workbook and while
    exporting each sheet. The caller keeps this printer selected throughout.
    """
    printer = _find_local_pdf_printer(win32print)
    if not printer:
        raise ConversionUnavailableError(
            "PDF変換に必要な「Microsoft Print to PDF」が見つかりません。"
            "Windowsの設定でこの機能を有効にしてから、もう一度お試しください。"
        )

    try:
        original = str(win32print.GetDefaultPrinter())  # type: ignore[attr-defined]
    except Exception:
        original = None

    changed = original != printer.name
    if changed:
        try:
            win32print.SetDefaultPrinter(printer.name)  # type: ignore[attr-defined]
        except Exception as error:
            raise ConversionUnavailableError(
                "プリンターの接続待ちを回避できませんでした。"
                "Windowsの既定のプリンターを「Microsoft Print to PDF」にしてから、"
                "もう一度お試しください。"
            ) from error
    return printer, original, changed


def _ensure_local_printer(
    excel: object, win32print: object, printer: LocalPrinter
) -> None:
    """Keep Windows and the running Excel instance on the local PDF printer."""
    try:
        current_default = str(win32print.GetDefaultPrinter())  # type: ignore[attr-defined]
    except Exception:
        current_default = ""
    if current_default != printer.name:
        try:
            win32print.SetDefaultPrinter(printer.name)  # type: ignore[attr-defined]
        except Exception as error:
            raise ConversionUnavailableError(
                "PDF変換用のローカルプリンターを選択できませんでした。"
                "社内IT担当者へご相談ください。"
            ) from error

    # Excel's ActivePrinter value contains both the display name and port.
    # Keeping the Windows default selected also covers localized Office builds.
    if printer.port:
        active_printer = f"{printer.name} on {printer.port}"
        try:
            setattr(excel, "ActivePrinter", active_printer)
        except Exception:
            try:
                current_excel_printer = str(getattr(excel, "ActivePrinter"))
            except Exception:
                current_excel_printer = ""
            if printer.name.casefold() not in current_excel_printer.casefold():
                raise ConversionUnavailableError(
                    "ExcelにPDF変換用プリンターを設定できませんでした。"
                    "Windowsの既定のプリンターを「Microsoft Print to PDF」にしてから、"
                    "もう一度お試しください。"
                )


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


def _set_if_supported(target: object, name: str, value: object) -> None:
    """Set an optional Excel property without failing on older installations."""
    try:
        setattr(target, name, value)
    except Exception:
        pass


def _write_performance_log(lines: list[str]) -> None:
    """Write timing data for diagnosing slow workbooks on Windows."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        return
    try:
        log_dir = Path(local_app_data) / "ExcelPDFConverter"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "conversion.log"
        with log_path.open("a", encoding="utf-8") as log_file:
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_file.write(f"[{stamp}] " + " | ".join(lines) + "\n")
    except Exception:
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
    total_started = time.perf_counter()
    timing_lines: list[str] = []
    excel = None
    local_printer: LocalPrinter | None = None
    original_printer: str | None = None
    printer_changed = False
    pythoncom.CoInitialize()
    try:
        printer_started = time.perf_counter()
        local_printer, original_printer, printer_changed = _prepare_local_printer(
            win32print
        )
        timing_lines.append(f"printer={time.perf_counter() - printer_started:.2f}s")
        excel_started = time.perf_counter()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as error:
            raise ConversionUnavailableError(
                "Microsoft Excelを起動できませんでした。Excelが入っているか確認してください。"
            ) from error
        timing_lines.append(f"excel_start={time.perf_counter() - excel_started:.2f}s")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        _set_if_supported(excel, "ScreenUpdating", False)
        _set_if_supported(excel, "EnableEvents", False)
        _set_if_supported(excel, "DisplayStatusBar", False)
        _set_if_supported(excel, "Calculation", XL_CALCULATION_MANUAL)
        _set_if_supported(excel, "CalculateBeforeSave", False)
        _set_if_supported(
            excel, "AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        )
        _ensure_local_printer(excel, win32print, local_printer)
        total = len(paths)

        for index, workbook_path in enumerate(paths, start=1):
            workbook = None
            if progress:
                progress(index - 1, total, f"変換中：{workbook_path.name}")
            workbook_started = time.perf_counter()
            try:
                _ensure_local_printer(excel, win32print, local_printer)
                workbook = excel.Workbooks.Open(
                    str(workbook_path),
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True,
                    Notify=False,
                    AddToMru=False,
                )
                opened_at = time.perf_counter()
                visible_sheets = [
                    worksheet
                    for worksheet in workbook.Worksheets
                    if worksheet.Visible == XL_SHEET_VISIBLE
                ]
                for sheet_index, worksheet in enumerate(visible_sheets, start=1):
                    if progress:
                        progress(
                            index - 1,
                            total,
                            f"変換中：{workbook_path.name} "
                            f"（{sheet_index}/{len(visible_sheets)}シート）",
                        )
                    sheet_started = time.perf_counter()
                    _ensure_local_printer(excel, win32print, local_printer)
                    output_path = available_pdf_path(
                        output_dir, workbook_path.stem, str(worksheet.Name)
                    )
                    worksheet.ExportAsFixedFormat(
                        Type=XL_TYPE_PDF,
                        Filename=str(output_path),
                        Quality=XL_QUALITY_STANDARD,
                        IncludeDocProperties=False,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False,
                    )
                    result.pdf_files.append(output_path)
                    timing_lines.append(
                        f"workbook{index}/sheet{sheet_index}="
                        f"{time.perf_counter() - sheet_started:.2f}s"
                    )
                if not visible_sheets:
                    result.failures.append(
                        ConversionFailure(workbook_path, "表示されているシートがありません。")
                    )
                timing_lines.append(
                    f"open:workbook{index}={opened_at - workbook_started:.2f}s"
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
        result.elapsed_seconds = time.perf_counter() - total_started
        timing_lines.append(f"total={result.elapsed_seconds:.2f}s")
        _write_performance_log(timing_lines)

    return result
