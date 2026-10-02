"""CustomTkinter user interface for the Excel PDF Converter."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import ctypes
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from .converter import (
    ConversionResult,
    ConversionUnavailableError,
    SUPPORTED_EXTENSIONS,
    convert_workbooks,
)

APP_TITLE = "Excel → PDF 変換"
ACCENT = "#2563EB"
ACCENT_HOVER = "#1D4ED8"
CONVERT = "#16803A"
CONVERT_HOVER = "#116B31"
TEXT = "#172033"
SUBTEXT = "#64748B"
BORDER = "#E2E8F0"
SURFACE = "#FFFFFF"
BACKGROUND = "#F4F7FB"
SUCCESS = "#16803A"
ERROR = "#C2413B"


def _resource_path(relative_path: str) -> Path:
    """Return a resource path both in source and in a PyInstaller bundle."""
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return bundle_root / relative_path


class ExcelPdfApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.selected_files: list[Path] = []
        self.output_dir: Path | None = None
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.is_converting = False

        self.title(APP_TITLE)
        if sys.platform == "win32":
            try:
                self.iconbitmap(str(_resource_path("assets/app_icon.ico")))
            except Exception:
                pass
        self.geometry("780x760")
        self.minsize(700, 700)
        self.configure(fg_color=BACKGROUND)
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self._build_ui()
        self.after(100, self._handle_worker_events)

    def _build_ui(self) -> None:
        shell = ctk.CTkFrame(self, fg_color="transparent")
        shell.pack(fill="both", expand=True, padx=46, pady=34)

        ctk.CTkLabel(
            shell,
            text=APP_TITLE,
            text_color=TEXT,
            font=ctk.CTkFont(size=30, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            shell,
            text="表示されているシートを、1シートずつPDFにします",
            text_color=SUBTEXT,
            font=ctk.CTkFont(size=14),
        ).pack(anchor="w", pady=(6, 24))

        card = ctk.CTkFrame(
            shell, fg_color=SURFACE, border_color=BORDER, border_width=1, corner_radius=14
        )
        card.pack(fill="both", expand=True)
        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=30, pady=28)

        self._step_title(content, "1", "Excelファイルを選ぶ")
        self.file_button = ctk.CTkButton(
            content,
            text="Excelファイルを選ぶ",
            command=self._choose_files,
            height=48,
            corner_radius=9,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.file_button.pack(fill="x", pady=(12, 10))
        self.file_summary = ctk.CTkLabel(
            content,
            text="まだ選択されていません",
            text_color=SUBTEXT,
            justify="left",
            anchor="w",
            wraplength=600,
            font=ctk.CTkFont(size=13),
        )
        self.file_summary.pack(fill="x", pady=(0, 24))

        self._separator(content)
        self._step_title(content, "2", "保存先を選ぶ")
        self.output_button = ctk.CTkButton(
            content,
            text="保存先を選ぶ",
            command=self._choose_output,
            height=44,
            corner_radius=9,
            fg_color="#FFFFFF",
            hover_color="#F1F5F9",
            border_color="#CBD5E1",
            border_width=1,
            text_color=TEXT,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.output_button.pack(fill="x", pady=(12, 10))
        self.output_summary = ctk.CTkLabel(
            content,
            text="まだ選択されていません",
            text_color=SUBTEXT,
            justify="left",
            anchor="w",
            wraplength=600,
            font=ctk.CTkFont(size=13),
        )
        self.output_summary.pack(fill="x", pady=(0, 24))

        self._separator(content)
        self._step_title(content, "3", "PDFに変換")
        self.convert_button = ctk.CTkButton(
            content,
            text="PDFに変換する",
            command=self._start_conversion,
            height=58,
            corner_radius=10,
            fg_color=CONVERT,
            hover_color=CONVERT_HOVER,
            font=ctk.CTkFont(size=17, weight="bold"),
            state="disabled",
        )
        self.convert_button.pack(fill="x", pady=(14, 16))

        self.progress_bar = ctk.CTkProgressBar(
            content, height=10, corner_radius=5, progress_color=CONVERT, fg_color="#E8EEF6"
        )
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            content,
            text="ファイルと保存先を選んでください",
            text_color=SUBTEXT,
            font=ctk.CTkFont(size=13),
            wraplength=600,
        )
        self.status_label.pack(pady=(12, 8))

        self.open_button = ctk.CTkButton(
            content,
            text="保存先を開く",
            command=self._open_output,
            height=42,
            corner_radius=9,
            fg_color="#E9F7EE",
            hover_color="#D9F0E2",
            text_color=SUCCESS,
            font=ctk.CTkFont(size=14, weight="bold"),
        )

        ctk.CTkLabel(
            shell,
            text="Microsoft Excel が入った Windows パソコンで使用できます",
            text_color="#94A3B8",
            font=ctk.CTkFont(size=12),
        ).pack(pady=(14, 0))

    @staticmethod
    def _separator(parent: ctk.CTkFrame) -> None:
        ctk.CTkFrame(parent, fg_color=BORDER, height=1).pack(fill="x", pady=(0, 24))

    @staticmethod
    def _step_title(parent: ctk.CTkFrame, number: str, text: str) -> None:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x")
        ctk.CTkLabel(
            row,
            text=number,
            width=30,
            height=30,
            corner_radius=15,
            fg_color="#E8F0FE",
            text_color=ACCENT,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            row,
            text=text,
            text_color=TEXT,
            font=ctk.CTkFont(size=17, weight="bold"),
        ).pack(side="left", padx=(10, 0))

    def _choose_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Excelファイルを選ぶ",
            filetypes=[
                ("Excelファイル", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("すべてのファイル", "*.*"),
            ],
        )
        if not paths:
            return
        self.selected_files = [
            Path(path) for path in paths if Path(path).suffix.lower() in SUPPORTED_EXTENSIONS
        ]
        names = [path.name for path in self.selected_files]
        shown = "、".join(names[:3])
        if len(names) > 3:
            shown += f" ほか{len(names) - 3}件"
        self.file_summary.configure(
            text=f"{len(names)}ファイル選択中\n{shown}" if names else "対応するファイルがありません",
            text_color=TEXT if names else ERROR,
        )
        self._refresh_ready_state()

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(title="PDFの保存先を選ぶ")
        if not path:
            return
        self.output_dir = Path(path)
        self.output_summary.configure(text=str(self.output_dir), text_color=TEXT)
        self._refresh_ready_state()

    def _refresh_ready_state(self) -> None:
        ready = bool(self.selected_files and self.output_dir and not self.is_converting)
        self.convert_button.configure(state="normal" if ready else "disabled")
        if ready:
            self.status_label.configure(text="準備できました", text_color=SUBTEXT)

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.file_button.configure(state=state)
        self.output_button.configure(state=state)
        self.convert_button.configure(state="normal" if enabled else "disabled")

    def _start_conversion(self) -> None:
        if not self.selected_files or not self.output_dir or self.is_converting:
            return
        self.is_converting = True
        self.open_button.pack_forget()
        self.progress_bar.set(0)
        self.status_label.configure(text="変換を開始しています…", text_color=CONVERT)
        self._set_controls_enabled(False)
        files = list(self.selected_files)
        output_dir = self.output_dir
        threading.Thread(
            target=self._conversion_worker, args=(files, output_dir), daemon=True
        ).start()

    def _conversion_worker(self, files: list[Path], output_dir: Path) -> None:
        def report(completed: int, total: int, message: str) -> None:
            self.events.put(("progress", (completed, total, message)))

        try:
            result = convert_workbooks(files, output_dir, report)
            self.events.put(("done", result))
        except (ValueError, ConversionUnavailableError) as error:
            self.events.put(("error", str(error)))
        except Exception:
            self.events.put(
                ("error", "予期しない問題が起きました。Excelを閉じて、もう一度お試しください。")
            )

    def _handle_worker_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "progress":
                    completed, total, message = payload  # type: ignore[misc]
                    self.progress_bar.set(completed / total if total else 0)
                    self.status_label.configure(text=message, text_color=CONVERT)
                elif event == "done":
                    self._show_result(payload)  # type: ignore[arg-type]
                elif event == "error":
                    self._show_error(str(payload))
        except queue.Empty:
            pass
        finally:
            self.after(100, self._handle_worker_events)

    def _show_result(self, result: ConversionResult) -> None:
        self.is_converting = False
        self.progress_bar.set(1)
        self._set_controls_enabled(True)
        self._refresh_ready_state()
        if result.failures:
            message = (
                f"{result.success_count}件のPDFを作成しました。"
                f"{len(result.failures)}ファイルは変換できませんでした。"
            )
            self.status_label.configure(text=message, text_color=ERROR)
            details = "\n".join(
                f"・{failure.workbook.name}：{failure.message}" for failure in result.failures
            )
            messagebox.showwarning("一部の変換が完了しませんでした", f"{message}\n\n{details}")
        else:
            message = (
                f"完了しました。{result.success_count}件のPDFを作成しました。"
                f"（{result.elapsed_seconds:.1f}秒）"
            )
            self.status_label.configure(text=message, text_color=SUCCESS)
        self.open_button.pack(fill="x", pady=(8, 0))

    def _show_error(self, message: str) -> None:
        self.is_converting = False
        self.progress_bar.set(0)
        self._set_controls_enabled(True)
        self._refresh_ready_state()
        self.status_label.configure(text=message, text_color=ERROR)
        messagebox.showerror("変換できませんでした", message)

    def _open_output(self) -> None:
        if not self.output_dir:
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(self.output_dir))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(self.output_dir)])
            else:
                subprocess.Popen(["xdg-open", str(self.output_dir)])
        except Exception:
            messagebox.showinfo("保存先", str(self.output_dir))


def main() -> None:
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(  # type: ignore[attr-defined]
                "InternalTools.ExcelPDFConverter.1.1"
            )
        except Exception:
            pass
    app = ExcelPdfApp()
    app.mainloop()
