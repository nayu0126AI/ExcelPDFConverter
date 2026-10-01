"""PDF file-name helpers that are safe on Windows."""

from __future__ import annotations

import re
from pathlib import Path

_INVALID_WINDOWS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}


def safe_filename_part(value: str, fallback: str = "名称未設定") -> str:
    """Return a value that can safely be used as part of a Windows file name."""
    cleaned = _INVALID_WINDOWS_CHARS.sub("_", value).strip().rstrip(". ")
    if not cleaned:
        cleaned = fallback
    if cleaned.upper() in _RESERVED_NAMES:
        cleaned = f"_{cleaned}"
    return cleaned[:120].rstrip(". ") or fallback


def available_pdf_path(output_dir: Path, workbook_stem: str, sheet_name: str) -> Path:
    """Create a non-overwriting PDF path, adding a sequence number if needed."""
    base = f"{safe_filename_part(workbook_stem)}_{safe_filename_part(sheet_name)}"
    candidate = output_dir / f"{base}.pdf"
    number = 2
    while candidate.exists():
        candidate = output_dir / f"{base}_{number}.pdf"
        number += 1
    return candidate

