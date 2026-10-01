from pathlib import Path

from excel_pdf_converter.naming import available_pdf_path, safe_filename_part


def test_safe_filename_replaces_windows_invalid_characters() -> None:
    assert safe_filename_part('売上/東京:*?"<>|') == "売上_東京_______"


def test_safe_filename_handles_reserved_names() -> None:
    assert safe_filename_part("CON") == "_CON"


def test_available_pdf_path_does_not_overwrite(tmp_path: Path) -> None:
    first = tmp_path / "帳票_明細.pdf"
    first.touch()
    second = available_pdf_path(tmp_path, "帳票", "明細")
    assert second.name == "帳票_明細_2.pdf"

