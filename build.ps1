$ErrorActionPreference = "Stop"

if (-not (Test-Path ".venv")) {
    py -m venv .venv
}

& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
& .\.venv\Scripts\python.exe -m pytest
& .\.venv\Scripts\python.exe -m PyInstaller --clean --noconfirm ExcelPDFConverter.spec

Write-Host ""
Write-Host "ビルドが完了しました: dist\ExcelPDFConverter.exe"

