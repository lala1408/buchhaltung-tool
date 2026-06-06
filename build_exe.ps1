$ErrorActionPreference = "Stop"

$python = ".\.venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    $python = "python"
}

& $python -m pip install --upgrade pip
& $python -m pip install -r requirements.txt
& $python -m pip install -r requirements-build.txt

& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name "SFK-Buchhaltung" `
    .\buchhaltung_launcher.py

Write-Host ""
Write-Host "Fertig: .\dist\SFK-Buchhaltung.exe"
