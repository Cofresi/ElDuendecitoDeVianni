$ErrorActionPreference = "Stop"

if (-not (Test-Path ".venv")) {
    py -3.12 -m venv .venv
}

.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\pyinstaller.exe `
    --noconfirm `
    --clean `
    --windowed `
    --onefile `
    --name "ElDuendecitoDeVianni" `
    --paths "src" `
    --add-data "demo;demo" `
    --add-data "config.example.json;." `
    launcher.py

if (Test-Path "politica_horario.xlsx") {
    Copy-Item "politica_horario.xlsx" "dist\politica_horario.xlsx" -Force
}

Write-Host "Executable generated at dist\ElDuendecitoDeVianni.exe"
