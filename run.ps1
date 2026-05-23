# Levanta la app web (Windows / PowerShell)
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
if (-not (Test-Path ".venv")) { py -m venv .venv }
& ".\.venv\Scripts\python.exe" -m pip install --quiet -r app\backend\requirements.txt
if (-not (Test-Path ".env")) { Copy-Item ".env.example" ".env"; Write-Host "Creé .env: completá GEMINI_API_KEY." -ForegroundColor Yellow }
& ".\.venv\Scripts\python.exe" -m uvicorn main:app --reload --port 8000 --app-dir app\backend
