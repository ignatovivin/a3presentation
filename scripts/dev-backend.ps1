$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

& ".\.venv\Scripts\python.exe" -m uvicorn a3presentation.main:app `
  --app-dir src `
  --host 127.0.0.1 `
  --port 8000 `
  --reload `
  --reload-dir src `
  --reload-dir storage\templates
