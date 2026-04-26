$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$logDir = Join-Path $root "run-logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

Start-Process powershell -ArgumentList @(
  "-NoLogo",
  "-NoExit",
  "-File",
  (Join-Path $PSScriptRoot "dev-backend.ps1")
) -WorkingDirectory $root `
  -RedirectStandardOutput (Join-Path $logDir "backend.out.log") `
  -RedirectStandardError (Join-Path $logDir "backend.err.log")

Start-Process powershell -ArgumentList @(
  "-NoLogo",
  "-NoExit",
  "-File",
  (Join-Path $PSScriptRoot "dev-frontend.ps1")
) -WorkingDirectory $root `
  -RedirectStandardOutput (Join-Path $logDir "frontend.out.log") `
  -RedirectStandardError (Join-Path $logDir "frontend.err.log")
