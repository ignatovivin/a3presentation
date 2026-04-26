$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location (Join-Path $root "frontend")

& "yarn.cmd" dev --host 127.0.0.1 --port 5173
