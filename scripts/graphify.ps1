Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$graphifyExe = Join-Path $repoRoot ".venv\\Scripts\\graphify.exe"

if (-not (Test-Path $graphifyExe)) {
    throw "graphify executable not found at $graphifyExe"
}

& $graphifyExe @args
exit $LASTEXITCODE
