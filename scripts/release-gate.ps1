$IncludeVisual = $false
if ($args -contains "-IncludeVisual") {
  $IncludeVisual = $true
}

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Invoke-Step {
  param(
    [string]$Title,
    [string]$FilePath,
    [string[]]$Arguments,
    [string]$WorkingDirectory = $root
  )

  Write-Host $Title
  Push-Location $WorkingDirectory
  try {
    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
      throw "$Title failed with exit code $LASTEXITCODE"
    }
  }
  finally {
    Pop-Location
  }
}

Invoke-Step "[release-gate] API contracts" ".\.venv\Scripts\python.exe" @("-m", "pytest", "tests\test_api.py")
Invoke-Step "[release-gate] Frontend verify" "cmd.exe" @("/c", "yarn", "verify") (Join-Path $root "frontend")
Invoke-Step "[release-gate] Frontend smoke" "cmd.exe" @("/c", "yarn", "test:smoke", "--workers", "1") (Join-Path $root "frontend")
Invoke-Step "[release-gate] Frontend runtime" "cmd.exe" @("/c", "yarn", "test:runtime") (Join-Path $root "frontend")

if ($IncludeVisual) {
  Invoke-Step "[release-gate] Frontend visual approval" "cmd.exe" @("/c", "yarn", "test:visual", "--workers", "1") (Join-Path $root "frontend")
}

Write-Host "[release-gate] Done"
