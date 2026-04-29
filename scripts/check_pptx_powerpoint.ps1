param(
    [string[]]$Path,
    [string]$Directory = "run-logs\selfcheck\pptx",
    [int]$Recent = 20
)

$ErrorActionPreference = "Continue"

if (-not $Path -or $Path.Count -eq 0) {
    $Path = Get-ChildItem -LiteralPath $Directory -Filter *.pptx |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First $Recent |
        ForEach-Object { $_.FullName }
}

$results = @()
foreach ($item in $Path) {
    $resolved = Resolve-Path -LiteralPath $item -ErrorAction SilentlyContinue
    if (-not $resolved) {
        $results += [pscustomobject]@{
            file = $item
            status = "missing"
            slides = $null
            error = "File not found"
        }
        continue
    }

    $ppt = $null
    $presentation = $null
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.DisplayAlerts = 1
        $presentation = $ppt.Presentations.Open($resolved.Path, -1, 0, 0)
        $results += [pscustomobject]@{
            file = Split-Path $resolved.Path -Leaf
            status = "opened"
            slides = $presentation.Slides.Count
            error = $null
        }
    } catch {
        $results += [pscustomobject]@{
            file = Split-Path $resolved.Path -Leaf
            status = "error"
            slides = $null
            error = $_.Exception.Message
        }
    } finally {
        if ($presentation) {
            try { $presentation.Close() } catch {}
        }
        if ($ppt) {
            try { $ppt.Quit() } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

$results | ConvertTo-Json -Depth 4
if ($results | Where-Object { $_.status -ne "opened" }) {
    exit 1
}
exit 0
