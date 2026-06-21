# Task 7 local ingest manual verification (throwaway script)
# - Drives the split pipeline via IngestRunner.Run in a clean PowerPoint COM instance
#   (no source/add-in changes). String-only entry point avoids __ComObject marshalling.
# - Prints the resulting plan + per-item debug.log progress, then verifies the outputs.
#
# Usage (no admin; close ALL PowerPoint first):
#   powershell -ExecutionPolicy Bypass -File scripts\manual-verify-task7.ps1 -Bundle "C:\Projects\teamppt-addin\assets\layout_test_aseet.pptx"

param(
    [Parameter(Mandatory = $true)]
    [string]$Bundle,
    [string]$OutDir = "C:\Projects\teamppt-addin\test"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Bundle)) { throw "Bundle pptx not found: $Bundle" }
$dll = "C:\Projects\teamppt-addin\src\TeampptAddin\bin\Debug\TeampptAddin.dll"
if (-not (Test-Path $dll)) { throw "DLL not found: $dll  (build first)" }

# Isolation guard: refuse to run if PowerPoint is already open (attaching to a live
# add-in instance multiplies panels and crashes the ingest).
$running = Get-Process POWERPNT -ErrorAction SilentlyContinue
if ($running) {
    Write-Host "[X] PowerPoint is running. Close ALL PowerPoint windows and re-run." -ForegroundColor Red
    exit 2
}

# Reset output folder
if (Test-Path $OutDir) { Remove-Item "$OutDir\*" -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

Write-Host "Bundle : $Bundle"
Write-Host "Output : $OutDir`n"

# Clean PowerPoint COM instance + load DLL, then drive the ingest through the
# string-only entry point. We do NOT call the adapters (SectionReader/SlideSplitter/...)
# from PowerShell: they take a strongly-typed PowerPoint.Presentation, and PowerShell
# cannot marshal a New-Object __ComObject into that parameter. IngestRunner.Run takes
# two strings, opens/closes its own Presentation internally, and returns the plan.
$ppt = New-Object -ComObject PowerPoint.Application
[void][Reflection.Assembly]::LoadFrom($dll)
[TeampptAddin.Globals]::Application = $ppt

# debug.log carries per-item ingest progress. Remember where it ends now so we can
# tail only the lines this run appends.
$logPath  = Join-Path $env:LOCALAPPDATA "TeampptAddin\debug.log"
$logStart = 0
if (Test-Path $logPath) { $logStart = (Get-Content $logPath -ErrorAction SilentlyContinue | Measure-Object -Line).Lines }

$plan = $null
Write-Host "=== RUN ===" -ForegroundColor Cyan
Write-Host "Calling IngestRunner.Run (string args only -> no COM marshalling)...`n"
try {
    $plan = [TeampptAddin.IngestRunner]::Run($Bundle, $OutDir)
}
catch {
    $ppt.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    throw
}

$total = $plan.Count

# Plan that Run produced (one item == one slide == one asset).
Write-Host "=== PLAN ===" -ForegroundColor Cyan
Write-Host ("{0} slides -> {0} pptx + {0} png ({1} files) expected`n" -f $total, ($total * 2))
foreach ($it in $plan) {
    Write-Host ("  - {0,-30} (slide {1}) -> {2} + {3}" -f $it.AssetId, $it.SourceSlideIndex, $it.PptxFileName, $it.ThumbFileName)
}

# Per-item progress IngestRunner logged to debug.log during this run.
Write-Host "`n=== debug.log (this run) ===" -ForegroundColor Cyan
if (Test-Path $logPath) {
    Get-Content $logPath |
        Select-Object -Skip $logStart |
        Where-Object { $_ -match "Ingest:" } |
        ForEach-Object { Write-Host ("  {0}" -f $_) }
}

# 3) Verify outputs (each pptx == 1 slide? png long edge == 768?)
Write-Host "`n=== VERIFY ===" -ForegroundColor Cyan
Add-Type -AssemblyName System.Drawing | Out-Null
$verifyFail = 0
foreach ($it in $plan) {
    $pptxPath = Join-Path $OutDir $it.PptxFileName
    $pngPath  = Join-Path $OutDir $it.ThumbFileName
    $okPptx = Test-Path $pptxPath
    $okPng  = Test-Path $pngPath
    $slides = 0
    if ($okPptx) { $p = $ppt.Presentations.Open($pptxPath, -1, 0, 0); $slides = $p.Slides.Count; $p.Close() }
    $longEdge = 0
    if ($okPng) { $img = [System.Drawing.Image]::FromFile($pngPath); $longEdge = [Math]::Max($img.Width, $img.Height); $img.Dispose() }
    $pass = $okPptx -and $okPng -and ($slides -eq 1) -and ($longEdge -eq 768)
    if ($pass) {
        Write-Host ("  OK   {0,-30} pptx:1slide png:{1}px" -f $it.AssetId, $longEdge) -ForegroundColor DarkGreen
    }
    else {
        $verifyFail++
        Write-Host ("  FAIL {0,-30} pptx:{1} png:{2} slides:{3} longEdge:{4}" -f $it.AssetId, $okPptx, $okPng, $slides, $longEdge) -ForegroundColor Red
    }
}

# Cleanup
$ppt.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)

Write-Host ("`nPlanned {0} items / Verify fail {1}" -f $total, $verifyFail)
if ($total -gt 0 -and $verifyFail -eq 0) {
    Write-Host "[PASS] All slides produced 1-slide pptx + 768px png pairs." -ForegroundColor Green
    exit 0
} else {
    Write-Host "[FAIL] See red lines above (or empty plan)." -ForegroundColor Red
    exit 1
}
