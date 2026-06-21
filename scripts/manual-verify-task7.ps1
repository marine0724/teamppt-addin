# Task 7 local ingest manual verification (throwaway script)
# - Runs the split pipeline in a clean PowerPoint COM instance (no source/add-in changes).
# - Prints the plan up front (N sections / N pptx) and shows per-item progress.
#
# Usage (no admin; close ALL PowerPoint first):
#   powershell -ExecutionPolicy Bypass -File .superpowers\sdd\manual-verify-task7.ps1 -Bundle "C:\Projects\teamppt-addin\assets\layout_test_aseet.pptx"

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

# Clean PowerPoint COM instance + load DLL, then call static methods directly.
$ppt = New-Object -ComObject PowerPoint.Application
[void][Reflection.Assembly]::LoadFrom($dll)
[TeampptAddin.Globals]::Application = $ppt

$source = $null
$fail = 0
$plan = $null
try {
    # Open source read-only, no window
    $source = $ppt.Presentations.Open($Bundle, -1, 0, 0)

    # 1) Read sections + build plan (show full plan first)
    $sections = [TeampptAddin.SectionReader]::Read($source)
    $plan     = [TeampptAddin.IngestPlanner]::Plan($sections)
    $total    = $plan.Count

    Write-Host "=== PLAN ===" -ForegroundColor Cyan
    Write-Host ("{0} sections -> {1} slides -> {1} pptx + {1} png ({2} files) expected`n" -f $sections.Count, $total, ($total * 2))
    foreach ($s in $sections) {
        Write-Host ("  - {0,-30} {1} slides" -f $s.Name, $s.SlideCount)
    }
    Write-Host ""

    # 2) Per-item split + render with progress
    $i = 0
    foreach ($it in $plan) {
        $i++
        $pct = [int](($i / $total) * 100)
        Write-Progress -Activity "Ingesting" -Status ("[{0}/{1}] {2}" -f $i, $total, $it.AssetId) -PercentComplete $pct
        Write-Host ("[{0,3}/{1}] {2,-30} (slide {3})" -f $i, $total, $it.AssetId, $it.SourceSlideIndex) -NoNewline

        $pptxPath = Join-Path $OutDir $it.PptxFileName
        $pngPath  = Join-Path $OutDir $it.ThumbFileName
        try {
            [TeampptAddin.SlideSplitter]::SplitSlide($ppt, $source, [int]$it.SourceSlideIndex, $pptxPath)
            [TeampptAddin.SlideImageRenderer]::Render($source, [int]$it.SourceSlideIndex, $pngPath, 768)
            Write-Host "  OK" -ForegroundColor DarkGreen
        }
        catch {
            $fail++
            $msg = $_.Exception.Message
            if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
            Write-Host ("  FAIL: {0}" -f $msg) -ForegroundColor Red
        }
    }
    Write-Progress -Activity "Ingesting" -Completed

    $source.Close()
}
finally {
    if ($source) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($source) }
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
    if (-not $pass) {
        $verifyFail++
        Write-Host ("  FAIL {0,-30} pptx:{1} png:{2} slides:{3} longEdge:{4}" -f $it.AssetId, $okPptx, $okPng, $slides, $longEdge) -ForegroundColor Red
    }
}

# Cleanup
$ppt.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)

$totalFail = $fail + $verifyFail
Write-Host ("`nGenerate fail {0} / Verify fail {1} / total {2} items" -f $fail, $verifyFail, $plan.Count)
if ($totalFail -eq 0) {
    Write-Host "[PASS] All slides produced 1-slide pptx + 768px png pairs." -ForegroundColor Green
    exit 0
} else {
    Write-Host "[FAIL] See red lines above." -ForegroundColor Red
    exit 1
}
