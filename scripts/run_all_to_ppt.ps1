param(
    [string]$DataPath = "PN_TEST_STEPID_YIELD.txt",
    [string]$RulePath = "Rule_list.xlsx",
    [string]$SrcGoalsPath = "bachmark SRC.xlsx",
    [string]$PptPath = "SDSS INAND YIELD WW45_2026_benchmark.pptx",
    [double]$MaxAgeHours = 36.0,
    [int]$MinRows = 100,
    [double]$LeftInch = 0.2,
    [double]$TopInch = 1.5
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
if (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
    $PSNativeCommandUseErrorActionPreference = $false
}

Write-Host "== One-click benchmark -> PPT =="
$now = Get-Date
Write-Host ("Realtime now: {0}" -f $now.ToString("yyyy-MM-dd HH:mm:ss"))

$out = "output/INAND_weekly_benchmark_fmt_{0}.xlsx" -f $now.ToString("yyyyMMdd_HHmmss")
Write-Host "Target Excel: $out"

# 1) Validate Spotfire export contract before consuming
$checkLogs = @()
$prevEa = $ErrorActionPreference
$ErrorActionPreference = "Continue"
$checkLogs = & uv run python scripts/validate_spotfire_export.py `
    --data "$DataPath" `
    --max-age-hours "$MaxAgeHours" `
    --min-rows "$MinRows" 2>&1
$checkExit = $LASTEXITCODE
$ErrorActionPreference = $prevEa
$checkLogs | ForEach-Object { Write-Host $_ }
if ($checkExit -ne 0) {
    throw "spotfire export validation failed"
}

# 2) Build Excel with realtime date (do NOT pass --today)
$logs = @()
$prevEa = $ErrorActionPreference
$ErrorActionPreference = "Continue"
$logs = & uv run python main.py `
    --data "$DataPath" `
    --rules "$RulePath" `
    --src-goals "$SrcGoalsPath" `
    --output "$out" 2>&1
$nativeExit = $LASTEXITCODE
$ErrorActionPreference = $prevEa

$logs | ForEach-Object { Write-Host $_ }
if ($nativeExit -ne 0) {
    throw "benchmark generation failed"
}

# Handle save fallback (file locked) by parsing "Saved Excel:"
$savedLine = $logs | Where-Object { $_ -match "Saved Excel:\s*(.+)$" } | Select-Object -Last 1
if (-not $savedLine) {
    throw "Cannot find 'Saved Excel:' line in logs"
}
$savedRel = [regex]::Match($savedLine, "Saved Excel:\s*(.+)$").Groups[1].Value.Trim()
$savedAbs = (Resolve-Path $savedRel).Path
Write-Host "Excel saved: $savedAbs"

# 3) Paste tables to PPT based on Rule_list third sheet (Sheet1)
& powershell -ExecutionPolicy Bypass -File "scripts/paste_excel_to_ppt.ps1" `
    -RulePath "$RulePath" `
    -ReportPath "$savedAbs" `
    -PptPath "$PptPath" `
    -LeftInch $LeftInch `
    -TopInch $TopInch

if ($LASTEXITCODE -ne 0) {
    throw "paste_excel_to_ppt failed"
}

Write-Host ""
Write-Host "DONE"
Write-Host "Excel: $savedAbs"
Write-Host "PPT  : $((Resolve-Path $PptPath).Path)"
