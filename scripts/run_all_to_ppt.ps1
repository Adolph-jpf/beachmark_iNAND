param(
    [string]$DataPath = "\\cvpfilip03\SDSS_MFG_Data\ENG_Data\Tempfile\ADolph\Spotfire_file\PN_TEST_STEPID_YIELD.csv",
    [string]$RulePath = "Rule_list.xlsx",
    [string]$SrcGoalsPath = "bachmark SRC.xlsx",
    [string]$PptPath = "SDSS INAND YIELD WW45_2026_benchmark.pptx",
    [string]$OutputPptPath = "",
    [string]$PublicOutputDir = "\\cvpfilip03\SDSS_MFG_Data\ENG_Data\Tempfile\ADolph\Spotfire_file\output",
    [switch]$SkipPublicOutputSync,
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

$weekLine = $logs | Where-Object { $_ -match "\d{4}FW\d{1,2}" } | Select-Object -Last 1
if (-not $weekLine) {
    throw "Cannot find fiscal week in logs"
}
$weekMatches = [regex]::Matches($weekLine, "(\d{4})FW(\d{1,2})")
if ($weekMatches.Count -lt 2) {
    throw "Cannot find previous fiscal week in logs"
}
# Report filename follows the report period, so use prev_week (the second FW in the period log).
$weekMatch = $weekMatches[1]
$fiscalYear = $weekMatch.Groups[1].Value
$fiscalWeek = [int]$weekMatch.Groups[2].Value
$reportWeekLabel = "W{0:00}'{1}" -f $fiscalWeek, $fiscalYear.Substring(2, 2)
if ([string]::IsNullOrWhiteSpace($OutputPptPath)) {
    $OutputPptPath = Join-Path "output" ("SDSS INAND YIELD WW{0:00}_{1}_benchmark.pptx" -f $fiscalWeek, $fiscalYear)
}
Write-Host "Target PPT: $OutputPptPath"
Write-Host "Report week label: $reportWeekLabel"

# 3) Paste tables to PPT based on Rule_list third sheet (Sheet1)
& powershell -ExecutionPolicy Bypass -File "scripts/paste_excel_to_ppt.ps1" `
    -RulePath "$RulePath" `
    -ReportPath "$savedAbs" `
    -PptPath "$PptPath" `
    -OutputPptPath "$OutputPptPath" `
    -ReportWeekLabel "$reportWeekLabel" `
    -LeftInch $LeftInch `
    -TopInch $TopInch

if ($LASTEXITCODE -ne 0) {
    throw "paste_excel_to_ppt failed"
}

if (-not $SkipPublicOutputSync) {
    New-Item -ItemType Directory -Force -Path $PublicOutputDir | Out-Null
    Write-Host "Sync output folder to public path: $PublicOutputDir"
    & robocopy "output" "$PublicOutputDir" /E /R:1 /W:2 /NFL /NDL /NP
    $copyExit = $LASTEXITCODE
    if ($copyExit -ge 8) {
        throw "output folder sync failed with robocopy exit code $copyExit"
    }
}

Write-Host ""
Write-Host "DONE"
Write-Host "Excel: $savedAbs"
Write-Host "PPT  : $((Resolve-Path $OutputPptPath).Path)"
if (-not $SkipPublicOutputSync) {
    Write-Host "Public output: $((Resolve-Path $PublicOutputDir).Path)"
}
