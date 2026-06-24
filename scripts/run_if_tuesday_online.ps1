param(
    [string]$ProjectDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,
    [string]$DataPath = "\\cvpfilip03\SDSS_MFG_Data\ENG_Data\Tempfile\ADolph\Spotfire_file\PN_TEST_STEPID_YIELD.csv",
    [string]$TaskStampDir = "$env:LOCALAPPDATA\INANDBenchmark",
    [double]$MaxAgeHours = 36.0,
    [int]$MinRows = 100
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-IsoWeekKey {
    $now = Get-Date
    $calendar = [System.Globalization.CultureInfo]::InvariantCulture.Calendar
    $rule = [System.Globalization.CalendarWeekRule]::FirstFourDayWeek
    $week = $calendar.GetWeekOfYear($now, $rule, [System.DayOfWeek]::Monday)
    return "{0}-W{1:00}" -f $now.Year, $week
}

New-Item -ItemType Directory -Force -Path $TaskStampDir | Out-Null
$log = Join-Path $TaskStampDir "scheduled_run.log"
$stamp = Join-Path $TaskStampDir ("last_success_{0}.txt" -f (Get-IsoWeekKey))

function Write-RunLog([string]$Message) {
    $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $log -Value $line -Encoding UTF8
}

if ((Get-Date).DayOfWeek -ne [System.DayOfWeek]::Tuesday) {
    Write-RunLog "Skip: today is not Tuesday."
    exit 0
}

if (Test-Path $stamp) {
    Write-RunLog "Skip: already succeeded this week."
    exit 0
}

if (-not (Test-Path $DataPath)) {
    Write-RunLog "Skip: data path unavailable: $DataPath"
    exit 0
}

Push-Location $ProjectDir
try {
    Write-RunLog "Start weekly benchmark."
    & powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\run_all_to_ppt.ps1" `
        -DataPath $DataPath `
        -MaxAgeHours $MaxAgeHours `
        -MinRows $MinRows *> (Join-Path $TaskStampDir "run_all_latest.log")

    if ($LASTEXITCODE -ne 0) {
        throw "run_all_to_ppt.ps1 failed with exit code $LASTEXITCODE"
    }

    Set-Content -Path $stamp -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Encoding UTF8
    Write-RunLog "Success."
} catch {
    Write-RunLog ("Failed: {0}" -f $_.Exception.Message)
    throw
} finally {
    Pop-Location
}
