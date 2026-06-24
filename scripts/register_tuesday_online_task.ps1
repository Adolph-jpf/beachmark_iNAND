param(
    [string]$TaskName = "INAND Benchmark Tuesday Online",
    [string]$ProjectDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,
    [string]$DataPath = "\\cvpfilip03\SDSS_MFG_Data\ENG_Data\Tempfile\ADolph\Spotfire_file\PN_TEST_STEPID_YIELD.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$runner = Join-Path $ProjectDir "scripts\run_if_tuesday_online.ps1"
if (-not (Test-Path $runner)) {
    throw "Runner script not found: $runner"
}

$args = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$runner`"",
    "-ProjectDir", "`"$ProjectDir`"",
    "-DataPath", "`"$DataPath`""
) -join " "

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $args -WorkingDirectory $ProjectDir

# Run at user logon and keep trying every 30 minutes on Tuesday.
# The runner script is idempotent: after one success in the ISO week it exits.
$logonTrigger = New-ScheduledTaskTrigger -AtLogOn
$weeklyTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -At 8:00am
$weeklyTrigger.Repetition.Interval = "PT30M"
$weeklyTrigger.Repetition.Duration = "P1D"

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger @($logonTrigger, $weeklyTrigger) `
    -Settings $settings `
    -Description "Run INAND benchmark once on Tuesday when the Spotfire export share is available." `
    -Force | Out-Null

Write-Host "Registered scheduled task: $TaskName"
Write-Host "ProjectDir: $ProjectDir"
Write-Host "DataPath  : $DataPath"
Write-Host "Log path  : $env:LOCALAPPDATA\INANDBenchmark\scheduled_run.log"
