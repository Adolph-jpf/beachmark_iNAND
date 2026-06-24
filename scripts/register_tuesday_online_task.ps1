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

# Domain policy on many corporate PCs blocks "At logon" triggers (Access denied).
# Use a Tuesday schedule with 30-minute repetition instead; run_if_tuesday_online.ps1
# still checks day-of-week and skips after one success per ISO week.
$weeklyTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -At 6:00am
$weeklyTrigger.Repetition = New-CimInstance `
    -Namespace Root/Microsoft/Windows/TaskScheduler `
    -ClassName MSFT_TaskRepetitionPattern `
    -ClientOnly `
    -Property @{
        Interval          = 'PT30M'
        Duration          = 'P1D'
        StopAtDurationEnd = $false
    }

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $weeklyTrigger `
    -Settings $settings `
    -User $env:USERNAME `
    -Description "Run INAND benchmark once on Tuesday when the Spotfire export share is available." `
    -Force | Out-Null

Write-Host "Registered scheduled task: $TaskName"
Write-Host "ProjectDir: $ProjectDir"
Write-Host "DataPath  : $DataPath"
Write-Host "Log path  : $env:LOCALAPPDATA\INANDBenchmark\scheduled_run.log"
