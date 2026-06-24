param(
    [ValidateSet("status", "run", "enable", "disable", "stop", "unregister", "logs")]
    [string]$Action = "status",
    [string]$TaskName = "INAND Benchmark Tuesday Online",
    [int]$Tail = 40
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$stateDir = Join-Path $env:LOCALAPPDATA "INANDBenchmark"
$scheduledLog = Join-Path $stateDir "scheduled_run.log"
$latestRunLog = Join-Path $stateDir "run_all_latest.log"

function Get-TaskOrNull {
    return Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
}

function Show-TaskStatus {
    $task = Get-TaskOrNull
    if ($null -eq $task) {
        Write-Host "Task not registered: $TaskName"
        return
    }

    $info = Get-ScheduledTaskInfo -TaskName $TaskName
    Write-Host "TaskName      : $TaskName"
    Write-Host "State         : $($task.State)"
    Write-Host "Enabled       : $($task.Settings.Enabled)"
    Write-Host "LastRunTime   : $($info.LastRunTime)"
    Write-Host "LastTaskResult: $($info.LastTaskResult)"
    Write-Host "NextRunTime   : $($info.NextRunTime)"
    Write-Host "LogDir        : $stateDir"
}

function Show-LogTail([string]$Path, [string]$Title) {
    Write-Host ""
    Write-Host "== $Title =="
    if (Test-Path $Path) {
        Get-Content -Path $Path -Tail $Tail
    }
    else {
        Write-Host "No log file: $Path"
    }
}

switch ($Action) {
    "status" {
        Show-TaskStatus
        Show-LogTail $scheduledLog "scheduled_run.log"
        break
    }
    "logs" {
        Show-LogTail $scheduledLog "scheduled_run.log"
        Show-LogTail $latestRunLog "run_all_latest.log"
        break
    }
    "run" {
        if ($null -eq (Get-TaskOrNull)) { throw "Task not registered: $TaskName" }
        Start-ScheduledTask -TaskName $TaskName
        Write-Host "Started task: $TaskName"
        break
    }
    "enable" {
        if ($null -eq (Get-TaskOrNull)) { throw "Task not registered: $TaskName" }
        Enable-ScheduledTask -TaskName $TaskName | Out-Null
        Write-Host "Enabled task: $TaskName"
        break
    }
    "disable" {
        if ($null -eq (Get-TaskOrNull)) { throw "Task not registered: $TaskName" }
        Disable-ScheduledTask -TaskName $TaskName | Out-Null
        Write-Host "Disabled task: $TaskName"
        break
    }
    "stop" {
        if ($null -eq (Get-TaskOrNull)) { throw "Task not registered: $TaskName" }
        Stop-ScheduledTask -TaskName $TaskName
        Write-Host "Stopped running task instance: $TaskName"
        break
    }
    "unregister" {
        if ($null -eq (Get-TaskOrNull)) {
            Write-Host "Task already absent: $TaskName"
            break
        }
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Unregistered task: $TaskName"
        break
    }
}
