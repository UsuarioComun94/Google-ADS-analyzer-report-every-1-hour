$TaskName = "HMA Weekly Local Backup"

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if (-not $task) {
    Write-Host "TASK_NOT_FOUND: $TaskName"
    exit 0
}

$info = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue

$task |
Select-Object TaskName, State |
Format-Table -AutoSize

if ($info) {
    $info |
    Select-Object LastRunTime, LastTaskResult, NextRunTime, NumberOfMissedRuns |
    Format-List
}
