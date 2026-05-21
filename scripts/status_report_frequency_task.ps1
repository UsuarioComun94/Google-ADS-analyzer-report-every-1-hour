$tasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
Where-Object { $_.TaskName -like "HMA Informe *" }

if (!$tasks) {
    Write-Host "No hay automatizacion de informes activa."
    exit 0
}

foreach ($task in $tasks) {
    $info = Get-ScheduledTaskInfo -TaskName $task.TaskName -ErrorAction SilentlyContinue
    Write-Host "============================================================"
    Write-Host "TaskName: $($task.TaskName)"
    Write-Host "State: $($task.State)"
    if ($info) {
        Write-Host "LastRunTime: $($info.LastRunTime)"
        Write-Host "LastTaskResult: $($info.LastTaskResult)"
        Write-Host "NextRunTime: $($info.NextRunTime)"
    }
}
