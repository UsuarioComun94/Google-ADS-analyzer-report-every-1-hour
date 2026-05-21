$tasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
Where-Object { $_.TaskName -like "HMA Informe *" }

if (!$tasks) {
    Write-Host "No hay tareas HMA Informe para pausar."
    exit 0
}

foreach ($task in $tasks) {
    Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "TASK_REMOVED: $($task.TaskName)"
}
