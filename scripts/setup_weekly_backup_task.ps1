$ErrorActionPreference = "Stop"

$TaskName = "HMA Weekly Local Backup"
$ScriptPath = "D:\Proyectos\hma-system\scripts\backup_hma_local.ps1"

$TaskRun = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -Silent"

schtasks /Create `
    /TN $TaskName `
    /TR $TaskRun `
    /SC WEEKLY `
    /D MON `
    /ST 09:00 `
    /F | Out-Host

Write-Host "TASK_OK: $TaskName"

Get-ScheduledTask -TaskName $TaskName |
Select-Object TaskName, State |
Format-Table -AutoSize
