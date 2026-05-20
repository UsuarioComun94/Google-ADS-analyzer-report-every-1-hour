$ErrorActionPreference = "Stop"

$TaskName = "HMA Full Cycle Every 12 Hours"
$ScriptPath = "D:\Proyectos\hma-system\scripts\hma_run_full_cycle.ps1"
$TaskRun = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -Silent"

schtasks /Create `
    /TN $TaskName `
    /TR $TaskRun `
    /SC HOURLY `
    /MO 12 `
    /ST 00:10 `
    /F | Out-Host

Write-Host "TASK_OK: $TaskName"

Get-ScheduledTask -TaskName $TaskName |
Select-Object TaskName, State |
Format-Table -AutoSize
