$ErrorActionPreference = "Stop"

$TaskName = "HMA Client Ads Export Every 12 Hours"
$ScriptPath = "D:\Proyectos\hma-system\scripts\export_all_clients.ps1"

$TaskRun = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""

schtasks /Create `
    /TN $TaskName `
    /TR $TaskRun `
    /SC HOURLY `
    /MO 12 `
    /ST 08:00 `
    /F | Out-Host

Write-Host "TASK_OK: $TaskName"

Get-ScheduledTask -TaskName $TaskName |
Select-Object TaskName, State |
Format-Table -AutoSize
