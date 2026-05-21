param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("1h","3h","5h","7h","12h","1d","2d","1w")]
    [string]$Frequency
)

$ErrorActionPreference = "Stop"

$BaseDir = "D:\Proyectos\hma-system"
$script = Join-Path $BaseDir "scripts\hma_report_engine.ps1"

if (!(Test-Path $script)) {
    throw "No existe hma_report_engine.ps1"
}

$oldTasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
Where-Object { $_.TaskName -like "HMA Informe *" }

foreach ($task in $oldTasks) {
    Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
}

$taskName = "HMA Informe $Frequency"
$startAt = (Get-Date).Date.AddMinutes(10)

switch ($Frequency) {
    "1h"  { $trigger = New-ScheduledTaskTrigger -Once -At $startAt -RepetitionInterval (New-TimeSpan -Hours 1)  -RepetitionDuration (New-TimeSpan -Days 3650) }
    "3h"  { $trigger = New-ScheduledTaskTrigger -Once -At $startAt -RepetitionInterval (New-TimeSpan -Hours 3)  -RepetitionDuration (New-TimeSpan -Days 3650) }
    "5h"  { $trigger = New-ScheduledTaskTrigger -Once -At $startAt -RepetitionInterval (New-TimeSpan -Hours 5)  -RepetitionDuration (New-TimeSpan -Days 3650) }
    "7h"  { $trigger = New-ScheduledTaskTrigger -Once -At $startAt -RepetitionInterval (New-TimeSpan -Hours 7)  -RepetitionDuration (New-TimeSpan -Days 3650) }
    "12h" { $trigger = New-ScheduledTaskTrigger -Once -At $startAt -RepetitionInterval (New-TimeSpan -Hours 12) -RepetitionDuration (New-TimeSpan -Days 3650) }
    "1d"  { $trigger = New-ScheduledTaskTrigger -Daily -DaysInterval 1 -At 9:00am }
    "2d"  { $trigger = New-ScheduledTaskTrigger -Daily -DaysInterval 2 -At 9:00am }
    "1w"  { $trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Monday -At 9:00am }
}

$actionArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -Frequency $Frequency"
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $actionArgs -WorkingDirectory $BaseDir
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description "HMA configurable report automation: $Frequency" -Force | Out-Null

Disable-ScheduledTask -TaskName "HMA Full Cycle Every 12 Hours" -ErrorAction SilentlyContinue | Out-Null

Write-Host "TASK_OK: $taskName"
Write-Host "Frecuencia activa: $Frequency"
Write-Host "Script: $script"
Write-Host "Nota: se deshabilito el task legacy HMA Full Cycle Every 12 Hours si existia."
Get-ScheduledTask -TaskName $taskName | Select-Object TaskName, State | Format-Table -AutoSize
