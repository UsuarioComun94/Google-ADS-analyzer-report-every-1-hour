$TaskName = "HMA Weekly Health Check"

Disable-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue | Out-Null

Write-Host "TASK_DISABLED: $TaskName"

Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue |
Select-Object TaskName, State |
Format-Table -AutoSize
