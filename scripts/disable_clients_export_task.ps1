$TaskName = "HMA Client Ads Export Every 12 Hours"

Disable-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue | Out-Null

Write-Host "TASK_DISABLED: $TaskName"
Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue | Select-Object TaskName, State | Format-Table -AutoSize
