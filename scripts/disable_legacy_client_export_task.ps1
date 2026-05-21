$TaskName = "HMA Client Ads Export Every 12 Hours"

Disable-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue | Out-Null

Write-Host "LEGACY_TASK_DISABLED: $TaskName"
Write-Host "Motivo: el ciclo completo cada 12h ya exporta metricas, reconstruye masters y ejecuta health check."

Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue |
Select-Object TaskName, State |
Format-Table -AutoSize
