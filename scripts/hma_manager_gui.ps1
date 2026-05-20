Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"
$Master = Join-Path $BaseDir "historico\HMA_Master.xlsx"
$ClientesDir = Join-Path $BaseDir "clientes"
$HistoricoDir = Join-Path $BaseDir "historico"
$LogsDir = Join-Path $BaseDir "logs"

function Show-Info($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Manager", "OK", "Information") | Out-Null
}

function Show-ErrorBox($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Manager - Error", "OK", "Error") | Out-Null
}

function Confirm-Action($msg) {
    return ([System.Windows.Forms.MessageBox]::Show($msg, "Confirmar", "YesNo", "Warning") -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Run-File($path) {
    if (-not (Test-Path $path)) {
        Show-ErrorBox "No existe:`n$path"
        return
    }
    Start-Process -FilePath $path -WorkingDirectory $BaseDir
}

function Open-Folder($path) {
    if (-not (Test-Path $path)) {
        Show-ErrorBox "No existe:`n$path"
        return
    }
    Start-Process explorer.exe $path
}

function Show-TextWindow($titleText, $bodyText) {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $titleText
    $f.Size = New-Object System.Drawing.Size(1000,700)
    $f.StartPosition = "CenterScreen"

    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ScrollBars = "Both"
    $box.ReadOnly = $true
    $box.Font = New-Object System.Drawing.Font("Consolas", 10)
    $box.Dock = "Fill"
    $box.Text = $bodyText

    $f.Controls.Add($box)
    [void]$f.ShowDialog()
}

function Show-HmaStatus {
    try {
        $tasks = Get-ScheduledTask |
            Where-Object {
                $_.TaskName -eq "HMA Download Artifacts Every Minute" -or
                $_.TaskName -eq "HMA Error Monitor Every Minute" -or
                $_.TaskName -eq "HMA Promote Pending Master"
            } |
            Select-Object TaskName, State |
            Format-Table -AutoSize |
            Out-String

        $masterInfo = if (Test-Path $Master) {
            Get-Item $Master | Select-Object FullName, Length, LastWriteTime | Format-List | Out-String
        } else {
            "NO EXISTE HMA_Master.xlsx"
        }

        $excelCheck = if ((Test-Path $Python) -and (Test-Path $Master)) {
            & $Python -c "import sys,zipfile,openpyxl; p=sys.argv[1]; print('is_zipfile:',zipfile.is_zipfile(p)); wb=openpyxl.load_workbook(p,read_only=True,data_only=True); ws=wb['recommendations']; headers=[c.value for c in next(ws.iter_rows(min_row=1,max_row=1))]; print('has_resumen_horario:', 'resumen_horario' in headers); print('metric_crosses_state:', wb['metric_crosses'].sheet_state); print('recommendations_rows:', ws.max_row); wb.close()" $Master 2>&1 | Out-String
        } else {
            "No se pudo validar Excel. Falta Python o master."
        }

        $procs = Get-CimInstance Win32_Process |
            Where-Object { $_.CommandLine -match "download_latest_hma_artifact.ps1|update_hma_master.py|build_metric_crosses.py|rebuild_recommendations_hourly.py|hma_error_monitor.py|promote_hma_pending.py" } |
            Select-Object ProcessId, ParentProcessId, Name, CommandLine |
            Format-List |
            Out-String

        if ([string]::IsNullOrWhiteSpace($procs)) {
            $procs = "Sin procesos HMA activos."
        }

        $git = git -C $BaseDir status --short 2>&1 | Out-String
        if ([string]::IsNullOrWhiteSpace($git)) {
            $git = "Git limpio."
        }

        $msg = @"
=== HMA ACTUAL / GITHUB ARTIFACTS ===

Tareas:
$tasks

Master principal:
$masterInfo

Validacion Excel:
$excelCheck

Procesos:
$procs

Git:
$git
"@

        Show-TextWindow "Estado HMA actual" $msg
    } catch {
        Show-ErrorBox $_.Exception.Message
    }
}

function Show-HmaClients {
    try {
        if (-not (Test-Path $ClientesDir)) {
            Show-Info "No existe carpeta clientes."
            return
        }

        $rows = @()

        Get-ChildItem $ClientesDir -Directory |
        Where-Object { $_.Name -ne "_template" } |
        ForEach-Object {
            $configPath = Join-Path $_.FullName "config\client_config.json"

            if (Test-Path $configPath) {
                try {
                    $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
                    $google = if ($cfg.platforms.google_ads.enabled) { "conectado" } else { "no conectado" }
                    $meta = if ($cfg.platforms.meta_ads.enabled) { "conectado" } else { "no conectado" }

                    $rows += [PSCustomObject]@{
                        ID = $cfg.client_id
                        Cliente = $cfg.client_name
                        Carpeta = $_.Name
                        GoogleAds = $google
                        MetaAds = $meta
                        Ruta = $_.FullName
                    }
                } catch {
                    $rows += [PSCustomObject]@{
                        ID = "ERROR"
                        Cliente = "Config invalida"
                        Carpeta = $_.Name
                        GoogleAds = "-"
                        MetaAds = "-"
                        Ruta = $_.FullName
                    }
                }
            }
        }

        if ($rows.Count -eq 0) {
            Show-Info "No hay clientes creados."
            return
        }

        Show-TextWindow "Clientes HMA" ($rows | Format-Table -AutoSize | Out-String)
    } catch {
        Show-ErrorBox $_.Exception.Message
    }
}

function Show-GitStatus {
    $git = git -C $BaseDir status --short 2>&1 | Out-String
    if ([string]::IsNullOrWhiteSpace($git)) {
        $git = "Git limpio."
    }
    Show-TextWindow "Git Status" $git
}

function Show-CurrentHmaTasks {
    $txt = Get-ScheduledTask |
        Where-Object {
            $_.TaskName -eq "HMA Download Artifacts Every Minute" -or
            $_.TaskName -eq "HMA Error Monitor Every Minute" -or
            $_.TaskName -eq "HMA Promote Pending Master"
        } |
        Select-Object TaskName, State |
        Format-Table -AutoSize |
        Out-String

    Show-TextWindow "Tareas HMA actual / GitHub artifacts" $txt
}

function Enable-CurrentHmaTasks {
    if (-not (Confirm-Action "Reanudar HMA actual basado en GitHub artifacts?")) { return }

    Enable-ScheduledTask -TaskName "HMA Download Artifacts Every Minute" -ErrorAction SilentlyContinue | Out-Null
    Enable-ScheduledTask -TaskName "HMA Error Monitor Every Minute" -ErrorAction SilentlyContinue | Out-Null
    Enable-ScheduledTask -TaskName "HMA Promote Pending Master" -ErrorAction SilentlyContinue | Out-Null

    Show-Info "HMA actual reanudado."
}

function Disable-CurrentHmaTasks {
    if (-not (Confirm-Action "Pausar HMA actual basado en GitHub artifacts?")) { return }

    Disable-ScheduledTask -TaskName "HMA Download Artifacts Every Minute" -ErrorAction SilentlyContinue | Out-Null
    Disable-ScheduledTask -TaskName "HMA Error Monitor Every Minute" -ErrorAction SilentlyContinue | Out-Null
    Disable-ScheduledTask -TaskName "HMA Promote Pending Master" -ErrorAction SilentlyContinue | Out-Null

    Show-Info "HMA actual pausado."
}

function Run-ClientExportNow {
    Run-File (Join-Path $BaseDir "export_all_clients.bat")
}

function Setup-ClientExportTask {
    $script = Join-Path $BaseDir "scripts\setup_clients_export_task.ps1"
    if (-not (Test-Path $script)) {
        Show-ErrorBox "No existe:`n$script"
        return
    }

    $result = powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script 2>&1 | Out-String
    Show-TextWindow "Activar metricas cada 12 horas" $result
}

function Disable-ClientExportTask {
    $script = Join-Path $BaseDir "scripts\disable_clients_export_task.ps1"
    if (-not (Test-Path $script)) {
        Show-ErrorBox "No existe:`n$script"
        return
    }

    $result = powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script 2>&1 | Out-String
    Show-TextWindow "Pausar metricas cada 12 horas" $result
}

function Show-ClientExportTask {
    $txt = Get-ScheduledTask -TaskName "HMA Client Ads Export Every 12 Hours" -ErrorAction SilentlyContinue |
        Select-Object TaskName, State |
        Format-Table -AutoSize |
        Out-String

    if ([string]::IsNullOrWhiteSpace($txt)) {
        $txt = "No existe la tarea: HMA Client Ads Export Every 12 Hours"
    }

    Show-TextWindow "Automatizacion metricas clientes 12h" $txt
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "HMA Manager - Dashboard Maestro"
$form.Size = New-Object System.Drawing.Size(820,760)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$title = New-Object System.Windows.Forms.Label
$title.Text = "HMA Manager"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(25,20)
$title.Size = New-Object System.Drawing.Size(740,35)
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Dashboard maestro local: clientes, datos, HMA actual, automatizacion y estado."
$subtitle.Location = New-Object System.Drawing.Point(25,60)
$subtitle.Size = New-Object System.Drawing.Size(740,25)
$form.Controls.Add($subtitle)

function Add-SectionLabel($text, $x, $y) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lbl.Location = New-Object System.Drawing.Point($x,$y)
    $lbl.Size = New-Object System.Drawing.Size(340,24)
    $form.Controls.Add($lbl)
}

function Add-Button($text, $x, $y, $action) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Location = New-Object System.Drawing.Point($x,$y)
    $btn.Size = New-Object System.Drawing.Size(330,34)
    $btn.Add_Click($action)
    $form.Controls.Add($btn)
}

Add-SectionLabel "Clientes" 35 105
Add-Button "Crear cliente" 35 135 { Run-File (Join-Path $BaseDir "create_client.bat") }
Add-Button "Ver clientes creados" 35 178 { Show-HmaClients }
Add-Button "Abrir carpeta clientes" 35 221 { Open-Folder $ClientesDir }

Add-SectionLabel "Ads / Datos por cliente" 430 105
Add-Button "Conectar Google Ads / Meta Ads" 430 135 { Run-File (Join-Path $BaseDir "connect_ads.bat") }
Add-Button "Actualizar metricas manual" 430 178 { Run-File (Join-Path $BaseDir "export_ads.bat") }
Add-Button "Actualizar metricas todos ahora" 430 221 { Run-ClientExportNow }
Add-Button "Abrir carpeta clientes/raw_exports" 430 264 { Open-Folder $ClientesDir }

Add-SectionLabel "HMA actual / GitHub artifacts" 35 310
Add-Button "Ver estado HMA actual" 35 340 { Show-HmaStatus }
Add-Button "Ver tareas HMA actual" 35 383 { Show-CurrentHmaTasks }
Add-Button "Reanudar HMA actual" 35 426 { Enable-CurrentHmaTasks }
Add-Button "Pausar HMA actual" 35 469 { Disable-CurrentHmaTasks }
Add-Button "Abrir HMA_Master.xlsx" 35 512 { Run-File $Master }

Add-SectionLabel "Metricas automaticas clientes" 430 310
Add-Button "Ver automatizacion metricas 12h" 430 340 { Show-ClientExportTask }
Add-Button "Activar metricas cada 12h" 430 383 { Setup-ClientExportTask }
Add-Button "Pausar metricas cada 12h" 430 426 { Disable-ClientExportTask }
Add-Button "Abrir logs" 430 469 { Open-Folder $LogsDir }
Add-Button "Ver Git status" 430 512 { Show-GitStatus }

Add-SectionLabel "Archivos locales" 35 565
Add-Button "Abrir historico" 35 595 { Open-Folder $HistoricoDir }

Add-Button "Salir" 430 595 { $form.Close() }

[void]$form.ShowDialog()
