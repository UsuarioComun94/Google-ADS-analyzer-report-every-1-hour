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

function Get-HmaStatusText {
    $tasks = Get-ScheduledTask |
        Where-Object { $_.TaskName -match "HMA" } |
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

    return @"
=== TAREAS HMA ===
$tasks

=== MASTER ===
$masterInfo

=== VALIDACION EXCEL ===
$excelCheck

=== PROCESOS HMA ===
$procs

=== GIT ===
$git
"@
}

function Show-HmaStatus {
    try {
        Show-TextWindow "Estado HMA" (Get-HmaStatusText)
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
                        Cliente = "Config inválida"
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

function Show-ScheduledTasks {
    $txt = Get-ScheduledTask |
        Where-Object { $_.TaskName -match "HMA" } |
        Select-Object TaskName, State |
        Format-Table -AutoSize |
        Out-String

    Show-TextWindow "Tareas HMA" $txt
}

function Enable-HmaTasks {
    if (-not (Confirm-Action "¿Activar todas las tareas HMA?")) { return }

    Enable-ScheduledTask -TaskName "HMA Download Artifacts Every Minute" -ErrorAction SilentlyContinue | Out-Null
    Enable-ScheduledTask -TaskName "HMA Error Monitor Every Minute" -ErrorAction SilentlyContinue | Out-Null
    Enable-ScheduledTask -TaskName "HMA Promote Pending Master" -ErrorAction SilentlyContinue | Out-Null

    Show-Info "Tareas HMA activadas."
}

function Disable-HmaTasks {
    if (-not (Confirm-Action "¿Pausar todas las tareas HMA?")) { return }

    Disable-ScheduledTask -TaskName "HMA Download Artifacts Every Minute" -ErrorAction SilentlyContinue | Out-Null
    Disable-ScheduledTask -TaskName "HMA Error Monitor Every Minute" -ErrorAction SilentlyContinue | Out-Null
    Disable-ScheduledTask -TaskName "HMA Promote Pending Master" -ErrorAction SilentlyContinue | Out-Null

    Show-Info "Tareas HMA pausadas."
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "HMA Manager - Dashboard Maestro"
$form.Size = New-Object System.Drawing.Size(760,620)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$title = New-Object System.Windows.Forms.Label
$title.Text = "HMA Manager"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(25,20)
$title.Size = New-Object System.Drawing.Size(700,35)
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Dashboard maestro local para clientes, conexiones, exports, tareas y estado del sistema."
$subtitle.Location = New-Object System.Drawing.Point(25,60)
$subtitle.Size = New-Object System.Drawing.Size(700,25)
$form.Controls.Add($subtitle)

function Add-SectionLabel($text, $x, $y) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lbl.Location = New-Object System.Drawing.Point($x,$y)
    $lbl.Size = New-Object System.Drawing.Size(300,24)
    $form.Controls.Add($lbl)
}

function Add-Button($text, $x, $y, $action) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Location = New-Object System.Drawing.Point($x,$y)
    $btn.Size = New-Object System.Drawing.Size(300,35)
    $btn.Add_Click($action)
    $form.Controls.Add($btn)
}

Add-SectionLabel "Clientes" 35 105
Add-Button "Crear cliente" 35 135 { Run-File (Join-Path $BaseDir "create_client.bat") }
Add-Button "Ver clientes creados" 35 180 { Show-HmaClients }
Add-Button "Abrir carpeta clientes" 35 225 { Open-Folder $ClientesDir }

Add-SectionLabel "Ads / Datos" 400 105
Add-Button "Conectar Google Ads / Meta Ads" 400 135 { Run-File (Join-Path $BaseDir "connect_ads.bat") }
Add-Button "Exportar Ads a CSV" 400 180 { Run-File (Join-Path $BaseDir "export_ads.bat") }
Add-Button "Abrir raw exports clientes" 400 225 { Open-Folder $ClientesDir }

Add-SectionLabel "Sistema" 35 285
Add-Button "Ver estado HMA" 35 315 { Show-HmaStatus }
Add-Button "Abrir HMA_Master.xlsx" 35 360 { Run-File $Master }
Add-Button "Abrir historico" 35 405 { Open-Folder $HistoricoDir }
Add-Button "Abrir logs" 35 450 { Open-Folder $LogsDir }

Add-SectionLabel "Operación" 400 285
Add-Button "Ver tareas programadas" 400 315 { Show-ScheduledTasks }
Add-Button "Activar tareas HMA" 400 360 { Enable-HmaTasks }
Add-Button "Pausar tareas HMA" 400 405 { Disable-HmaTasks }
Add-Button "Ver Git status" 400 450 { Show-GitStatus }

Add-Button "Salir" 220 525 { $form.Close() }

[void]$form.ShowDialog()
