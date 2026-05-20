Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"
$Master = Join-Path $BaseDir "historico\HMA_Master.xlsx"

function Show-Info($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Manager", "OK", "Information") | Out-Null
}

function Show-ErrorBox($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Manager - Error", "OK", "Error") | Out-Null
}

function Run-File($path) {
    if (-not (Test-Path $path)) {
        Show-ErrorBox "No existe:`n$path"
        return
    }
    Start-Process -FilePath $path -WorkingDirectory $BaseDir
}

function Show-HmaStatus {
    try {
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

        $msg = @"
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

        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Estado HMA"
        $form.Size = New-Object System.Drawing.Size(900,650)
        $form.StartPosition = "CenterScreen"

        $box = New-Object System.Windows.Forms.TextBox
        $box.Multiline = $true
        $box.ScrollBars = "Vertical"
        $box.ReadOnly = $true
        $box.Font = New-Object System.Drawing.Font("Consolas", 10)
        $box.Dock = "Fill"
        $box.Text = $msg

        $form.Controls.Add($box)
        [void]$form.ShowDialog()
    } catch {
        Show-ErrorBox $_.Exception.Message
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "HMA Manager"
$form.Size = New-Object System.Drawing.Size(460,430)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$title = New-Object System.Windows.Forms.Label
$title.Text = "HMA Manager"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(25,20)
$title.Size = New-Object System.Drawing.Size(390,35)
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Panel local para manejar clientes, conexiones y estado del sistema."
$subtitle.Location = New-Object System.Drawing.Point(25,60)
$subtitle.Size = New-Object System.Drawing.Size(390,35)
$form.Controls.Add($subtitle)

function Add-Button($text, $y, $action) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Location = New-Object System.Drawing.Point(80,$y)
    $btn.Size = New-Object System.Drawing.Size(290,35)
    $btn.Add_Click($action)
    $form.Controls.Add($btn)
}

Add-Button "Crear cliente" 110 { Run-File (Join-Path $BaseDir "create_client.bat") }
Add-Button "Conectar Google Ads / Meta Ads" 155 { Run-File (Join-Path $BaseDir "connect_ads.bat") }
Add-Button "Abrir carpeta clientes" 200 { Start-Process explorer.exe (Join-Path $BaseDir "clientes") }
Add-Button "Ver estado HMA" 245 { Show-HmaStatus }
Add-Button "Salir" 290 { $form.Close() }

[void]$form.ShowDialog()

