Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Parent $PSScriptRoot
$ClientesDir = Join-Path $BaseDir "clientes"
$HistoricoDir = Join-Path $BaseDir "historico"
$LogsDir = Join-Path $BaseDir "logs"
$MasterFile = Join-Path $HistoricoDir "HMA_Master.xlsx"

function Msg($title, $text) {
    [System.Windows.Forms.MessageBox]::Show($text, $title, "OK", "Information") | Out-Null
}

function Err($text) {
    [System.Windows.Forms.MessageBox]::Show($text, "HMA Manager - Error", "OK", "Error") | Out-Null
}

function Open-Folder($path) {
    if (-not (Test-Path $path)) {
        Err "No existe:`n$path"
        return
    }
    Start-Process explorer.exe $path
}

function Open-File($path) {
    if (-not (Test-Path $path)) {
        Err "No existe:`n$path"
        return
    }
    Start-Process $path
}

function Run-Bat($name) {
    $path = Join-Path $BaseDir $name
    if (-not (Test-Path $path)) {
        Err "No existe:`n$path"
        return
    }
    Start-Process -FilePath $path -WorkingDirectory $BaseDir
}

function Run-PS1($relativePath) {
    $path = Join-Path $BaseDir $relativePath
    if (-not (Test-Path $path)) {
        Err "No existe:`n$path"
        return
    }
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$path`"" -WorkingDirectory $BaseDir
}

function TextWindow($title, $text) {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $title
    $f.Size = New-Object System.Drawing.Size(980,640)
    $f.StartPosition = "CenterScreen"

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.ScrollBars = "Both"
    $tb.WordWrap = $false
    $tb.ReadOnly = $true
    $tb.Dock = "Fill"
    $tb.Font = New-Object System.Drawing.Font("Consolas", 10)
    $tb.Text = $text

    $f.Controls.Add($tb)
    [void]$f.ShowDialog()
}

function Show-GitStatus {
    $txt = git -C $BaseDir status --short 2>&1 | Out-String
    if ([string]::IsNullOrWhiteSpace($txt)) {
        $txt = "Git limpio."
    }
    TextWindow "Git status" $txt
}

function Show-Clients {
    if (-not (Test-Path $ClientesDir)) {
        Msg "Clientes" "No existe carpeta clientes."
        return
    }

    $rows = @()

    Get-ChildItem $ClientesDir -Directory |
    Where-Object { $_.Name -ne "_template" } |
    Sort-Object Name |
    ForEach-Object {
        $cfgPath = Join-Path $_.FullName "config\client_config.json"

        if (Test-Path $cfgPath) {
            try {
                $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
                $rows += [PSCustomObject]@{
                    ID = $cfg.client_id
                    Cliente = $cfg.client_name
                    Carpeta = $_.Name
                    Ruta = $_.FullName
                }
            } catch {
                $rows += [PSCustomObject]@{
                    ID = "ERROR"
                    Cliente = "Config invalida"
                    Carpeta = $_.Name
                    Ruta = $_.FullName
                }
            }
        }
    }

    if ($rows.Count -eq 0) {
        Msg "Clientes" "No hay clientes creados."
        return
    }

    TextWindow "Clientes creados" ($rows | Format-Table -AutoSize | Out-String)
}

function Show-HmaTasks {
    $taskNames = @(
        "HMA Download Artifacts Every Minute",
        "HMA Error Monitor Every Minute",
        "HMA Promote Pending Master"
    )

    $rows = foreach ($name in $taskNames) {
        $task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue

        if ($task) {
            [PSCustomObject]@{
                TaskName = $task.TaskName
                State = $task.State
            }
        } else {
            [PSCustomObject]@{
                TaskName = $name
                State = "No existe"
            }
        }
    }

    TextWindow "Tareas HMA actual" ($rows | Format-Table -AutoSize | Out-String)
}

function Enable-HmaTasks {
    @(
        "HMA Download Artifacts Every Minute",
        "HMA Error Monitor Every Minute",
        "HMA Promote Pending Master"
    ) | ForEach-Object {
        Enable-ScheduledTask -TaskName $_ -ErrorAction SilentlyContinue | Out-Null
    }

    Msg "HMA actual" "Tareas HMA actual reanudadas."
}

function Disable-HmaTasks {
    @(
        "HMA Download Artifacts Every Minute",
        "HMA Error Monitor Every Minute",
        "HMA Promote Pending Master"
    ) | ForEach-Object {
        Disable-ScheduledTask -TaskName $_ -ErrorAction SilentlyContinue | Out-Null
    }

    Msg "HMA actual" "Tareas HMA actual pausadas."
}

function Show-ClientExportTask {
    $task = Get-ScheduledTask -TaskName "HMA Client Ads Export Every 12 Hours" -ErrorAction SilentlyContinue

    if (-not $task) {
        Msg "Metricas 12h" "No existe la tarea: HMA Client Ads Export Every 12 Hours"
        return
    }

    TextWindow "Automatizacion metricas 12h" ($task | Select-Object TaskName, State | Format-Table -AutoSize | Out-String)
}

function Show-ClientMetrics {
    if (-not (Test-Path $ClientesDir)) {
        Msg "Metricas clientes" "No existe carpeta clientes."
        return
    }

    $rows = @()

    Get-ChildItem $ClientesDir -Directory |
    Where-Object { $_.Name -ne "_template" } |
    Sort-Object Name |
    ForEach-Object {
        $clientPath = $_.FullName
        $cfgPath = Join-Path $clientPath "config\client_config.json"

        if (-not (Test-Path $cfgPath)) {
            return
        }

        try {
            $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json

            $googleDir = Join-Path $clientPath "raw_exports\google_ads"
            $metaDir = Join-Path $clientPath "raw_exports\meta_ads"

            $googleFiles = @()
            $metaFiles = @()

            if (Test-Path $googleDir) {
                $googleFiles = @(Get-ChildItem $googleDir -Filter "*.csv" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
            }

            if (Test-Path $metaDir) {
                $metaFiles = @(Get-ChildItem $metaDir -Filter "*.csv" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
            }

            $rows += [PSCustomObject]@{
                Cliente = $cfg.client_name
                ID = $cfg.client_id
                GoogleAds = if ($cfg.platforms.google_ads.enabled) { "conectado" } else { "no conectado" }
                MetaAds = if ($cfg.platforms.meta_ads.enabled) { "conectado" } else { "no conectado" }
                CSV_Google = $googleFiles.Count
                Ultimo_Google = if ($googleFiles.Count -gt 0) { $googleFiles[0].LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "sin CSV" }
                CSV_Meta = $metaFiles.Count
                Ultimo_Meta = if ($metaFiles.Count -gt 0) { $metaFiles[0].LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "sin CSV" }
            }
        } catch {
            $rows += [PSCustomObject]@{
                Cliente = $_.Name
                ID = "ERROR"
                GoogleAds = "-"
                MetaAds = "-"
                CSV_Google = "-"
                Ultimo_Google = "config invalida"
                CSV_Meta = "-"
                Ultimo_Meta = "config invalida"
            }
        }
    }

    if ($rows.Count -eq 0) {
        Msg "Metricas clientes" "No hay clientes para revisar."
        return
    }

    TextWindow "Estado metricas clientes" ($rows | Format-Table -AutoSize | Out-String)
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "HMA Manager - Dashboard V2"
$form.Size = New-Object System.Drawing.Size(1180,720)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$tree = New-Object System.Windows.Forms.TreeView
$tree.Location = New-Object System.Drawing.Point(10,55)
$tree.Size = New-Object System.Drawing.Size(285,610)
$tree.HideSelection = $false
$form.Controls.Add($tree)

$titleLeft = New-Object System.Windows.Forms.Label
$titleLeft.Text = "HMA Manager"
$titleLeft.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLeft.Location = New-Object System.Drawing.Point(10,10)
$titleLeft.Size = New-Object System.Drawing.Size(280,35)
$form.Controls.Add($titleLeft)

$right = New-Object System.Windows.Forms.Panel
$right.Location = New-Object System.Drawing.Point(310,10)
$right.Size = New-Object System.Drawing.Size(850,655)
$right.BorderStyle = "FixedSingle"
$form.Controls.Add($right)

function Add-Root($name, $children) {
    $root = New-Object System.Windows.Forms.TreeNode($name)

    foreach ($child in $children) {
        [void]$root.Nodes.Add((New-Object System.Windows.Forms.TreeNode($child)))
    }

    [void]$tree.Nodes.Add($root)
}

Add-Root "Google Ads" @("Manual", "Automatizacion 12h")
Add-Root "Meta Ads" @("Manual", "Automatizacion 12h")
Add-Root "Local" @("HMA actual", "Historico", "Logs")
Add-Root "Administrador" @("Clientes", "Carpetas", "Estado / Git")

$tree.ExpandAll()

function Clear-Panel {
    $right.Controls.Clear()
}

function Title($text, $y) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
    $lbl.Location = New-Object System.Drawing.Point(20,$y)
    $lbl.Size = New-Object System.Drawing.Size(780,35)
    $right.Controls.Add($lbl)
}

function Desc($text, $y) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Location = New-Object System.Drawing.Point(20,$y)
    $lbl.Size = New-Object System.Drawing.Size(780,45)
    $right.Controls.Add($lbl)
}

function Action($text, $y, $scriptBlock, $helpText) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Location = New-Object System.Drawing.Point(20,$y)
    $btn.Size = New-Object System.Drawing.Size(350,36)
    $btn.Add_Click($scriptBlock)
    $right.Controls.Add($btn)

    $help = New-Object System.Windows.Forms.Button
    $help.Text = "?"
    $help.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $help.Location = New-Object System.Drawing.Point(376,($y - 5))
    $help.Size = New-Object System.Drawing.Size(25,25)
    $help.Tag = [PSCustomObject]@{
        Title = "Ayuda - $text"
        Body = $helpText
    }
    $help.Add_Click({
        Msg $this.Tag.Title $this.Tag.Body
    })
    $right.Controls.Add($help)
}

function Render($path) {
    Clear-Panel

    switch ($path) {
        "Google Ads\Manual" {
            Title "Google Ads > Manual" 20
            Desc "Conexion, prueba y actualizacion manual de metricas." 60
            Action "Conectar / editar cuenta" 120 { Run-Bat "connect_ads.bat" } "Abre el asistente para elegir cliente y cargar credenciales. Guarda secretos dentro del cliente."
            Action "Actualizar metricas manual" 165 { Run-Bat "export_ads.bat" } "Permite elegir cliente, plataforma y periodo. Sirve como fallback manual."
            Action "Actualizar Google Ads ahora" 210 { Run-Bat "export_google_all_clients.bat" } "Exporta solo Google Ads de todos los clientes que tengan Google Ads conectado. Guarda CSV crudos."
            Action "Ver estado metricas clientes" 255 { Show-ClientMetrics } "Muestra conexion, cantidad de CSV y ultima exportacion por cliente."
        }

        "Google Ads\Automatizacion 12h" {
            Title "Google Ads > Automatizacion 12h" 20
            Desc "Control de la tarea automatica de metricas cada 12 horas." 60
            Action "Ver automatizacion metricas 12h" 120 { Show-ClientExportTask } "Muestra si la tarea automatica existe y si esta activa."
            Action "Activar metricas cada 12h" 165 { Run-PS1 "scripts\setup_clients_export_task.ps1" } "Crea o repara la tarea programada de metricas cada 12 horas."
            Action "Pausar metricas cada 12h" 210 { Run-PS1 "scripts\disable_clients_export_task.ps1" } "Desactiva la tarea automatica. No borra datos."
            Action "Abrir logs" 255 { Open-Folder $LogsDir } "Abre logs para revisar errores de exportacion."
        }

        "Meta Ads\Manual" {
            Title "Meta Ads > Manual" 20
            Desc "Conexion, prueba y actualizacion manual de metricas." 60
            Action "Conectar / editar cuenta" 120 { Run-Bat "connect_ads.bat" } "Abre el asistente para cargar token, ad account ID y version API."
            Action "Actualizar metricas manual" 165 { Run-Bat "export_ads.bat" } "Permite exportar metricas manualmente por cliente y periodo."
            Action "Actualizar Meta Ads ahora" 210 { Run-Bat "export_meta_all_clients.bat" } "Exporta solo Meta Ads de todos los clientes que tengan Meta Ads conectado. Guarda CSV crudos."
            Action "Ver estado metricas clientes" 255 { Show-ClientMetrics } "Muestra conexion, cantidad de CSV y ultima exportacion por cliente."
        }

        "Meta Ads\Automatizacion 12h" {
            Title "Meta Ads > Automatizacion 12h" 20
            Desc "Control de la tarea automatica de metricas cada 12 horas." 60
            Action "Ver automatizacion metricas 12h" 120 { Show-ClientExportTask } "Muestra estado de la tarea automatica."
            Action "Activar metricas cada 12h" 165 { Run-PS1 "scripts\setup_clients_export_task.ps1" } "Crea o repara la tarea programada."
            Action "Pausar metricas cada 12h" 210 { Run-PS1 "scripts\disable_clients_export_task.ps1" } "Pausa la automatizacion de metricas."
            Action "Abrir logs" 255 { Open-Folder $LogsDir } "Abre logs de ejecucion."
        }

        "Local\HMA actual" {
            Title "Local > HMA actual" 20
            Desc "Sistema local actual basado en GitHub artifacts." 60
            Action "Ver tareas HMA actual" 120 { Show-HmaTasks } "Muestra estado de tareas que descargan artifacts y actualizan el master principal."
            Action "Reanudar HMA actual" 165 { Enable-HmaTasks } "Activa tareas locales del HMA actual."
            Action "Pausar HMA actual" 210 { Disable-HmaTasks } "Pausa tareas locales del HMA actual."
            Action "Abrir HMA_Master.xlsx" 255 { Open-File $MasterFile } "Abre el Excel principal del HMA actual."
        }

        "Local\Historico" {
            Title "Local > Historico" 20
            Desc "Acceso a historico y Excel maestro." 60
            Action "Abrir historico" 120 { Open-Folder $HistoricoDir } "Abre carpeta historico."
            Action "Abrir HMA_Master.xlsx" 165 { Open-File $MasterFile } "Abre Excel maestro principal."
        }

        "Local\Logs" {
            Title "Local > Logs" 20
            Desc "Revision de logs locales." 60
            Action "Abrir logs" 120 { Open-Folder $LogsDir } "Abre carpeta logs."
            Action "Ver Git status" 165 { Show-GitStatus } "Muestra archivos modificados o pendientes."
        }

        "Administrador\Clientes" {
            Title "Administrador > Clientes" 20
            Desc "Gestion de clientes del sistema multi-cliente." 60
            Action "Crear cliente" 120 { Run-Bat "create_client.bat" } "Crea un nuevo cliente con ID automatico CL-XXX."
            Action "Ver clientes creados" 165 { Show-Clients } "Lista clientes existentes."
            Action "Abrir carpeta clientes" 210 { Open-Folder $ClientesDir } "Abre carpeta raiz de clientes."
        }

        "Administrador\Carpetas" {
            Title "Administrador > Carpetas" 20
            Desc "Accesos rapidos a carpetas principales." 60
            Action "Abrir clientes" 120 { Open-Folder $ClientesDir } "Abre carpeta clientes."
            Action "Abrir historico" 165 { Open-Folder $HistoricoDir } "Abre carpeta historico."
            Action "Abrir logs" 210 { Open-Folder $LogsDir } "Abre carpeta logs."
        }

        "Administrador\Estado / Git" {
            Title "Administrador > Estado / Git" 20
            Desc "Estado general del sistema local." 60
            Action "Ver Git status" 120 { Show-GitStatus } "Muestra cambios pendientes del repositorio."
            Action "Ver estado metricas clientes" 165 { Show-ClientMetrics } "Muestra estado de exports por cliente."
            Action "Ver automatizacion metricas 12h" 210 { Show-ClientExportTask } "Muestra estado de la tarea automatica de metricas."
            Action "Actualizar todas las plataformas ahora" 255 { Run-Bat "export_all_clients.bat" } "Exporta Google Ads y Meta Ads de todos los clientes conectados. Es la actualizacion global."
        }

        default {
            Title "HMA Manager" 20
            Desc "Selecciona una subcategoria del panel izquierdo." 60
        }
    }
}

$tree.Add_AfterSelect({
    $n = $tree.SelectedNode

    if ($n -and $n.Parent) {
        Render ($n.Parent.Text + "\" + $n.Text)
    } else {
        Clear-Panel
        Title "HMA Manager" 20
        Desc "Selecciona una subcategoria del panel izquierdo." 60
    }
})

Render ""
[void]$form.ShowDialog()
