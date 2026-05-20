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

    $shell = New-Object -ComObject WScript.Shell
    $cmd = 'cmd.exe /c ""' + $path + '""'
    $shell.Run($cmd, 0, $false) | Out-Null
}

function Run-PS1($relativePath) {
    $path = Join-Path $BaseDir $relativePath

    if (-not (Test-Path $path)) {
        Err "No existe:`n$path"
        return
    }

    $shell = New-Object -ComObject WScript.Shell
    $cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $path + '"'
    $shell.Run($cmd, 0, $false) | Out-Null
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


function Add-RootNode($name) {
    $root = New-Object System.Windows.Forms.TreeNode($name)
    [void]$tree.Nodes.Add($root)
    return $root
}

function Add-ChildNode($parent, $name) {
    $child = New-Object System.Windows.Forms.TreeNode($name)
    [void]$parent.Nodes.Add($child)
    return $child
}

$googleRoot = Add-RootNode "Google Ads"
$googleManual = Add-ChildNode $googleRoot "Manual"
[void](Add-ChildNode $googleManual "Conectar / editar cuenta")
[void](Add-ChildNode $googleManual "Actualizar metricas manual")
[void](Add-ChildNode $googleManual "Actualizar Google Ads ahora")
[void](Add-ChildNode $googleManual "Ver estado metricas clientes")
$googleAuto = Add-ChildNode $googleRoot "Automatizacion 12h"
[void](Add-ChildNode $googleAuto "Ver automatizacion metricas 12h")
[void](Add-ChildNode $googleAuto "Activar metricas cada 12h")
[void](Add-ChildNode $googleAuto "Pausar metricas cada 12h")
[void](Add-ChildNode $googleAuto "Abrir logs")

$metaRoot = Add-RootNode "Meta Ads"
$metaManual = Add-ChildNode $metaRoot "Manual"
[void](Add-ChildNode $metaManual "Conectar / editar cuenta")
[void](Add-ChildNode $metaManual "Actualizar metricas manual")
[void](Add-ChildNode $metaManual "Actualizar Meta Ads ahora")
[void](Add-ChildNode $metaManual "Ver estado metricas clientes")
$metaAuto = Add-ChildNode $metaRoot "Automatizacion 12h"
[void](Add-ChildNode $metaAuto "Ver automatizacion metricas 12h")
[void](Add-ChildNode $metaAuto "Activar metricas cada 12h")
[void](Add-ChildNode $metaAuto "Pausar metricas cada 12h")
[void](Add-ChildNode $metaAuto "Abrir logs")

$localRoot = Add-RootNode "Local"
$localHma = Add-ChildNode $localRoot "HMA actual"
[void](Add-ChildNode $localHma "Ver tareas HMA actual")
[void](Add-ChildNode $localHma "Reanudar HMA actual")
[void](Add-ChildNode $localHma "Pausar HMA actual")
[void](Add-ChildNode $localHma "Abrir HMA_Master.xlsx")
$localMasters = Add-ChildNode $localRoot "Masters clientes"
[void](Add-ChildNode $localMasters "Construir masters de todos los clientes")
[void](Add-ChildNode $localMasters "Ver estado metricas clientes")
[void](Add-ChildNode $localMasters "Abrir carpeta clientes")
[void](Add-ChildNode $localMasters "Abrir logs")
$localHistorico = Add-ChildNode $localRoot "Historico"
[void](Add-ChildNode $localHistorico "Abrir historico")
[void](Add-ChildNode $localHistorico "Abrir HMA_Master.xlsx")
$localLogs = Add-ChildNode $localRoot "Logs"
[void](Add-ChildNode $localLogs "Abrir logs")
[void](Add-ChildNode $localLogs "Ver Git status")

$adminRoot = Add-RootNode "Administrador"
$adminClientes = Add-ChildNode $adminRoot "Clientes"
[void](Add-ChildNode $adminClientes "Crear cliente")
[void](Add-ChildNode $adminClientes "Ver clientes creados")
[void](Add-ChildNode $adminClientes "Abrir carpeta clientes")
$adminCarpetas = Add-ChildNode $adminRoot "Carpetas"
[void](Add-ChildNode $adminCarpetas "Abrir clientes")
[void](Add-ChildNode $adminCarpetas "Abrir historico")
[void](Add-ChildNode $adminCarpetas "Abrir logs")
$adminEstado = Add-ChildNode $adminRoot "Estado / Git"
[void](Add-ChildNode $adminEstado "Ver Git status")
[void](Add-ChildNode $adminEstado "Ver estado metricas clientes")
[void](Add-ChildNode $adminEstado "Ver automatizacion metricas 12h")
[void](Add-ChildNode $adminEstado "Actualizar todas las plataformas ahora")
[void](Add-ChildNode $adminEstado "Health check sistema")

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

        "Google Ads" {
            Title "Google Ads" 20
            Desc "Selecciona Manual para acciones puntuales o Automatizacion 12h para tareas automaticas." 60
            Action "Ir a Google Ads > Manual" 120 { Render "Google Ads\Manual" } "Muestra acciones manuales de Google Ads."
            Action "Ir a Google Ads > Automatizacion 12h" 165 { Render "Google Ads\Automatizacion 12h" } "Muestra controles de automatizacion cada 12 horas."
        }

        "Meta Ads" {
            Title "Meta Ads" 20
            Desc "Selecciona Manual para acciones puntuales o Automatizacion 12h para tareas automaticas." 60
            Action "Ir a Meta Ads > Manual" 120 { Render "Meta Ads\Manual" } "Muestra acciones manuales de Meta Ads."
            Action "Ir a Meta Ads > Automatizacion 12h" 165 { Render "Meta Ads\Automatizacion 12h" } "Muestra controles de automatizacion cada 12 horas."
        }

        "Local" {
            Title "Local" 20
            Desc "Acciones locales del sistema: HMA actual, masters de clientes, historico y logs." 60
            Action "Ir a Local > HMA actual" 120 { Render "Local\HMA actual" } "Muestra controles del HMA actual basado en GitHub artifacts."
            Action "Ir a Local > Masters clientes" 165 { Render "Local\Masters clientes" } "Muestra herramientas para construir masters por cliente."
            Action "Ir a Local > Historico" 210 { Render "Local\Historico" } "Muestra accesos al historico."
            Action "Ir a Local > Logs" 255 { Render "Local\Logs" } "Muestra accesos a logs."
        }

        "Administrador" {
            Title "Administrador" 20
            Desc "Gestion general del sistema: clientes, carpetas y estado." 60
            Action "Ir a Administrador > Clientes" 120 { Render "Administrador\Clientes" } "Muestra opciones para crear y listar clientes."
            Action "Ir a Administrador > Carpetas" 165 { Render "Administrador\Carpetas" } "Muestra accesos rapidos a carpetas."
            Action "Ir a Administrador > Estado / Git" 210 { Render "Administrador\Estado / Git" } "Muestra estado general y Git."
        }


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


        "Local\Masters clientes" {
            Title "Local > Masters clientes" 20
            Desc "Construccion de HMA_Master.xlsx por cliente usando los CSV crudos de raw_exports." 60
            Action "Construir masters de todos los clientes" 120 { Run-Bat "build_all_client_masters.bat" } "Genera o actualiza clientes\\CL-XXX\\historico\\HMA_Master.xlsx usando raw_exports de Google Ads y Meta Ads. No modifica el HMA_Master principal."
            Action "Ver estado metricas clientes" 165 { Show-ClientMetrics } "Muestra por cliente si hay CSV de Google Ads o Meta Ads y la ultima fecha de exportacion."
            Action "Abrir carpeta clientes" 210 { Open-Folder $ClientesDir } "Abre la carpeta donde estan los clientes y sus masters individuales."
            Action "Abrir logs" 255 { Open-Folder $LogsDir } "Abre los logs para revisar ejecuciones o errores del builder."
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
            Action "Health check sistema" 300 { Run-Bat "hma_health_check.bat" } "Ejecuta un diagnostico local: tareas, clientes, masters, CSV, logs y Git status."
        }

        default {
            Title "HMA Manager" 20
            Desc "Selecciona una subcategoria del panel izquierdo." 60
        }
    }
}


function Get-TreePath($node) {
    $parts = New-Object System.Collections.ArrayList

    while ($node -ne $null) {
        [void]$parts.Insert(0, $node.Text)
        $node = $node.Parent
    }

    return ($parts -join "\")
}

function Get-RenderablePath($path) {
    $parts = $path -split "\\\\"

    if ($parts.Count -le 2) {
        return $path
    }

    return ($parts[0..1] -join "\")
}


function Get-NodeDepth($node) {
    $depth = 0

    while ($node -ne $null) {
        $depth += 1
        $node = $node.Parent
    }

    return $depth
}

function Get-RenderPath($node) {
    if ($node -eq $null) {
        return ""
    }

    $depth = Get-NodeDepth $node

    if ($depth -eq 1) {
        return $node.Text
    }

    if ($depth -eq 2) {
        return "$($node.Parent.Text)\$($node.Text)"
    }

    if ($depth -ge 3) {
        return "$($node.Parent.Parent.Text)\$($node.Parent.Text)"
    }

    return ""
}

function Get-ActionPath($node) {
    if ($node -eq $null) {
        return ""
    }

    $depth = Get-NodeDepth $node

    if ($depth -ge 3) {
        return "$($node.Parent.Parent.Text)\$($node.Parent.Text)\$($node.Text)"
    }

    return ""
}

function Render-SelectedNode {
    $node = $tree.SelectedNode

    if ($node -eq $null) {
        Clear-Panel
        Title "HMA Manager" 20
        Desc "Selecciona una categoria del panel izquierdo." 60
        return
    }

    $renderPath = Get-RenderPath $node

    if ([string]::IsNullOrWhiteSpace($renderPath)) {
        Clear-Panel
        Title "HMA Manager" 20
        Desc "Selecciona una categoria del panel izquierdo." 60
        return
    }

    Render $renderPath
}

function Invoke-ActionPath($actionPath) {
    switch ($actionPath) {
        "Google Ads\Manual\Conectar / editar cuenta" { Run-Bat "connect_ads.bat" }
        "Google Ads\Manual\Actualizar metricas manual" { Run-Bat "export_ads.bat" }
        "Google Ads\Manual\Actualizar Google Ads ahora" { Run-Bat "export_google_all_clients.bat" }
        "Google Ads\Manual\Ver estado metricas clientes" { Show-ClientMetrics }

        "Google Ads\Automatizacion 12h\Ver automatizacion metricas 12h" { Show-ClientExportTask }
        "Google Ads\Automatizacion 12h\Activar metricas cada 12h" { Run-PS1 "scripts\setup_clients_export_task.ps1" }
        "Google Ads\Automatizacion 12h\Pausar metricas cada 12h" { Run-PS1 "scripts\disable_clients_export_task.ps1" }
        "Google Ads\Automatizacion 12h\Abrir logs" { Open-Folder $LogsDir }

        "Meta Ads\Manual\Conectar / editar cuenta" { Run-Bat "connect_ads.bat" }
        "Meta Ads\Manual\Actualizar metricas manual" { Run-Bat "export_ads.bat" }
        "Meta Ads\Manual\Actualizar Meta Ads ahora" { Run-Bat "export_meta_all_clients.bat" }
        "Meta Ads\Manual\Ver estado metricas clientes" { Show-ClientMetrics }

        "Meta Ads\Automatizacion 12h\Ver automatizacion metricas 12h" { Show-ClientExportTask }
        "Meta Ads\Automatizacion 12h\Activar metricas cada 12h" { Run-PS1 "scripts\setup_clients_export_task.ps1" }
        "Meta Ads\Automatizacion 12h\Pausar metricas cada 12h" { Run-PS1 "scripts\disable_clients_export_task.ps1" }
        "Meta Ads\Automatizacion 12h\Abrir logs" { Open-Folder $LogsDir }

        "Local\HMA actual\Ver tareas HMA actual" { Show-HmaTasks }
        "Local\HMA actual\Reanudar HMA actual" { Enable-HmaTasks }
        "Local\HMA actual\Pausar HMA actual" { Disable-HmaTasks }
        "Local\HMA actual\Abrir HMA_Master.xlsx" { Open-File $MasterFile }

        "Local\Masters clientes\Construir masters de todos los clientes" { Run-Bat "build_all_client_masters.bat" }
        "Local\Masters clientes\Ver estado metricas clientes" { Show-ClientMetrics }
        "Local\Masters clientes\Abrir carpeta clientes" { Open-Folder $ClientesDir }
        "Local\Masters clientes\Abrir logs" { Open-Folder $LogsDir }

        "Local\Historico\Abrir historico" { Open-Folder $HistoricoDir }
        "Local\Historico\Abrir HMA_Master.xlsx" { Open-File $MasterFile }

        "Local\Logs\Abrir logs" { Open-Folder $LogsDir }
        "Local\Logs\Ver Git status" { Show-GitStatus }

        "Administrador\Clientes\Crear cliente" { Run-Bat "create_client.bat" }
        "Administrador\Clientes\Ver clientes creados" { Show-Clients }
        "Administrador\Clientes\Abrir carpeta clientes" { Open-Folder $ClientesDir }

        "Administrador\Carpetas\Abrir clientes" { Open-Folder $ClientesDir }
        "Administrador\Carpetas\Abrir historico" { Open-Folder $HistoricoDir }
        "Administrador\Carpetas\Abrir logs" { Open-Folder $LogsDir }

        "Administrador\Estado / Git\Ver Git status" { Show-GitStatus }
        "Administrador\Estado / Git\Ver estado metricas clientes" { Show-ClientMetrics }
        "Administrador\Estado / Git\Ver automatizacion metricas 12h" { Show-ClientExportTask }
        "Administrador\Estado / Git\Actualizar todas las plataformas ahora" { Run-Bat "export_all_clients.bat" }
        "Administrador\Estado / Git\Health check sistema" { Run-Bat "hma_health_check.bat" }

        default { Render-SelectedNode }
    }
}

function Invoke-SelectedNodeAction {
    $node = $tree.SelectedNode

    if ($node -eq $null) {
        return
    }

    $depth = Get-NodeDepth $node

    if ($depth -ge 3) {
        Invoke-ActionPath (Get-ActionPath $node)
    } else {
        Render-SelectedNode
    }
}

$tree.Add_AfterSelect({
    Render-SelectedNode
})

$tree.Add_NodeMouseDoubleClick({
    param($sender, $e)

    if ($e.Node -ne $null) {
        $tree.SelectedNode = $e.Node
        Invoke-SelectedNodeAction
    }
})

$tree.Add_KeyDown({
    param($sender, $e)

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Invoke-SelectedNodeAction
    }
})

Render ""
[void]$form.ShowDialog()
