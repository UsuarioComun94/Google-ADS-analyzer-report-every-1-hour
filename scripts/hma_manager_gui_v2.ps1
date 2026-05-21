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

$localBackups = Add-ChildNode $localRoot "Backups"
[void](Add-ChildNode $localBackups "Crear backup local ahora")
[void](Add-ChildNode $localBackups "Ver backup semanal")
[void](Add-ChildNode $localBackups "Programar backup semanal")
[void](Add-ChildNode $localBackups "Pausar backup semanal")
[void](Add-ChildNode $localBackups "Abrir carpeta backups")
[void](Add-ChildNode $localBackups "Limpiar backups antiguos")
[void](Add-ChildNode $localBackups "Restaurar backup local")
[void](Add-ChildNode $localBackups "Crear paquete portable")


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
[void](Add-ChildNode $adminEstado "Ejecutar ciclo completo ahora")
[void](Add-ChildNode $adminEstado "Ver ciclo completo 12h")
[void](Add-ChildNode $adminEstado "Ver todas las automatizaciones")
[void](Add-ChildNode $adminEstado "Validacion QA completa")
[void](Add-ChildNode $adminEstado "Pausar tarea legacy export")
[void](Add-ChildNode $adminEstado "Activar ciclo completo 12h")
[void](Add-ChildNode $adminEstado "Pausar ciclo completo 12h")
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


function Get-DetailedHelp($action, $defaultHelp) {
    switch ($action) {
        "Conectar / editar cuenta" {
            return @"
Sirve para vincular una cuenta publicitaria a un cliente del sistema.

Primero elegis el cliente. Despues elegis la plataforma: Google Ads o Meta Ads.
Luego el sistema te pide las credenciales necesarias y las guarda dentro de la carpeta de ese cliente.

Usalo cuando:
- creaste un cliente nuevo
- queres conectar Google Ads o Meta Ads
- queres corregir credenciales
- queres cambiar el Customer ID, Ad Account ID, token o datos de OAuth

Resultado esperado:
el cliente queda marcado como conectado para esa plataforma.
"@
        }

        "Actualizar metricas manual" {
            return @"
Ejecuta una extraccion manual de metricas para un cliente y una plataforma.

Sirve como plan B cuando no queres esperar la automatizacion de 12 horas o cuando queres probar si una conexion funciona.

Usalo cuando:
- acabas de conectar una cuenta
- queres probar si las credenciales funcionan
- queres traer datos ahora mismo
- la tarea automatica fallo o todavia no corrio

Resultado esperado:
se crean CSV crudos dentro de la carpeta raw_exports del cliente.
"@
        }

        "Actualizar Google Ads ahora" {
            return @"
Ejecuta ahora la extraccion de metricas SOLO de Google Ads para todos los clientes que tengan Google Ads conectado.

No toca Meta Ads.
No modifica directamente el HMA_Master principal.
Genera o actualiza archivos CSV crudos de Google Ads por cliente.

Usalo cuando:
- queres actualizar Google Ads sin tocar Meta
- queres probar las conexiones de Google Ads
- queres traer datos antes del ciclo automatico de 12 horas

Resultado esperado:
cada cliente conectado a Google Ads recibe nuevos datos en raw_exports\google_ads.
"@
        }

        "Actualizar Meta Ads ahora" {
            return @"
Ejecuta ahora la extraccion de metricas SOLO de Meta Ads para todos los clientes que tengan Meta Ads conectado.

No toca Google Ads.
No modifica directamente el HMA_Master principal.
Genera o actualiza archivos CSV crudos de Meta Ads por cliente.

Usalo cuando:
- queres actualizar Meta Ads sin tocar Google
- queres probar tokens o Ad Account IDs
- queres traer datos antes del ciclo automatico de 12 horas

Resultado esperado:
cada cliente conectado a Meta Ads recibe nuevos datos en raw_exports\meta_ads.
"@
        }

        "Actualizar todas las plataformas ahora" {
            return @"
Ejecuta una actualizacion global de plataformas.

Intenta exportar Google Ads y Meta Ads de todos los clientes que tengan alguna cuenta conectada.

Usalo cuando:
- queres traer todos los datos disponibles
- queres forzar una actualizacion general
- queres verificar si todas las conexiones siguen funcionando

Resultado esperado:
se actualizan los CSV crudos por cliente y plataforma.
"@
        }

        "Ejecutar ciclo completo ahora" {
            return @"
Ejecuta el flujo completo del sistema en una sola corrida.

Hace tres cosas:
1. Exporta metricas de todos los clientes conectados.
2. Reconstruye los HMA_Master.xlsx individuales de cada cliente.
3. Ejecuta un health check para diagnosticar estado general.

Usalo cuando:
- queres probar todo el sistema de punta a punta
- queres actualizar datos y masters sin esperar 12 horas
- hiciste cambios y queres verificar que nada se rompio

Resultado esperado:
datos actualizados, masters reconstruidos y log de diagnostico generado.
"@
        }

        "Ver estado metricas clientes" {
            return @"
Muestra un resumen por cliente.

Indica:
- si Google Ads esta conectado
- si Meta Ads esta conectado
- cuantos CSV existen por plataforma
- fecha de la ultima exportacion
- si el cliente tiene datos disponibles o no

Usalo para revisar rapidamente que clientes tienen datos reales y cuales todavia estan sin conectar.
"@
        }

        "Ver automatizacion metricas 12h" {
            return @"
Muestra el estado de la tarea automatica que actualiza metricas cada 12 horas.

Sirve para saber si la tarea existe, si esta activa, pausada o si debe repararse.

Usalo cuando:
- queres comprobar si la automatizacion esta funcionando
- no ves datos nuevos
- queres confirmar que Windows Task Scheduler tiene la tarea registrada

Resultado esperado:
ver el estado de la tarea programada.
"@
        }

        "Activar metricas cada 12h" {
            return @"
Crea o repara la tarea automatica de metricas cada 12 horas.

Esta tarea ejecuta la extraccion de datos de clientes conectados sin que tengas que abrir manualmente el dashboard.

Usalo cuando:
- instalaste el sistema por primera vez
- la tarea no existe
- la tarea fue pausada
- queres reactivar la actualizacion automatica

Resultado esperado:
Windows queda programado para actualizar metricas cada 12 horas.
"@
        }

        "Pausar metricas cada 12h" {
            return @"
Desactiva la automatizacion de metricas cada 12 horas.

No borra clientes.
No borra credenciales.
No borra CSV.
No borra masters.

Solo evita que Windows ejecute la extraccion automatica.

Usalo cuando:
- estas haciendo cambios importantes
- queres evitar procesos automaticos temporalmente
- estas depurando errores
"@
        }

        "Ver ciclo completo 12h" {
            return @"
Muestra el estado del ciclo completo automatico cada 12 horas.

Este ciclo es mas completo que solo exportar metricas:
1. Exporta datos.
2. Reconstruye masters por cliente.
3. Ejecuta health check.

Usalo para confirmar si el sistema esta trabajando solo cada 12 horas.
"@
        }

        "Activar ciclo completo 12h" {
            return @"
Programa el ciclo completo cada 12 horas.

Este es el flujo automatico principal recomendado:
1. Trae metricas de Google Ads y Meta Ads.
2. Actualiza los HMA_Master.xlsx de cada cliente.
3. Ejecuta diagnostico del sistema.

Usalo cuando queres que HMA trabaje solo sin tener que correr botones manuales.
"@
        }

        "Pausar ciclo completo 12h" {
            return @"
Pausa el ciclo completo automatico cada 12 horas.

No elimina datos.
No elimina clientes.
No elimina backups.
No elimina la tarea.

Solo la deja desactivada hasta que vuelvas a activarla.

Usalo si estas haciendo mantenimiento o cambios delicados.
"@
        }

        "Ver todas las automatizaciones" {
            return @"
Genera un reporte general de todas las tareas programadas de HMA.

Incluye:
- estado de cada tarea
- ultima ejecucion
- proxima ejecucion
- resultado de la ultima corrida
- si alguna tarea no existe

Usalo como panel de control rapido para saber si el sistema automatico esta sano.
"@
        }

        "Crear cliente" {
            return @"
Crea una nueva carpeta de cliente dentro del sistema multi-cliente.

El sistema asigna un ID automatico del tipo CL-001, CL-002, CL-003 y permite escribir el nombre del cliente.

Resultado esperado:
se crea una estructura separada para ese cliente con config, historico, raw_exports y carpetas internas.
"@
        }

        "Ver clientes creados" {
            return @"
Lista todos los clientes existentes en el sistema.

Sirve para ver:
- ID del cliente
- nombre del cliente
- carpeta local
- ruta completa

Usalo cuando no recordas que clientes ya estan cargados o queres verificar la estructura multi-cliente.
"@
        }

        "Abrir carpeta clientes" {
            return @"
Abre la carpeta principal donde se guardan todos los clientes.

Cada cliente tiene su propia carpeta separada.
Ejemplo:
clientes\CL-001-Vics Solutions

Usalo para revisar archivos, masters, exports o configuraciones de un cliente.
"@
        }

        "Ver tareas HMA actual" {
            return @"
Muestra las tareas del HMA original basado en GitHub artifacts.

Estas tareas pertenecen al flujo anterior:
- descarga artifacts
- monitorea errores
- promueve pending al master

No es lo mismo que el ciclo multi-cliente nuevo.
"@
        }

        "Reanudar HMA actual" {
            return @"
Reactiva las tareas locales del HMA original basado en GitHub artifacts.

Usalo solo si queres que el flujo viejo vuelva a trabajar:
descargar artifacts, controlar errores y actualizar el master principal.
"@
        }

        "Pausar HMA actual" {
            return @"
Pausa las tareas locales del HMA original.

Conviene pausarlo si estas trabajando con el nuevo sistema multi-cliente y no queres que el flujo viejo modifique archivos al mismo tiempo.
"@
        }

        "Abrir HMA_Master.xlsx" {
            return @"
Abre el archivo HMA_Master.xlsx correspondiente a esta seccion.

En Local > HMA actual abre el master principal viejo.
En cada cliente, el master correcto esta dentro de la carpeta historico del cliente.

Usalo para revisar datos, recomendaciones y cruces de metricas.
"@
        }

        "Construir masters de todos los clientes" {
            return @"
Reconstruye los HMA_Master.xlsx individuales de cada cliente usando los CSV crudos ya exportados.

No trae datos nuevos de Google Ads o Meta Ads.
Solo toma lo que ya existe en raw_exports y lo convierte en planillas ordenadas.

Usalo cuando:
- ya exportaste metricas
- queres actualizar los excels por cliente
- queres regenerar formato, recomendaciones y cruces
"@
        }

        "Abrir historico" {
            return @"
Abre la carpeta historico.

Ahi se guardan archivos Excel principales o historicos segun la seccion del sistema.

Usalo para encontrar masters, revisar archivos generados o validar fechas de modificacion.
"@
        }

        "Abrir logs" {
            return @"
Abre la carpeta de logs del sistema.

Los logs sirven para diagnosticar:
- exportaciones
- errores
- health checks
- backups
- ciclos completos
- procesos automaticos

Usalo cuando algo no funcione o quieras revisar que hizo HMA.
"@
        }

        "Ver Git status" {
            return @"
Muestra el estado del repositorio Git local.

Sirve para saber si hay archivos modificados, archivos nuevos o cambios pendientes de commit.

Usalo antes de cerrar una etapa para confirmar si hay que guardar cambios en GitHub.
"@
        }

        "Health check sistema" {
            return @"
Ejecuta un diagnostico general del sistema.

Revisa:
- Python venv
- carpetas principales
- tareas programadas
- clientes
- masters por cliente
- CSV disponibles
- logs recientes
- estado de Git

Al final genera un TXT de diagnostico.
"@
        }

        "Ver health check semanal" {
            return @"
Muestra el estado de la tarea semanal de diagnostico.

Sirve para saber si el health check semanal esta programado, cuando corrio por ultima vez y cuando volvera a correr.
"@
        }

        "Programar health check semanal" {
            return @"
Programa un diagnostico automatico semanal.

El objetivo es detectar problemas aunque no estes usando el sistema manualmente:
- tareas caidas
- masters faltantes
- errores de logs
- clientes sin datos
- problemas de estructura
"@
        }

        "Pausar health check semanal" {
            return @"
Pausa el diagnostico semanal automatico.

No elimina reportes anteriores.
Solo evita que Windows lo ejecute semanalmente.
"@
        }

        "Crear backup local ahora" {
            return @"
Crea un backup ZIP inmediato del sistema local.

Incluye:
- scripts
- dashboard
- estructura de clientes
- historicos
- raw_exports
- masters
- logs relevantes

No deberia incluir secrets ni tokens.

Usalo antes de cambios importantes o despues de dejar algo funcionando.
"@
        }

        "Ver backup semanal" {
            return @"
Muestra el estado de la tarea automatica de backup semanal.

Sirve para confirmar si Windows esta programado para crear backups sin intervencion manual.
"@
        }

        "Programar backup semanal" {
            return @"
Programa un backup automatico semanal.

La idea es que el sistema tenga una copia recuperable aunque te olvides de hacer backup manual.

Recomendado para proteger cambios, masters, clientes y exports locales.
"@
        }

        "Pausar backup semanal" {
            return @"
Desactiva el backup semanal automatico.

No borra backups existentes.
Solo evita que se creen nuevos backups automaticamente.
"@
        }

        "Abrir carpeta backups" {
            return @"
Abre la carpeta donde se guardan los backups ZIP.

Cada ZIP representa una copia recuperable del sistema local en una fecha y hora determinada.
"@
        }

        "Limpiar backups antiguos" {
            return @"
Elimina backups viejos y conserva solo los ultimos backups definidos por el sistema.

Sirve para que la carpeta backups no crezca indefinidamente.

Actualmente conserva los ultimos 8 backups.
"@
        }

        "Restaurar backup local" {
            return @"
Abre el restaurador visual de backups.

Permite elegir un ZIP anterior y restaurarlo.

Importante:
antes de restaurar, el sistema crea un PRE_RESTORE del estado actual.
Aun asi, debe usarse con cuidado porque puede sobrescribir scripts, clientes e historicos.
"@
        }

        "Crear paquete portable" {
            return @"
Crea un ZIP portable del sistema base para llevar en pendrive.

No es un backup operativo completo.
No incluye datos reales, secrets ni tokens.

Sirve para transportar la estructura base del sistema a otra PC.
El portable final conviene generarlo cuando el sistema este terminado.
"@
        }

        default {
            if ([string]::IsNullOrWhiteSpace($defaultHelp)) {
                return "Esta accion ejecuta una funcion interna del sistema HMA. Si no estas seguro, revisa primero el estado del sistema o consulta los logs."
            }

            return $defaultHelp
        }
    }
}



function Build-HelpText($actionName, $shortHelp) {
    $map = @{
        "Conectar / editar cuenta" = "Vincula una cuenta publicitaria a un cliente. Primero elegis el cliente, despues la plataforma y luego cargas las credenciales necesarias. Usalo cuando crees un cliente nuevo, cambies una clave o tengas que corregir una conexion."
        "Actualizar metricas manual" = "Trae metricas manualmente para un cliente y una plataforma. Sirve para probar una conexion o actualizar datos sin esperar la automatizacion."
        "Actualizar Google Ads ahora" = "Exporta solo Google Ads para todos los clientes que tengan Google Ads conectado. No toca Meta Ads."
        "Actualizar Meta Ads ahora" = "Exporta solo Meta Ads para todos los clientes que tengan Meta Ads conectado. No toca Google Ads."
        "Actualizar todas las plataformas ahora" = "Exporta Google Ads y Meta Ads de todos los clientes conectados. Es una actualizacion global."
        "Ejecutar ciclo completo ahora" = "Ejecuta el flujo completo: exporta metricas, reconstruye masters por cliente y corre health check."
        "Ver estado metricas clientes" = "Muestra por cliente si Google Ads o Meta Ads estan conectados, cuantos CSV existen y cuando fue la ultima exportacion."
        "Ver automatizacion metricas 12h" = "Muestra si la tarea automatica de metricas cada 12 horas existe, esta activa o esta pausada."
        "Activar metricas cada 12h" = "Crea o reactiva la tarea de Windows que exporta metricas automaticamente cada 12 horas."
        "Pausar metricas cada 12h" = "Pausa la tarea automatica de metricas. No borra clientes, datos ni credenciales."
        "Ver ciclo completo 12h" = "Muestra el estado del ciclo completo automatico cada 12 horas: ultima ejecucion, proxima ejecucion y resultado."
        "Activar ciclo completo 12h" = "Programa el ciclo completo automatico cada 12 horas: exportar metricas, construir masters y hacer health check."
        "Pausar ciclo completo 12h" = "Pausa el ciclo completo automatico sin borrar datos ni tareas."
        "Ver todas las automatizaciones" = "Genera un reporte con todas las tareas programadas de HMA y su estado."
        "Crear cliente" = "Crea una carpeta nueva para un cliente con ID automatico, configuracion, historico y carpetas internas."
        "Ver clientes creados" = "Lista todos los clientes creados, sus IDs y sus carpetas."
        "Abrir carpeta clientes" = "Abre la carpeta principal donde estan todos los clientes separados."
        "Construir masters de todos los clientes" = "Reconstruye los HMA_Master.xlsx individuales usando los CSV ya exportados. No trae datos nuevos."
        "Health check sistema" = "Ejecuta un diagnostico del sistema: tareas, clientes, masters, CSV, logs y Git."
        "Crear backup local ahora" = "Crea un ZIP de respaldo del sistema local actual. Es util antes de cambios importantes."
        "Ver backup semanal" = "Muestra si el backup semanal esta programado y cuando corre."
        "Programar backup semanal" = "Crea la tarea semanal que genera backups automaticamente."
        "Pausar backup semanal" = "Pausa el backup semanal sin borrar backups existentes."
        "Abrir carpeta backups" = "Abre la carpeta donde se guardan los ZIP de backup."
        "Limpiar backups antiguos" = "Elimina backups viejos y conserva los ultimos definidos por el sistema."
        "Restaurar backup local" = "Abre el restaurador visual. Permite volver a un backup anterior y crea un PRE_RESTORE antes de restaurar."
        "Crear paquete portable" = "Crea un ZIP portable del sistema base para pendrive. No incluye secrets ni datos reales."
        "Ver Git status" = "Muestra archivos modificados, nuevos o pendientes de commit."
        "Abrir logs" = "Abre los logs para revisar errores, ejecuciones y diagnosticos."
        "Abrir historico" = "Abre la carpeta de historicos y planillas generadas."
        "Abrir HMA_Master.xlsx" = "Abre el Excel maestro correspondiente."
        "Ver tareas HMA actual" = "Muestra las tareas del HMA original basado en GitHub artifacts."
        "Reanudar HMA actual" = "Reactiva las tareas locales del HMA original."
        "Pausar HMA actual" = "Pausa las tareas locales del HMA original."
        "Ver health check semanal" = "Muestra si el diagnostico semanal esta programado, cuando corrio por ultima vez y cuando vuelve a correr."
        "Programar health check semanal" = "Programa un diagnostico automatico semanal del sistema."
        "Pausar health check semanal" = "Pausa el diagnostico semanal automatico sin borrar reportes existentes."
    }

    if ($map.ContainsKey($actionName)) {
        $body = $map[$actionName]
    } elseif (-not [string]::IsNullOrWhiteSpace($shortHelp)) {
        $body = $shortHelp
    } else {
        $body = "Ejecuta una accion interna del sistema HMA."
    }

    return @"
Accion:
$actionName

Que hace:
$body

Cuando usarla:
Usala cuando necesites ejecutar, revisar o corregir esta parte del sistema.

Resultado esperado:
La accion deberia abrir una ventana, generar un archivo, actualizar datos o mostrar un estado segun corresponda. Si algo falla, revisa los logs o ejecuta Health check sistema.
"@
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
        Body = (Build-HelpText $text $helpText)
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


        "Local\Backups" {
            Title "Local > Backups" 20
            Desc "Respaldo local del sistema HMA, clientes, masters y exports." 60
            Action "Crear backup local ahora" 120 { Run-Bat "backup_hma_local.bat" } "Crea un respaldo ZIP en backups/ con scripts, historico, clientes, masters y raw_exports."
            Action "Abrir carpeta backups" 165 { Open-Folder (Join-Path $BaseDir "backups") } "Abre la carpeta donde quedan guardados los ZIP de backup."
            Action "Health check sistema" 210 { Run-Bat "hma_health_check.bat" } "Ejecuta diagnostico local antes o despues de respaldar."
            Action "Ver health check semanal" 255 { Run-PS1 "scripts\status_weekly_health_check_task.ps1" } "Muestra estado, ultima ejecucion y proxima ejecucion del diagnostico semanal."
            Action "Programar health check semanal" 300 { Run-PS1 "scripts\setup_weekly_health_check_task.ps1" } "Programa un diagnostico semanal todos los lunes a las 09:20."
            Action "Pausar health check semanal" 345 { Run-PS1 "scripts\disable_weekly_health_check_task.ps1" } "Desactiva el diagnostico semanal sin borrar reportes existentes."
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
            Action "Ejecutar ciclo completo ahora" 255 { Run-Bat "hma_run_full_cycle.bat" } "Ejecuta exportacion de metricas, construccion de masters por cliente y health check en una sola corrida."
            Action "Ver ciclo completo 12h" 300 { Run-PS1 "scripts\status_full_cycle_12h_task.ps1" } "Muestra estado, ultima ejecucion y proxima ejecucion del ciclo completo automatico."
            Action "Ver todas las automatizaciones" 345 { Run-Bat "hma_automation_overview.bat" } "Muestra un reporte general de todas las tareas programadas HMA: estado, ultima ejecucion y proxima ejecucion."
            Action "Validacion QA completa" 390 { Run-Bat "hma_qa_validation.bat" } "Ejecuta una revision general: archivos, dashboard, tareas, clientes, masters, backups y Git."
            Action "Pausar tarea legacy export" 435 { Run-PS1 "scripts\disable_legacy_client_export_task.ps1" } "Pausa la tarea antigua HMA Client Ads Export Every 12 Hours para evitar duplicidad con el ciclo completo 12h."
            Action "Activar ciclo completo 12h" 345 { Run-PS1 "scripts\setup_full_cycle_12h_task.ps1" } "Programa el ciclo completo cada 12 horas: exportar metricas, construir masters y health check."
            Action "Pausar ciclo completo 12h" 390 { Run-PS1 "scripts\disable_full_cycle_12h_task.ps1" } "Desactiva el ciclo completo automatico cada 12 horas."
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

        "Local\Backups\Crear backup local ahora" { Run-Bat "backup_hma_local.bat" }
        "Local\Backups\Abrir carpeta backups" { Open-Folder (Join-Path $BaseDir "backups") }
        "Local\Backups\Limpiar backups antiguos" { Run-Bat "cleanup_hma_backups.bat" }
        "Local\Backups\Restaurar backup local" { Run-PS1 "scripts\restore_hma_backup_gui.ps1" }
        "Local\Backups\Crear paquete portable" { Run-Bat "create_hma_portable_package.bat" }
        "Local\Backups\Health check sistema" { Run-Bat "hma_health_check.bat" }
        "Local\Backups\Ver health check semanal" { Run-PS1 "scripts\status_weekly_health_check_task.ps1" }
        "Local\Backups\Programar health check semanal" { Run-PS1 "scripts\setup_weekly_health_check_task.ps1" }
        "Local\Backups\Pausar health check semanal" { Run-PS1 "scripts\disable_weekly_health_check_task.ps1" }

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
        "Administrador\Estado / Git\Ejecutar ciclo completo ahora" { Run-Bat "hma_run_full_cycle.bat" }
        "Administrador\Estado / Git\Ver ciclo completo 12h" { Run-PS1 "scripts\status_full_cycle_12h_task.ps1" }
        "Administrador\Estado / Git\Ver todas las automatizaciones" { Run-Bat "hma_automation_overview.bat" }
        "Administrador\Estado / Git\Validacion QA completa" { Run-Bat "hma_qa_validation.bat" }
        "Administrador\Estado / Git\Pausar tarea legacy export" { Run-PS1 "scripts\disable_legacy_client_export_task.ps1" }
        "Administrador\Estado / Git\Activar ciclo completo 12h" { Run-PS1 "scripts\setup_full_cycle_12h_task.ps1" }
        "Administrador\Estado / Git\Pausar ciclo completo 12h" { Run-PS1 "scripts\disable_full_cycle_12h_task.ps1" }
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
