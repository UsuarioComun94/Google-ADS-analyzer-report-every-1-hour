$ErrorActionPreference = "Continue"

$BaseDir = Split-Path -Parent $PSScriptRoot
$ReportDir = Join-Path $BaseDir "diagnosticos"
$ClientsRoot = Join-Path $BaseDir "clientes"
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"

New-Item -ItemType Directory -Force -Path $ReportDir | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$report = Join-Path $ReportDir "HMA_QA_VALIDATION_$stamp.txt"

function Add-Line($text = "") {
    Write-Host $text
    Add-Content -Path $report -Value $text -Encoding UTF8
}

function Check-Path($label, $path) {
    if (Test-Path $path) {
        Add-Line "[OK] $label -> $path"
    } else {
        Add-Line "[ERROR] $label NO EXISTE -> $path"
    }
}

function Check-Task($taskName) {
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if (-not $task) {
        Add-Line "[ERROR] TASK NO EXISTE -> $taskName"
        return
    }

    $info = Get-ScheduledTaskInfo -TaskName $taskName -ErrorAction SilentlyContinue

    Add-Line "[OK] TASK -> $taskName | State=$($task.State)"

    if ($info) {
        Add-Line "     LastRunTime=$($info.LastRunTime)"
        Add-Line "     LastTaskResult=$($info.LastTaskResult)"
        Add-Line "     NextRunTime=$($info.NextRunTime)"
        Add-Line "     MissedRuns=$($info.NumberOfMissedRuns)"
    }
}

Add-Line "============================================================"
Add-Line "HMA QA VALIDATION"
Add-Line "Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Add-Line "BaseDir: $BaseDir"
Add-Line "============================================================"
Add-Line ""

Add-Line "=== 1. ARCHIVOS PRINCIPALES ==="
Check-Path "Dashboard principal BAT" (Join-Path $BaseDir "hma_manager.bat")
Check-Path "Dashboard VBS sin consola" (Join-Path $BaseDir "hma_manager.vbs")
Check-Path "Dashboard V2 PS1" (Join-Path $BaseDir "scripts\hma_manager_gui_v2.ps1")
Check-Path "Crear cliente" (Join-Path $BaseDir "create_client.bat")
Check-Path "Conectar Ads" (Join-Path $BaseDir "connect_ads.bat")
Check-Path "Export manual Ads" (Join-Path $BaseDir "export_ads.bat")
Check-Path "Export global" (Join-Path $BaseDir "export_all_clients.bat")
Check-Path "Export Google only" (Join-Path $BaseDir "export_google_all_clients.bat")
Check-Path "Export Meta only" (Join-Path $BaseDir "export_meta_all_clients.bat")
Check-Path "Build masters clientes" (Join-Path $BaseDir "build_all_client_masters.bat")
Check-Path "Health check" (Join-Path $BaseDir "hma_health_check.bat")
Check-Path "Backup local" (Join-Path $BaseDir "backup_hma_local.bat")
Check-Path "Restore GUI" (Join-Path $BaseDir "scripts\restore_hma_backup_gui.ps1")
Check-Path "Full cycle" (Join-Path $BaseDir "hma_run_full_cycle.bat")
Add-Line ""

Add-Line "=== 2. VALIDACION SINTAXIS POWERSHELL DASHBOARD ==="
$dashboard = Join-Path $BaseDir "scripts\hma_manager_gui_v2.ps1"

if (Test-Path $dashboard) {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($dashboard, [ref]$tokens, [ref]$errors) | Out-Null

    if ($errors.Count -eq 0) {
        Add-Line "[OK] Dashboard V2 sin errores de sintaxis PowerShell."
    } else {
        Add-Line "[ERROR] Dashboard V2 tiene errores de sintaxis:"
        $errors | ForEach-Object { Add-Line "     $($_.Message)" }
    }
} else {
    Add-Line "[ERROR] No se pudo validar dashboard porque no existe."
}
Add-Line ""

Add-Line "=== 3. TAREAS PROGRAMADAS ==="
$tasks = @(
    "HMA Download Artifacts Every Minute",
    "HMA Error Monitor Every Minute",
    "HMA Promote Pending Master",
    "HMA Client Ads Export Every 12 Hours",
    "HMA Full Cycle Every 12 Hours",
    "HMA Weekly Local Backup",
    "HMA Weekly Health Check"
)

foreach ($task in $tasks) {
    Check-Task $task
}
Add-Line ""

Add-Line "=== 4. CLIENTES Y MASTERS ==="

if (-not (Test-Path $ClientsRoot)) {
    Add-Line "[ERROR] No existe carpeta clientes."
} else {
    $clients = @(Get-ChildItem $ClientsRoot -Directory | Where-Object { $_.Name -ne "_template" } | Sort-Object Name)

    if ($clients.Count -eq 0) {
        Add-Line "[WARN] No hay clientes creados."
    }

    foreach ($client in $clients) {
        Add-Line ""
        Add-Line "--- CLIENTE: $($client.Name) ---"

        $cfg = Join-Path $client.FullName "config\client_config.json"
        $master = Join-Path $client.FullName "historico\HMA_Master.xlsx"
        $googleRaw = Join-Path $client.FullName "raw_exports\google_ads"
        $metaRaw = Join-Path $client.FullName "raw_exports\meta_ads"

        Check-Path "Config cliente" $cfg
        Check-Path "Master cliente" $master

        $googleCsv = if (Test-Path $googleRaw) { @(Get-ChildItem $googleRaw -Filter "*.csv" -File -ErrorAction SilentlyContinue).Count } else { 0 }
        $metaCsv = if (Test-Path $metaRaw) { @(Get-ChildItem $metaRaw -Filter "*.csv" -File -ErrorAction SilentlyContinue).Count } else { 0 }

        Add-Line "Google CSV count: $googleCsv"
        Add-Line "Meta CSV count: $metaCsv"

        if (Test-Path $cfg) {
            try {
                $json = Get-Content $cfg -Raw | ConvertFrom-Json
                Add-Line "Client ID: $($json.client_id)"
                Add-Line "Client Name: $($json.client_name)"
                Add-Line "Google enabled: $($json.platforms.google_ads.enabled)"
                Add-Line "Meta enabled: $($json.platforms.meta_ads.enabled)"
            } catch {
                Add-Line "[ERROR] Config invalida: $($_.Exception.Message)"
            }
        }

        if ((Test-Path $Python) -and (Test-Path $master)) {
            $check = & $Python -c "import sys,zipfile,openpyxl; p=sys.argv[1]; print('is_zipfile:', zipfile.is_zipfile(p)); wb=openpyxl.load_workbook(p, read_only=True, data_only=True); required=['hourly_summary','campaign_metrics','metric_comparison','recommendations','creative_assets','metric_crosses']; print('sheets:', ','.join(wb.sheetnames)); print('missing:', ','.join([s for s in required if s not in wb.sheetnames]) or 'none'); print('metric_crosses_state:', wb['metric_crosses'].sheet_state if 'metric_crosses' in wb.sheetnames else 'missing'); print('recommendations_rows:', wb['recommendations'].max_row if 'recommendations' in wb.sheetnames else 'missing'); wb.close()" $master 2>&1

            $check | ForEach-Object { Add-Line "Excel: $_" }
        }
    }
}
Add-Line ""

Add-Line "=== 5. BACKUPS ==="
$BackupRoot = Join-Path $BaseDir "backups"
Check-Path "Carpeta backups" $BackupRoot

if (Test-Path $BackupRoot) {
    $zips = @(Get-ChildItem $BackupRoot -Filter "HMA_BACKUP_*.zip" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)

    Add-Line "Backups ZIP encontrados: $($zips.Count)"

    $zips | Select-Object -First 8 | ForEach-Object {
        Add-Line "     $($_.LastWriteTime) | $([Math]::Round($_.Length / 1MB, 2)) MB | $($_.Name)"
    }
}
Add-Line ""

Add-Line "=== 6. GIT STATUS ==="
$git = git -C $BaseDir status --short 2>&1

if ($git) {
    $git | ForEach-Object { Add-Line $_ }
} else {
    Add-Line "Git limpio."
}
Add-Line ""

Add-Line "=== 7. CHECKLIST MANUAL ==="
Add-Line "[ ] Abrir hma_manager.bat y confirmar que no aparece CMD detras."
Add-Line "[ ] Revisar que el arbol izquierdo muestre Google Ads, Meta Ads, Local y Administrador."
Add-Line "[ ] Confirmar que Local > Backups existe."
Add-Line "[ ] Probar botones ? y revisar que los textos sean legibles."
Add-Line "[ ] Probar doble clic o Enter en una accion del arbol izquierdo."
Add-Line "[ ] Abrir Administrador > Estado / Git > Ver todas las automatizaciones."
Add-Line "[ ] Ejecutar Health check sistema desde dashboard."
Add-Line "[ ] Abrir un master cliente y revisar formato, idioma y hojas."
Add-Line "[ ] Cuando haya cuenta real conectada: export manual -> build master -> full cycle."
Add-Line ""

Add-Line "============================================================"
Add-Line "FIN QA VALIDATION"
Add-Line "Reporte: $report"
Add-Line "============================================================"

Start-Process notepad.exe $report
