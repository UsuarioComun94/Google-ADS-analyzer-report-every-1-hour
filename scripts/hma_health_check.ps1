param([switch]$Silent)

$ErrorActionPreference = "Continue"

$BaseDir = Split-Path -Parent $PSScriptRoot
$ClientsRoot = Join-Path $BaseDir "clientes"
$LogsDir = Join-Path $BaseDir "logs"
$ReportDir = Join-Path $BaseDir "diagnosticos"
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"

New-Item -ItemType Directory -Force -Path $ReportDir | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$report = Join-Path $ReportDir "HMA_HEALTH_CHECK_$stamp.txt"

function Add-Line($text = "") {
    Write-Host $text
    Add-Content -Path $report -Value $text -Encoding UTF8
}

function Status($name, $ok, $detail) {
    $state = if ($ok) { "OK" } else { "ERROR" }
    Add-Line ("[{0}] {1} - {2}" -f $state, $name, $detail)
}

Add-Line "============================================================"
Add-Line "HMA HEALTH CHECK"
Add-Line "Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Add-Line "BaseDir: $BaseDir"
Add-Line "============================================================"
Add-Line ""

Status "Python venv" (Test-Path $Python) $Python
Status "Carpeta clientes" (Test-Path $ClientsRoot) $ClientsRoot
Status "Carpeta logs" (Test-Path $LogsDir) $LogsDir

Add-Line ""
Add-Line "=== TAREAS PROGRAMADAS ==="

$taskNames = @(
    "HMA Download Artifacts Every Minute",
    "HMA Error Monitor Every Minute",
    "HMA Promote Pending Master",
    "HMA Client Ads Export Every 12 Hours"
)

foreach ($taskName in $taskNames) {
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if ($task) {
        Status $taskName $true "State=$($task.State)"
    } else {
        Status $taskName $false "No existe"
    }
}

Add-Line ""
Add-Line "=== CLIENTES ==="

if (Test-Path $ClientsRoot) {
    $clients = Get-ChildItem $ClientsRoot -Directory | Where-Object { $_.Name -ne "_template" } | Sort-Object Name

    if (-not $clients) {
        Add-Line "[INFO] No hay clientes creados."
    }

    foreach ($client in $clients) {
        Add-Line ""
        Add-Line "--- $($client.Name) ---"

        $cfgPath = Join-Path $client.FullName "config\client_config.json"
        $master = Join-Path $client.FullName "historico\HMA_Master.xlsx"
        $googleRaw = Join-Path $client.FullName "raw_exports\google_ads"
        $metaRaw = Join-Path $client.FullName "raw_exports\meta_ads"

        Status "Config cliente" (Test-Path $cfgPath) $cfgPath
        Status "Master cliente" (Test-Path $master) $master

        if (Test-Path $master) {
            $item = Get-Item $master
            Add-Line "Master LastWriteTime: $($item.LastWriteTime)"
            Add-Line "Master Length: $($item.Length)"
        }

        $googleCount = if (Test-Path $googleRaw) { @(Get-ChildItem $googleRaw -Filter "*.csv" -File -ErrorAction SilentlyContinue).Count } else { 0 }
        $metaCount = if (Test-Path $metaRaw) { @(Get-ChildItem $metaRaw -Filter "*.csv" -File -ErrorAction SilentlyContinue).Count } else { 0 }

        Add-Line "Google CSV count: $googleCount"
        Add-Line "Meta CSV count: $metaCount"

        if (Test-Path $cfgPath) {
            try {
                $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
                Add-Line "Client ID: $($cfg.client_id)"
                Add-Line "Client Name: $($cfg.client_name)"
                Add-Line "Google Ads enabled: $($cfg.platforms.google_ads.enabled)"
                Add-Line "Meta Ads enabled: $($cfg.platforms.meta_ads.enabled)"
            } catch {
                Add-Line "[ERROR] Config invalida: $($_.Exception.Message)"
            }
        }

        if ((Test-Path $Python) -and (Test-Path $master)) {
            $check = & $Python -c "import sys,zipfile,openpyxl; p=sys.argv[1]; print('is_zipfile:',zipfile.is_zipfile(p)); wb=openpyxl.load_workbook(p,read_only=True,data_only=True); print('sheets:', ','.join(wb.sheetnames)); print('has_recommendations:', 'recommendations' in wb.sheetnames); print('has_metric_crosses:', 'metric_crosses' in wb.sheetnames); print('metric_crosses_state:', wb['metric_crosses'].sheet_state if 'metric_crosses' in wb.sheetnames else 'missing'); wb.close()" $master 2>&1
            Add-Line "Excel check:"
            $check | ForEach-Object { Add-Line "  $_" }
        }
    }
}

Add-Line ""
Add-Line "=== LOGS RECIENTES ==="

if (Test-Path $LogsDir) {
    Get-ChildItem $LogsDir -Recurse -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 15 |
    ForEach-Object {
        Add-Line "$($_.LastWriteTime) | $($_.FullName)"
    }
}

Add-Line ""
Add-Line "=== GIT STATUS ==="
$git = git -C $BaseDir status --short 2>&1
if ($git) {
    $git | ForEach-Object { Add-Line $_ }
} else {
    Add-Line "Git limpio."
}

Add-Line ""
Add-Line "============================================================"
Add-Line "FIN HEALTH CHECK"
Add-Line "Reporte: $report"
Add-Line "============================================================"

if (-not $Silent) { Start-Process notepad.exe $report }
