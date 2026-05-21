param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("1h","3h","5h","7h","12h","1d","2d","1w")]
    [string]$Frequency
)

$ErrorActionPreference = "Continue"

$BaseDir = "D:\Proyectos\hma-system"
$env:HMA_DASHBOARD_MODE = "1"
$env:HMA_NO_NOTEPAD = "1"
$env:HMA_NO_OPEN_TXT = "1"

$labels = @{
    "1h"  = "Informe_1h"
    "3h"  = "Informe_3h"
    "5h"  = "Informe_5h"
    "7h"  = "Informe_7h"
    "12h" = "Informe_12h"
    "1d"  = "Informe_1d"
    "2d"  = "Informe_2d"
    "1w"  = "Informe_1w"
}

$label = $labels[$Frequency]
$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logDir = Join-Path $BaseDir "logs\report_engine"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir "hma_report_engine_${Frequency}_$stamp.log"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

Log "============================================================"
Log "HMA REPORT ENGINE START"
Log "Frequency: $Frequency"
Log "ReportLabel: $label"
Log "BaseDir: $BaseDir"
Log "============================================================"

$fullCycleBat = Join-Path $BaseDir "hma_run_full_cycle.bat"

if (Test-Path $fullCycleBat) {
    Log "Ejecutando ciclo completo base..."
    cmd.exe /d /c call "`"$fullCycleBat`"" 2>&1 | ForEach-Object { Log $_ }
} else {
    Log "ADVERTENCIA: no existe hma_run_full_cycle.bat. Se intentara construir informes con masters existentes."
}

$clientesDir = Join-Path $BaseDir "clientes"

if (!(Test-Path $clientesDir)) {
    Log "ERROR: no existe carpeta clientes."
    exit 1
}

$clientes = Get-ChildItem $clientesDir -Directory | Where-Object { $_.Name -notlike "_*" }

foreach ($cliente in $clientes) {
    $clientName = $cliente.Name
    $master = Join-Path $cliente.FullName "historico\HMA_Master.xlsx"
    $reportDir = Join-Path $cliente.FullName "informes\$label"
    $reportFile = Join-Path $reportDir "HMA_$label.xlsx"
    $manifest = Join-Path $reportDir "HMA_${label}_manifest.txt"

    New-Item -ItemType Directory -Force -Path $reportDir | Out-Null

    Log "Cliente: $clientName"

    if (Test-Path $master) {
        Copy-Item $master $reportFile -Force
        Log "[OK] Informe generado: $reportFile"
    } else {
        Log "[FALTA] No existe master para cliente: $master"
    }

    @"
HMA REPORT MANIFEST
Fecha: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Cliente: $clientName
Frecuencia: $Frequency
Etiqueta: $label
Archivo: $reportFile
Origen: $master

Nota:
Este informe usa el HMA_Master.xlsx del cliente como fuente base.
La siguiente etapa puede agregar hojas filtradas/agregadas por frecuencia real.
"@ | Set-Content $manifest -Encoding UTF8

    Log "[OK] Manifest generado: $manifest"
}

Log "============================================================"
Log "HMA REPORT ENGINE END"
Log "Log: $log"
Log "============================================================"
