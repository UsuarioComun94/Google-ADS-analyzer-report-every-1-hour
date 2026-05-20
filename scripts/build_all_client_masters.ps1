$ErrorActionPreference = "Continue"

$BaseDir = Split-Path -Parent $PSScriptRoot
$ClientsRoot = Join-Path $BaseDir "clientes"
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"
$Builder = Join-Path $BaseDir "scripts\build_client_master_from_exports.py"
$LogDir = Join-Path $BaseDir "logs\client_masters"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$log = Join-Path $LogDir "build_all_client_masters_$stamp.log"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

Log "=== BUILD CLIENT MASTERS START ==="

if (-not (Test-Path $ClientsRoot)) {
    Log "ERROR: no existe carpeta clientes."
    exit 1
}

if (-not (Test-Path $Python)) {
    Log "ERROR: no existe Python venv: $Python"
    exit 1
}

if (-not (Test-Path $Builder)) {
    Log "ERROR: no existe builder: $Builder"
    exit 1
}

$clients = Get-ChildItem $ClientsRoot -Directory | Where-Object { $_.Name -ne "_template" }

if (-not $clients) {
    Log "No hay clientes creados."
    exit 0
}

foreach ($clientDir in $clients) {
    try {
        Log "CLIENTE: $($clientDir.Name)"
        & $Python $Builder --client-dir $clientDir.FullName 2>&1 |
        ForEach-Object { Log $_ }
    } catch {
        Log "ERROR cliente $($clientDir.Name): $($_.Exception.Message)"
    }
}

Log "=== BUILD CLIENT MASTERS END ==="
Log "LOG_FILE=$log"
