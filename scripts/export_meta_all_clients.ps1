$ErrorActionPreference = "Continue"

$BaseDir = Split-Path -Parent $PSScriptRoot
$ClientsRoot = Join-Path $BaseDir "clientes"
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"
$MetaExporter = Join-Path $BaseDir "scripts\export_meta_ads_client.py"
$LogDir = Join-Path $BaseDir "logs\client_exports"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$log = Join-Path $LogDir "export_meta_all_clients_$stamp.log"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

Log "=== HMA META ADS EXPORT START ==="

if (-not (Test-Path $ClientsRoot)) {
    Log "ERROR: no existe carpeta clientes."
    exit 1
}

if (-not (Test-Path $Python)) {
    Log "ERROR: no existe Python venv: $Python"
    exit 1
}

if (-not (Test-Path $MetaExporter)) {
    Log "ERROR: falta export_meta_ads_client.py"
    exit 1
}

$clients = Get-ChildItem $ClientsRoot -Directory | Where-Object { $_.Name -ne "_template" }

if (-not $clients) {
    Log "No hay clientes creados."
    exit 0
}

foreach ($clientDir in $clients) {
    $clientPath = $clientDir.FullName
    $configPath = Join-Path $clientPath "config\client_config.json"

    if (-not (Test-Path $configPath)) {
        Log "SKIP: sin config: $($clientDir.Name)"
        continue
    }

    try {
        $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
        Log "CLIENTE: $($cfg.client_id) | $($cfg.client_name)"

        if ($cfg.platforms.meta_ads.enabled -eq $true) {
            Log "Export Meta Ads..."
            & $Python $MetaExporter --client-dir $clientPath --date-preset today 2>&1 |
            ForEach-Object { Log "META: $_" }
        } else {
            Log "Meta Ads no conectado."
        }
    } catch {
        Log "ERROR cliente $($clientDir.Name): $($_.Exception.Message)"
    }
}


$BuildMastersScript = Join-Path $BaseDir "scripts\build_all_client_masters.ps1"

if (Test-Path $BuildMastersScript) {
    Log "=== BUILD CLIENT MASTERS AFTER EXPORT START ==="
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File $BuildMastersScript 2>&1 |
    ForEach-Object { Log "BUILD_MASTER: $_" }
    Log "=== BUILD CLIENT MASTERS AFTER EXPORT END ==="
} else {
    Log "SKIP_BUILD_MASTERS: no existe $BuildMastersScript"
}

Log "=== HMA META ADS EXPORT END ==="
Log "LOG_FILE=$log"
