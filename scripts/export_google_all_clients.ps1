$ErrorActionPreference = "Continue"

$BaseDir = Split-Path -Parent $PSScriptRoot
$ClientsRoot = Join-Path $BaseDir "clientes"
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"
$GoogleExporter = Join-Path $BaseDir "scripts\export_google_ads_client.py"
$LogDir = Join-Path $BaseDir "logs\client_exports"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$log = Join-Path $LogDir "export_google_all_clients_$stamp.log"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

Log "=== HMA GOOGLE ADS EXPORT START ==="

if (-not (Test-Path $ClientsRoot)) {
    Log "ERROR: no existe carpeta clientes."
    exit 1
}

if (-not (Test-Path $Python)) {
    Log "ERROR: no existe Python venv: $Python"
    exit 1
}

if (-not (Test-Path $GoogleExporter)) {
    Log "ERROR: falta export_google_ads_client.py"
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

        if ($cfg.platforms.google_ads.enabled -eq $true) {
            Log "Export Google Ads..."
            & $Python $GoogleExporter --client-dir $clientPath --date-range TODAY 2>&1 |
            ForEach-Object { Log "GOOGLE: $_" }
        } else {
            Log "Google Ads no conectado."
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

Log "=== HMA GOOGLE ADS EXPORT END ==="
Log "LOG_FILE=$log"
