param(
    [switch]$Silent
)

$ErrorActionPreference = "Continue"

$BaseDir = Split-Path -Parent $PSScriptRoot
$LogDir = Join-Path $BaseDir "logs\full_cycle"
$ExportAll = Join-Path $BaseDir "scripts\export_all_clients.ps1"
$BuildMasters = Join-Path $BaseDir "scripts\build_all_client_masters.ps1"
$HealthCheck = Join-Path $BaseDir "scripts\hma_health_check.ps1"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$log = Join-Path $LogDir "hma_full_cycle_$stamp.log"

function Log {
    param([string]$msg)

    $line = ("{0} | {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg)
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

function Run-Step {
    param(
        [string]$name,
        [string]$script,
        [string[]]$scriptArgs = @()
    )

    Log ("=== STEP START: {0} ===" -f $name)

    if (-not (Test-Path $script)) {
        Log ("ERROR: no existe {0}" -f $script)
        Log ("=== STEP END: {0} ===" -f $name)
        return
    }

    try {
        if ($scriptArgs.Count -gt 0) {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script @scriptArgs 2>&1 |
            ForEach-Object {
                Log ("{0}: {1}" -f $name, $_)
            }
        } else {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script 2>&1 |
            ForEach-Object {
                Log ("{0}: {1}" -f $name, $_)
            }
        }
    } catch {
        Log ("ERROR_STEP {0}: {1}" -f $name, $_.Exception.Message)
    }

    Log ("=== STEP END: {0} ===" -f $name)
}

Log "============================================================"
Log "HMA FULL CYCLE START"
Log ("BaseDir: {0}" -f $BaseDir)
Log "============================================================"

Run-Step -name "EXPORT_ALL_CLIENTS" -script $ExportAll
Run-Step -name "BUILD_CLIENT_MASTERS" -script $BuildMasters
Run-Step -name "HEALTH_CHECK" -script $HealthCheck -scriptArgs @("-Silent")

Log "============================================================"
Log "HMA FULL CYCLE END"
Log ("LOG_FILE={0}" -f $log)
Log "============================================================"

if (-not $Silent) {
    Start-Process notepad.exe $log
}
