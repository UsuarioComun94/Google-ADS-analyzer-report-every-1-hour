param([switch]$Silent)

$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Parent $PSScriptRoot
$BackupRoot = Join-Path $BaseDir "backups"
$ClientsRoot = Join-Path $BaseDir "clientes"
$HistoricoDir = Join-Path $BaseDir "historico"
$ScriptsDir = Join-Path $BaseDir "scripts"
$LogsDir = Join-Path $BaseDir "logs"

New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$backupDir = Join-Path $BackupRoot "HMA_BACKUP_$stamp"
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null

function Copy-Safe($source, $dest) {
    if (Test-Path $source) {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
        Copy-Item $source $dest -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Copy-FileSafe($source, $dest) {
    if (Test-Path $source) {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
        Copy-Item $source $dest -Force -ErrorAction SilentlyContinue
    }
}

$manifest = Join-Path $backupDir "BACKUP_MANIFEST.txt"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
    Write-Host $line
    Add-Content -Path $manifest -Value $line -Encoding UTF8
}

Log "=== HMA BACKUP START ==="
Log "BaseDir: $BaseDir"
Log "BackupDir: $backupDir"

Copy-FileSafe (Join-Path $BaseDir ".gitignore") (Join-Path $backupDir ".gitignore")
Copy-FileSafe (Join-Path $BaseDir "hma_manager.bat") (Join-Path $backupDir "hma_manager.bat")
Copy-FileSafe (Join-Path $BaseDir "hma_manager.vbs") (Join-Path $backupDir "hma_manager.vbs")
Copy-FileSafe (Join-Path $BaseDir "hma_manager_v2.bat") (Join-Path $backupDir "hma_manager_v2.bat")
Copy-FileSafe (Join-Path $BaseDir "connect_ads.bat") (Join-Path $backupDir "connect_ads.bat")
Copy-FileSafe (Join-Path $BaseDir "create_client.bat") (Join-Path $backupDir "create_client.bat")
Copy-FileSafe (Join-Path $BaseDir "export_ads.bat") (Join-Path $backupDir "export_ads.bat")
Copy-FileSafe (Join-Path $BaseDir "export_all_clients.bat") (Join-Path $backupDir "export_all_clients.bat")
Copy-FileSafe (Join-Path $BaseDir "export_google_all_clients.bat") (Join-Path $backupDir "export_google_all_clients.bat")
Copy-FileSafe (Join-Path $BaseDir "export_meta_all_clients.bat") (Join-Path $backupDir "export_meta_all_clients.bat")
Copy-FileSafe (Join-Path $BaseDir "build_all_client_masters.bat") (Join-Path $backupDir "build_all_client_masters.bat")
Copy-FileSafe (Join-Path $BaseDir "hma_health_check.bat") (Join-Path $backupDir "hma_health_check.bat")

Copy-Safe $ScriptsDir (Join-Path $backupDir "scripts")
Copy-Safe $HistoricoDir (Join-Path $backupDir "historico")

if (Test-Path $ClientsRoot) {
    $clientBackupRoot = Join-Path $backupDir "clientes"
    New-Item -ItemType Directory -Force -Path $clientBackupRoot | Out-Null

    Get-ChildItem $ClientsRoot -Directory |
    Where-Object { $_.Name -ne "_template" } |
    ForEach-Object {
        $clientDir = $_.FullName
        $clientDest = Join-Path $clientBackupRoot $_.Name

        New-Item -ItemType Directory -Force -Path $clientDest | Out-Null

        Copy-Safe (Join-Path $clientDir "config") (Join-Path $clientDest "config")
        Copy-Safe (Join-Path $clientDir "historico") (Join-Path $clientDest "historico")
        Copy-Safe (Join-Path $clientDir "raw_exports") (Join-Path $clientDest "raw_exports")
        Copy-Safe (Join-Path $clientDir "logs") (Join-Path $clientDest "logs")
        Copy-Safe (Join-Path $clientDir "error") (Join-Path $clientDest "error")

        Log "Cliente respaldado: $($_.Name)"
    }

    Copy-Safe (Join-Path $ClientsRoot "_template") (Join-Path $clientBackupRoot "_template")
}

$zipPath = Join-Path $BackupRoot "HMA_BACKUP_$stamp.zip"

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Compress-Archive -Path (Join-Path $backupDir "*") -DestinationPath $zipPath -Force

Log "ZIP creado: $zipPath"
Log "=== HMA BACKUP END ==="

if (-not $Silent) { Start-Process explorer.exe $BackupRoot }
