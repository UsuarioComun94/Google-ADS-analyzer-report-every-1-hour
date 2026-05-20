param(
    [int]$KeepLast = 8
)

$ErrorActionPreference = "Continue"

$BaseDir = Split-Path -Parent $PSScriptRoot
$BackupRoot = Join-Path $BaseDir "backups"
$LogDir = Join-Path $BaseDir "logs\backup_cleanup"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$log = Join-Path $LogDir "cleanup_hma_backups_$stamp.log"

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

Log "=== HMA BACKUP CLEANUP START ==="
Log "BackupRoot: $BackupRoot"
Log "KeepLast: $KeepLast"

if (-not (Test-Path $BackupRoot)) {
    Log "No existe carpeta backups. Nada para limpiar."
    exit 0
}

$zipBackups = @(Get-ChildItem $BackupRoot -Filter "HMA_BACKUP_*.zip" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
$dirBackups = @(Get-ChildItem $BackupRoot -Directory -Filter "HMA_BACKUP_*" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)

Log "ZIP encontrados: $($zipBackups.Count)"
Log "Carpetas encontradas: $($dirBackups.Count)"

$zipToDelete = @()
$dirToDelete = @()

if ($zipBackups.Count -gt $KeepLast) {
    $zipToDelete = $zipBackups | Select-Object -Skip $KeepLast
}

if ($dirBackups.Count -gt $KeepLast) {
    $dirToDelete = $dirBackups | Select-Object -Skip $KeepLast
}

foreach ($item in $zipToDelete) {
    try {
        Remove-Item $item.FullName -Force
        Log "ZIP eliminado: $($item.FullName)"
    } catch {
        Log "ERROR eliminando ZIP $($item.FullName): $($_.Exception.Message)"
    }
}

foreach ($item in $dirToDelete) {
    try {
        Remove-Item $item.FullName -Recurse -Force
        Log "Carpeta eliminada: $($item.FullName)"
    } catch {
        Log "ERROR eliminando carpeta $($item.FullName): $($_.Exception.Message)"
    }
}

Log "Backups ZIP conservados: $([Math]::Min($zipBackups.Count, $KeepLast))"
Log "Backups carpeta conservados: $([Math]::Min($dirBackups.Count, $KeepLast))"
Log "=== HMA BACKUP CLEANUP END ==="
Log "LOG_FILE=$log"
