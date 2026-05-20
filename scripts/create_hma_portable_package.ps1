$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Parent $PSScriptRoot
$PortableRoot = Join-Path $BaseDir "portable_packages"
$TempRoot = Join-Path $env:TEMP ("hma_portable_" + (Get-Date -Format "yyyyMMdd_HHmmss"))

New-Item -ItemType Directory -Force -Path $PortableRoot | Out-Null
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$packageName = "HMA_PORTABLE_$stamp"
$packageDir = Join-Path $TempRoot $packageName
$zipPath = Join-Path $PortableRoot "$packageName.zip"

New-Item -ItemType Directory -Force -Path $packageDir | Out-Null

function Copy-Safe($source, $dest) {
    if (Test-Path $source) {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
        Copy-Item $source $dest -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$items = @(
    ".gitignore",
    "README.md",
    "requirements.txt",
    "hma_manager.bat",
    "hma_manager.vbs",
    "hma_manager_v2.bat",
    "connect_ads.bat",
    "create_client.bat",
    "export_ads.bat",
    "export_all_clients.bat",
    "export_google_all_clients.bat",
    "export_meta_all_clients.bat",
    "build_all_client_masters.bat",
    "hma_health_check.bat",
    "backup_hma_local.bat",
    "cleanup_hma_backups.bat",
    "create_hma_desktop_shortcut.bat",
    "scripts",
    "clientes\_template"
)

foreach ($item in $items) {
    Copy-Safe (Join-Path $BaseDir $item) (Join-Path $packageDir $item)
}

$readme = Join-Path $packageDir "README_PORTABLE_HMA.txt"

@"
HMA PORTABLE PACKAGE

Contenido:
- Dashboard HMA Manager
- Scripts principales
- Template multi-cliente
- Exportadores Google Ads / Meta Ads
- Builder de masters por cliente
- Health check
- Backup / restore / cleanup
- Automatizaciones semanales

No incluye:
- Secrets
- Tokens
- Raw exports reales
- Masters reales de clientes
- Logs locales
- Backups previos

Uso recomendado:
1. Copiar este ZIP a un pendrive.
2. Descomprimir en una PC Windows.
3. Abrir hma_manager.bat o hma_manager.vbs.
4. Crear clientes o restaurar un backup operativo si corresponde.
5. Conectar cuentas desde el dashboard.

Nota:
Este paquete es para transportar el sistema base. Para transportar datos reales, usar backups locales.
"@ | Set-Content $readme -Encoding UTF8

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Compress-Archive -Path (Join-Path $packageDir "*") -DestinationPath $zipPath -Force

Remove-Item $TempRoot -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "PORTABLE_PACKAGE_OK"
Write-Host $zipPath

Start-Process explorer.exe $PortableRoot
