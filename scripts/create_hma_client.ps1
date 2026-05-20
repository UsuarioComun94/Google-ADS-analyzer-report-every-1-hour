param(
    [Parameter(Mandatory=$true)]
    [string]$ClientId,

    [Parameter(Mandatory=$true)]
    [string]$ClientName
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$ClientsRoot = Join-Path $ProjectRoot "clientes"
$TemplateRoot = Join-Path $ClientsRoot "_template"
$ClientRoot = Join-Path $ClientsRoot $ClientId

if ($ClientId -notmatch "^[a-zA-Z0-9_-]+$") {
    throw "ClientId inválido. Usá solo letras, números, guion o guion bajo. Ejemplo: cliente_001"
}

if (Test-Path $ClientRoot) {
    throw "El cliente ya existe: $ClientRoot"
}

$dirs = @(
    "config",
    "google_ads",
    "meta_ads",
    "historico",
    "raw_exports",
    "downloads",
    "logs",
    "error",
    "secrets"
)

foreach ($dir in $dirs) {
    New-Item -ItemType Directory -Force -Path (Join-Path $ClientRoot $dir) | Out-Null
}

$templateConfig = Join-Path $TemplateRoot "config\client_config.json"
$clientConfig = Join-Path $ClientRoot "config\client_config.json"

if (Test-Path $templateConfig) {
    Copy-Item $templateConfig $clientConfig -Force
    $config = Get-Content $clientConfig -Raw | ConvertFrom-Json

    $config.client_id = $ClientId
    $config.client_name = $ClientName
    $config.locks.writer_lock = "HMA_${ClientId}_Writer_Lock"
    $config.locks.update_master_lock = "HMA_${ClientId}_UpdateMaster_Lock"

    $config | ConvertTo-Json -Depth 10 | Set-Content $clientConfig -Encoding UTF8
} else {
@"
{
  "client_id": "$ClientId",
  "client_name": "$ClientName",
  "enabled": true,
  "platforms": {
    "google_ads": {
      "enabled": false,
      "customer_id": "",
      "login_customer_id": ""
    },
    "meta_ads": {
      "enabled": false,
      "ad_account_id": ""
    }
  },
  "paths": {
    "historico": "historico",
    "raw_exports": "raw_exports",
    "downloads": "downloads",
    "logs": "logs",
    "error": "error",
    "secrets": "secrets"
  },
  "locks": {
    "writer_lock": "HMA_${ClientId}_Writer_Lock",
    "update_master_lock": "HMA_${ClientId}_UpdateMaster_Lock"
  }
}
"@ | Set-Content $clientConfig -Encoding UTF8
}

Write-Host "CLIENTE_CREADO_OK"
Write-Host $ClientRoot
Write-Host $clientConfig
