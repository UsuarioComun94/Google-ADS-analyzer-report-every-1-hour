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

function ConvertTo-SafeFolderName {
    param([string]$Name)

    $safe = $Name.Trim()
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()

    foreach ($ch in $invalid) {
        $safe = $safe.Replace($ch, "-")
    }

    $safe = $safe -replace "\s+", " "
    $safe = $safe.Trim(" .-_")

    return $safe
}

if ($ClientId -notmatch "^[a-zA-Z0-9_-]+$") {
    throw "ClientId inválido. Usá solo letras, números, guion o guion bajo. Ejemplo: CL-001"
}

$safeClientName = ConvertTo-SafeFolderName $ClientName
$ClientFolderName = "$ClientId-$safeClientName"
$ClientRoot = Join-Path $ClientsRoot $ClientFolderName

if (Test-Path $ClientRoot) {
    throw "El cliente ya existe: $ClientRoot"
}

$existingClient = Get-ChildItem $ClientsRoot -Directory -ErrorAction SilentlyContinue |
Where-Object { $_.Name -ne "_template" } |
Where-Object {
    $cfgPath = Join-Path $_.FullName "config\client_config.json"
    if (Test-Path $cfgPath) {
        try {
            $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
            return $cfg.client_id -eq $ClientId
        } catch {
            return $false
        }
    }
    return $false
} |
Select-Object -First 1

if ($existingClient) {
    throw "Ya existe un cliente con ClientId $ClientId en: $($existingClient.FullName)"
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
    $config | Add-Member -NotePropertyName "client_folder" -NotePropertyValue $ClientFolderName -Force
    $config.locks.writer_lock = "HMA_${ClientId}_Writer_Lock"
    $config.locks.update_master_lock = "HMA_${ClientId}_UpdateMaster_Lock"

    $config | ConvertTo-Json -Depth 10 | Set-Content $clientConfig -Encoding UTF8
} else {
@"
{
  "client_id": "$ClientId",
  "client_name": "$ClientName",
  "client_folder": "$ClientFolderName",
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
