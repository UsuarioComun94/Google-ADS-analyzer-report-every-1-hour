$ErrorActionPreference = "Continue"

$BaseDir = "D:\Proyectos\hma-system"
$AppPath = Join-Path $BaseDir "hma-desktop\src\App.jsx"
$MainPath = Join-Path $BaseDir "hma-desktop\electron\main.js"
$ReportDir = Join-Path $BaseDir "diagnosticos"
$Stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$Report = Join-Path $ReportDir "HMA_ELECTRON_ACTION_MAP_VALIDATION_$Stamp.txt"

New-Item -ItemType Directory -Force -Path $ReportDir | Out-Null

$global:ErrorsFound = 0
$global:WarningsFound = 0

function Line($msg = "") {
    Write-Host $msg
    Add-Content -Path $Report -Value $msg -Encoding UTF8
}

function OK($msg) {
    Line "[OK] $msg"
}

function WARN($msg) {
    $global:WarningsFound++
    Line "[WARN] $msg"
}

function FAIL($msg) {
    $global:ErrorsFound++
    Line "[ERROR] $msg"
}

Line "============================================================"
Line "HMA ELECTRON ACTION MAP VALIDATION"
Line "Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Line "BaseDir: $BaseDir"
Line "============================================================"
Line ""

if (!(Test-Path $AppPath)) {
    FAIL "No existe App.jsx -> $AppPath"
    exit 1
}

if (!(Test-Path $MainPath)) {
    FAIL "No existe main.js -> $MainPath"
    exit 1
}

$app = Get-Content $AppPath -Raw
$main = Get-Content $MainPath -Raw

Line "=== 1. COMANDOS DECLARADOS EN APP.JSX ==="

$matches = [regex]::Matches($app, 'command:\s*"([^"]+)"')
$commands = @()

foreach ($m in $matches) {
    $commands += $m.Groups[1].Value
}

$commands = $commands | Sort-Object -Unique

if (!$commands -or $commands.Count -eq 0) {
    FAIL "No se encontraron comandos en App.jsx"
} else {
    OK "Comandos encontrados: $($commands.Count)"
}

Line ""

$specialCommands = @(
    "git-status",
    "clients-list",
    "choose-report-frequency"
)

foreach ($command in $commands) {
    if ($specialCommands -contains $command) {
        OK "Comando especial válido -> $command"
        continue
    }

    if ($command -like "set-report-frequency:*") {
        OK "Comando frecuencia válido -> $command"
        continue
    }

    $target = Join-Path $BaseDir $command

    if (Test-Path $target) {
        OK "Ruta válida -> $command"
    } else {
        FAIL "Ruta inexistente -> $command | Esperado: $target"
    }
}

Line ""
Line "=== 2. FRECUENCIAS FIJAS ==="

$requiredFrequencies = @("1h","3h","5h","7h","12h","1d","2d","1w")

foreach ($freq in $requiredFrequencies) {
    if ($app -match [regex]::Escape($freq)) {
        OK "Frecuencia presente -> $freq"
    } else {
        FAIL "Frecuencia faltante -> $freq"
    }
}

if ($app -match "Personalizado|custom|customized") {
    FAIL "La UI todavía contiene opción personalizada."
} else {
    OK "No hay opción personalizada."
}

Line ""
Line "=== 3. TEXTO UI AUTOMATIZACION ==="

if ($app -match "Automatizacion 12h|Automatización 12h") {
    WARN "Todavía aparece texto 'Automatizacion 12h'. Conviene dejar solo 'Automatizacion'."
} else {
    OK "Texto rígido 'Automatizacion 12h' no detectado."
}

if ($app -match "Automatizacion|Automatización") {
    OK "Texto general 'Automatizacion' detectado."
} else {
    WARN "No se detectó texto 'Automatizacion'."
}

Line ""
Line "=== 4. IPC / EJECUCION REAL ==="

$checks = @(
    'ipcMain.handle("hma:run-action"',
    'contextIsolation: true',
    'HMA_NO_NOTEPAD',
    'HMA_NO_OPEN_TXT',
    'createCmdWrapper',
    'set-report-frequency:'
)

foreach ($check in $checks) {
    if ($main -match [regex]::Escape($check)) {
        OK "main.js contiene -> $check"
    } else {
        FAIL "main.js NO contiene -> $check"
    }
}

Line ""
Line "=== 5. SCRIPTS DE FRECUENCIA ==="

$freqFiles = @(
    "scripts\hma_report_engine.ps1",
    "scripts\setup_report_frequency_task.ps1",
    "scripts\status_report_frequency_task.ps1",
    "scripts\disable_report_frequency_task.ps1",
    "hma_report_frequency_status.bat",
    "hma_report_frequency_disable.bat"
)

foreach ($file in $freqFiles) {
    $target = Join-Path $BaseDir $file
    if (Test-Path $target) {
        OK "Existe -> $file"
    } else {
        FAIL "Falta -> $file"
    }
}

Line ""
Line "=== 6. GIT STATUS ==="

$gitStatus = git -C $BaseDir status --short 2>&1

if ($gitStatus) {
    WARN "Git tiene cambios pendientes:"
    $gitStatus | ForEach-Object { Line "     $_" }
} else {
    OK "Git limpio."
}

Line ""
Line "=== RESULTADO FINAL ==="
Line "Errores: $global:ErrorsFound"
Line "Warnings: $global:WarningsFound"

if ($global:ErrorsFound -eq 0 -and $global:WarningsFound -eq 0) {
    Line "RESULTADO: OK TOTAL"
} elseif ($global:ErrorsFound -eq 0) {
    Line "RESULTADO: OK CON WARNINGS"
} else {
    Line "RESULTADO: HAY ERRORES"
}

Line "Reporte: $Report"
Line "============================================================"
