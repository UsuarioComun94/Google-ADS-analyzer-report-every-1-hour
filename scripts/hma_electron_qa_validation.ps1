$ErrorActionPreference = "Continue"

$BaseDir = "D:\Proyectos\hma-system"
$ScriptsDir = Join-Path $BaseDir "scripts"
$DesktopDir = Join-Path $BaseDir "hma-desktop"
$DiagnosticsDir = Join-Path $BaseDir "diagnosticos"
$Stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$Report = Join-Path $DiagnosticsDir "HMA_ELECTRON_QA_VALIDATION_$Stamp.txt"

New-Item -ItemType Directory -Force -Path $DiagnosticsDir | Out-Null

$global:ErrorsFound = 0
$global:WarningsFound = 0

function Write-Line($msg = "") {
    Write-Host $msg
    Add-Content -Path $Report -Value $msg -Encoding UTF8
}

function OK($msg) {
    Write-Line "[OK] $msg"
}

function WARN($msg) {
    $global:WarningsFound++
    Write-Line "[WARN] $msg"
}

function FAIL($msg) {
    $global:ErrorsFound++
    Write-Line "[ERROR] $msg"
}

function Check-Path($label, $path, $type = "Any") {
    if ($type -eq "Directory") {
        if (Test-Path $path -PathType Container) { OK "$label -> $path" } else { FAIL "$label no existe -> $path" }
        return
    }

    if ($type -eq "File") {
        if (Test-Path $path -PathType Leaf) { OK "$label -> $path" } else { FAIL "$label no existe -> $path" }
        return
    }

    if (Test-Path $path) { OK "$label -> $path" } else { FAIL "$label no existe -> $path" }
}

function Test-PS1Syntax($path) {
    if (!(Test-Path $path -PathType Leaf)) {
        FAIL "No existe PS1 para validar sintaxis -> $path"
        return
    }

    try {
        $tokens = $null
        $parseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$parseErrors) | Out-Null

        if ($parseErrors -and $parseErrors.Count -gt 0) {
            FAIL "Sintaxis PowerShell con errores -> $path"
            foreach ($e in $parseErrors) {
                Write-Line "       Linea $($e.Extent.StartLineNumber): $($e.Message)"
            }
        } else {
            OK "Sintaxis PowerShell OK -> $path"
        }
    } catch {
        FAIL "No se pudo validar sintaxis PowerShell -> $path | $($_.Exception.Message)"
    }
}

function Test-JsonFile($label, $path) {
    if (!(Test-Path $path -PathType Leaf)) {
        FAIL "$label no existe -> $path"
        return
    }

    try {
        Get-Content $path -Raw | ConvertFrom-Json | Out-Null
        OK "$label JSON valido -> $path"
    } catch {
        FAIL "$label JSON invalido -> $path | $($_.Exception.Message)"
    }
}

function Test-Contains($label, $path, $pattern) {
    if (!(Test-Path $path -PathType Leaf)) {
        FAIL "$label no se puede revisar porque no existe -> $path"
        return
    }

    $text = Get-Content $path -Raw -ErrorAction SilentlyContinue

    if ($text -match [regex]::Escape($pattern)) {
        OK "$label contiene: $pattern"
    } else {
        FAIL "$label no contiene: $pattern"
    }
}

function Check-Task($name, $expectedRequired = $false) {
    $task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue

    if (!$task) {
        if ($expectedRequired) {
            FAIL "Tarea no existe -> $name"
        } else {
            WARN "Tarea no existe -> $name"
        }
        return
    }

    $info = Get-ScheduledTaskInfo -TaskName $name -ErrorAction SilentlyContinue

    OK "Task -> $name | State=$($task.State)"

    if ($info) {
        Write-Line "     LastRunTime=$($info.LastRunTime)"
        Write-Line "     LastTaskResult=$($info.LastTaskResult)"
        Write-Line "     NextRunTime=$($info.NextRunTime)"
    }
}

Write-Line "============================================================"
Write-Line "HMA ELECTRON QA VALIDATION"
Write-Line "Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Line "BaseDir: $BaseDir"
Write-Line "============================================================"
Write-Line ""

Write-Line "=== 1. RUTAS PRINCIPALES ==="
Check-Path "BaseDir" $BaseDir "Directory"
Check-Path "Scripts" $ScriptsDir "Directory"
Check-Path "Diagnosticos" $DiagnosticsDir "Directory"
Check-Path "Clientes" (Join-Path $BaseDir "clientes") "Directory"
Check-Path "Backups" (Join-Path $BaseDir "backups") "Directory"
Check-Path "Logs" (Join-Path $BaseDir "logs") "Directory"
Check-Path "Electron desktop" $DesktopDir "Directory"
Write-Line ""

Write-Line "=== 2. NODE / NPM / ELECTRON ==="
$node = Get-Command node -ErrorAction SilentlyContinue
$npm = Get-Command npm -ErrorAction SilentlyContinue

if ($node) {
    OK "Node detectado -> $(node -v)"
} else {
    FAIL "Node no detectado en PATH"
}

if ($npm) {
    OK "npm detectado -> $(npm -v)"
} else {
    FAIL "npm no detectado en PATH"
}

Check-Path "Electron main.js" (Join-Path $DesktopDir "electron\main.js") "File"
Check-Path "Electron preload.js" (Join-Path $DesktopDir "electron\preload.js") "File"
Check-Path "React App.jsx" (Join-Path $DesktopDir "src\App.jsx") "File"
Check-Path "React styles.css" (Join-Path $DesktopDir "src\styles.css") "File"
Check-Path "React main.jsx" (Join-Path $DesktopDir "src\main.jsx") "File"
Test-JsonFile "package.json" (Join-Path $DesktopDir "package.json")

if (Test-Path (Join-Path $DesktopDir "node_modules") -PathType Container) {
    OK "node_modules existe"
} else {
    WARN "node_modules no existe. Si Electron no abre, ejecutar npm install dentro de hma-desktop."
}
Write-Line ""

Write-Line "=== 3. CHEQUEO CODIGO ELECTRON / IPC ==="
$mainJs = Join-Path $DesktopDir "electron\main.js"
$preloadJs = Join-Path $DesktopDir "electron\preload.js"
$appJsx = Join-Path $DesktopDir "src\App.jsx"

Test-Contains "main.js" $mainJs 'ipcMain.handle("hma:run-action"'
Test-Contains "main.js" $mainJs 'HMA_NO_NOTEPAD'
Test-Contains "main.js" $mainJs 'HMA_NO_OPEN_TXT'
Test-Contains "main.js" $mainJs 'createCmdWrapper'
Test-Contains "main.js" $mainJs 'set-report-frequency:'
Test-Contains "preload.js" $preloadJs 'runAction'
Test-Contains "App.jsx" $appJsx 'choose-report-frequency'
Test-Contains "App.jsx" $appJsx 'frequencyOptions'
Test-Contains "App.jsx" $appJsx '1h'
Test-Contains "App.jsx" $appJsx '3h'
Test-Contains "App.jsx" $appJsx '5h'
Test-Contains "App.jsx" $appJsx '7h'
Test-Contains "App.jsx" $appJsx '12h'
Test-Contains "App.jsx" $appJsx '1d'
Test-Contains "App.jsx" $appJsx '2d'
Test-Contains "App.jsx" $appJsx '1w'
Write-Line ""

Write-Line "=== 4. ARCHIVOS BAT PRINCIPALES ==="
$batFiles = @(
    "hma_desktop_start.bat",
    "hma_desktop_start_debug.bat",
    "hma_manager.bat",
    "hma_manager_v2.bat",
    "hma_automation_overview.bat",
    "build_all_client_masters.bat",
    "hma_health_check.bat",
    "hma_qa_validation.bat",
    "hma_run_full_cycle.bat",
    "hma_report_frequency_status.bat",
    "hma_report_frequency_disable.bat",
    "backup_hma_local.bat",
    "create_client.bat",
    "connect_ads.bat",
    "export_ads.bat",
    "export_google_all_clients.bat",
    "export_meta_all_clients.bat"
)

foreach ($bat in $batFiles) {
    Check-Path $bat (Join-Path $BaseDir $bat) "File"
}
Write-Line ""

Write-Line "=== 5. SCRIPTS POWERSHELL CLAVE ==="
$ps1Files = @(
    "hma_report_engine.ps1",
    "setup_report_frequency_task.ps1",
    "status_report_frequency_task.ps1",
    "disable_report_frequency_task.ps1",
    "hma_qa_validation.ps1",
    "hma_automation_overview.ps1",
    "restore_hma_backup_gui.ps1"
)

foreach ($ps1 in $ps1Files) {
    $path = Join-Path $ScriptsDir $ps1
    Check-Path $ps1 $path "File"
    Test-PS1Syntax $path
}
Write-Line ""

Write-Line "=== 6. TAREAS PROGRAMADAS HMA ==="
$hmaTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like "HMA*" }

if (!$hmaTasks) {
    FAIL "No hay tareas programadas HMA."
} else {
    foreach ($task in $hmaTasks | Sort-Object TaskName) {
        $info = Get-ScheduledTaskInfo -TaskName $task.TaskName -ErrorAction SilentlyContinue
        OK "Task -> $($task.TaskName) | State=$($task.State)"
        if ($info) {
            Write-Line "     LastRunTime=$($info.LastRunTime)"
            Write-Line "     LastTaskResult=$($info.LastTaskResult)"
            Write-Line "     NextRunTime=$($info.NextRunTime)"
        }
    }
}

Write-Line ""
Write-Line "=== 7. AUTOMATIZACION DE INFORMES ACTIVA ==="
$reportTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like "HMA Informe *" }

if (!$reportTasks) {
    WARN "No hay frecuencia de informes activa. Esto es normal si todavia no configuraste 1h/3h/5h/7h/12h/1d/2d/1w desde el dashboard."
} else {
    foreach ($task in $reportTasks) {
        $info = Get-ScheduledTaskInfo -TaskName $task.TaskName -ErrorAction SilentlyContinue
        OK "Informe activo -> $($task.TaskName) | State=$($task.State)"
        if ($info) {
            Write-Line "     LastRunTime=$($info.LastRunTime)"
            Write-Line "     LastTaskResult=$($info.LastTaskResult)"
            Write-Line "     NextRunTime=$($info.NextRunTime)"
        }
    }
}
Write-Line ""

Write-Line "=== 8. CLIENTES / INFORMES / MASTERS ==="
$clientesDir = Join-Path $BaseDir "clientes"

if (!(Test-Path $clientesDir)) {
    FAIL "No existe carpeta clientes."
} else {
    $clientes = Get-ChildItem $clientesDir -Directory | Where-Object { $_.Name -notlike "_*" }

    if (!$clientes) {
        WARN "No hay clientes creados."
    }

    foreach ($cliente in $clientes) {
        Write-Line ""
        Write-Line "--- CLIENTE: $($cliente.Name) ---"

        $config = Join-Path $cliente.FullName "config\client_config.json"
        $master = Join-Path $cliente.FullName "historico\HMA_Master.xlsx"
        $informes = Join-Path $cliente.FullName "informes"

        Check-Path "Config cliente" $config "File"
        Check-Path "Master cliente" $master "File"

        if (Test-Path $informes -PathType Container) {
            OK "Carpeta informes -> $informes"

            $expectedFolders = @("Informe_1h","Informe_3h","Informe_5h","Informe_7h","Informe_12h","Informe_1d","Informe_2d","Informe_1w")

            foreach ($folder in $expectedFolders) {
                $folderPath = Join-Path $informes $folder
                if (Test-Path $folderPath -PathType Container) {
                    OK "Informe folder existe -> $folder"
                } else {
                    WARN "Informe folder aun no existe -> $folder"
                }
            }
        } else {
            WARN "Carpeta informes aun no existe para cliente. Se crea cuando ejecutes una frecuencia."
        }
    }
}
Write-Line ""

Write-Line "=== 9. BACKUPS ==="
$backupDir = Join-Path $BaseDir "backups"

if (Test-Path $backupDir -PathType Container) {
    OK "Carpeta backups existe -> $backupDir"

    $zips = Get-ChildItem $backupDir -Filter "*.zip" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending

    Write-Line "Backups ZIP encontrados: $($zips.Count)"

    foreach ($zip in $zips | Select-Object -First 10) {
        Write-Line "     $($zip.LastWriteTime) | $([math]::Round($zip.Length / 1MB, 2)) MB | $($zip.Name)"
    }
} else {
    FAIL "No existe carpeta backups."
}
Write-Line ""

Write-Line "=== 10. PRUEBA SEGURA DE CMD WRAPPER ==="
$runtimeDir = Join-Path $DesktopDir ".runtime"
New-Item -ItemType Directory -Force -Path $runtimeDir | Out-Null

$testCmd = Join-Path $runtimeDir "hma_cmd_wrapper_smoke_test.cmd"

@"
@echo off
cd /d "$BaseDir"
echo HMA_CMD_WRAPPER_OK
echo BASE=%CD%
exit /b 0
"@ | Set-Content $testCmd -Encoding ASCII

$cmdOutput = cmd.exe /d /c $testCmd 2>&1
$exitCode = $LASTEXITCODE

Remove-Item $testCmd -Force -ErrorAction SilentlyContinue

if ($exitCode -eq 0 -and ($cmdOutput -join "`n") -match "HMA_CMD_WRAPPER_OK") {
    OK "CMD wrapper smoke test OK"
    Write-Line ($cmdOutput -join "`n")
} else {
    FAIL "CMD wrapper smoke test fallo"
    Write-Line ($cmdOutput -join "`n")
}
Write-Line ""

Write-Line "=== 11. GIT STATUS ==="
$git = Get-Command git -ErrorAction SilentlyContinue

if (!$git) {
    WARN "Git no detectado en PATH."
} else {
    $gitStatus = git -C $BaseDir status --short 2>&1

    if ($gitStatus) {
        WARN "Git tiene cambios pendientes:"
        $gitStatus | ForEach-Object { Write-Line "     $_" }
    } else {
        OK "Git limpio."
    }
}
Write-Line ""

Write-Line "=== 12. RESULTADO FINAL ==="

if ($global:ErrorsFound -eq 0 -and $global:WarningsFound -eq 0) {
    Write-Line "RESULTADO: OK TOTAL"
} elseif ($global:ErrorsFound -eq 0 -and $global:WarningsFound -gt 0) {
    Write-Line "RESULTADO: OK CON WARNINGS"
} else {
    Write-Line "RESULTADO: HAY ERRORES"
}

Write-Line "Errores: $global:ErrorsFound"
Write-Line "Warnings: $global:WarningsFound"
Write-Line "Reporte: $Report"
Write-Line "============================================================"

Write-Host ""
Write-Host "Reporte generado:"
Write-Host $Report
