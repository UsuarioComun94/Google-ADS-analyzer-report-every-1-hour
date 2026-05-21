$ErrorActionPreference = "Continue"

$BaseDir = Split-Path -Parent $PSScriptRoot
$ReportDir = Join-Path $BaseDir "diagnosticos"

New-Item -ItemType Directory -Force -Path $ReportDir | Out-Null

$stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$report = Join-Path $ReportDir "HMA_AUTOMATION_OVERVIEW_$stamp.txt"

function Add-Line($text = "") {
    Write-Host $text
    Add-Content -Path $report -Value $text -Encoding UTF8
}

$tasks = @(
    "HMA Download Artifacts Every Minute",
    "HMA Error Monitor Every Minute",
    "HMA Promote Pending Master",
    "HMA Full Cycle Every 12 Hours",
    "HMA Weekly Local Backup",
    "HMA Weekly Health Check"
)

Add-Line "============================================================"
Add-Line "HMA AUTOMATION OVERVIEW"
Add-Line "Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Add-Line "BaseDir: $BaseDir"
Add-Line "============================================================"
Add-Line ""

foreach ($taskName in $tasks) {
    Add-Line "------------------------------------------------------------"
    Add-Line "TASK: $taskName"

    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if (-not $task) {
        Add-Line "Estado: NO EXISTE"
        Add-Line ""
        continue
    }

    $info = Get-ScheduledTaskInfo -TaskName $taskName -ErrorAction SilentlyContinue

    Add-Line "Estado: $($task.State)"

    if ($info) {
        Add-Line "Ultima ejecucion: $($info.LastRunTime)"
        Add-Line "Resultado ultima ejecucion: $($info.LastTaskResult)"
        Add-Line "Proxima ejecucion: $($info.NextRunTime)"
        Add-Line "Ejecuciones perdidas: $($info.NumberOfMissedRuns)"
    }

    try {
        $triggerText = ($task.Triggers | Format-List * | Out-String).Trim()
        Add-Line "Triggers:"
        $triggerText -split "`r?`n" | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_)) {
                Add-Line "  $_"
            }
        }
    } catch {
        Add-Line "Triggers: no disponible"
    }

    Add-Line ""
}

Add-Line "============================================================"
Add-Line "INTERPRETACION RAPIDA"
Add-Line "============================================================"
Add-Line "[Ready] significa programada y disponible."
Add-Line "[Running] significa ejecutandose ahora."
Add-Line "[Disabled] significa pausada."
Add-Line "LastTaskResult = 0 normalmente indica ejecucion correcta."
Add-Line "Si una tarea NO EXISTE, hay que recrearla desde el dashboard."
Add-Line ""
Add-Line "Reporte: $report"

Start-Process notepad.exe $report

