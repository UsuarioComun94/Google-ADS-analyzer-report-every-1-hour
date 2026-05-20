Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Parent $PSScriptRoot
$BackupRoot = Join-Path $BaseDir "backups"
$PreRestoreRoot = Join-Path $BaseDir "backups\pre_restore"
$TempRoot = Join-Path $env:TEMP ("hma_restore_" + (Get-Date -Format "yyyyMMdd_HHmmss"))

function Msg($title, $text) {
    [System.Windows.Forms.MessageBox]::Show($text, $title, "OK", "Information") | Out-Null
}

function Err($text) {
    [System.Windows.Forms.MessageBox]::Show($text, "HMA Restore - Error", "OK", "Error") | Out-Null
}

function Confirm($text) {
    return ([System.Windows.Forms.MessageBox]::Show($text, "Confirmar restauracion", "YesNo", "Warning") -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Get-Backups {
    if (-not (Test-Path $BackupRoot)) {
        return @()
    }

    return @(Get-ChildItem $BackupRoot -Filter "HMA_BACKUP_*.zip" -File | Sort-Object LastWriteTime -Descending)
}

function Make-PreRestoreBackup {
    New-Item -ItemType Directory -Force -Path $PreRestoreRoot | Out-Null

    $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $dest = Join-Path $PreRestoreRoot "PRE_RESTORE_$stamp"

    New-Item -ItemType Directory -Force -Path $dest | Out-Null

    $items = @(
        "scripts",
        "historico",
        "clientes",
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
        ".gitignore"
    )

    foreach ($item in $items) {
        $src = Join-Path $BaseDir $item
        $dst = Join-Path $dest $item

        if (Test-Path $src) {
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst) | Out-Null
            Copy-Item $src $dst -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    return $dest
}

function Restore-Backup($zipPath) {
    if (-not (Test-Path $zipPath)) {
        Err "No existe el backup seleccionado:`n$zipPath"
        return
    }

    $question = @"
Vas a restaurar este backup:

$zipPath

Esto puede sobrescribir scripts, clientes, historico y archivos BAT/VBS actuales.

Antes de restaurar se creara un PRE_RESTORE local.

Continuar?
"@

    if (-not (Confirm $question)) {
        return
    }

    $preRestore = Make-PreRestoreBackup

    if (Test-Path $TempRoot) {
        Remove-Item $TempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

    Expand-Archive -Path $zipPath -DestinationPath $TempRoot -Force

    $items = @(
        "scripts",
        "historico",
        "clientes",
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
        ".gitignore"
    )

    foreach ($item in $items) {
        $src = Join-Path $TempRoot $item
        $dst = Join-Path $BaseDir $item

        if (Test-Path $src) {
            if (Test-Path $dst) {
                Remove-Item $dst -Recurse -Force -ErrorAction SilentlyContinue
            }

            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dst) | Out-Null
            Copy-Item $src $dst -Recurse -Force
        }
    }

    Remove-Item $TempRoot -Recurse -Force -ErrorAction SilentlyContinue

    Msg "Restauracion completada" "Backup restaurado correctamente.`n`nPre-restore guardado en:`n$preRestore"
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "HMA Restore Backup"
$form.Size = New-Object System.Drawing.Size(820,520)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$title = New-Object System.Windows.Forms.Label
$title.Text = "Restaurar backup local HMA"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(20,20)
$title.Size = New-Object System.Drawing.Size(760,35)
$form.Controls.Add($title)

$desc = New-Object System.Windows.Forms.Label
$desc.Text = "Selecciona un ZIP de backup. Antes de restaurar se crea una copia PRE_RESTORE del estado actual."
$desc.Location = New-Object System.Drawing.Point(20,60)
$desc.Size = New-Object System.Drawing.Size(760,35)
$form.Controls.Add($desc)

$list = New-Object System.Windows.Forms.ListBox
$list.Location = New-Object System.Drawing.Point(20,105)
$list.Size = New-Object System.Drawing.Size(760,250)
$list.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($list)

$backups = Get-Backups

foreach ($b in $backups) {
    [void]$list.Items.Add(("{0} | {1:N2} MB | {2}" -f $b.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), ($b.Length / 1MB), $b.FullName))
}

$refresh = New-Object System.Windows.Forms.Button
$refresh.Text = "Actualizar lista"
$refresh.Location = New-Object System.Drawing.Point(20,380)
$refresh.Size = New-Object System.Drawing.Size(150,35)
$refresh.Add_Click({
    $list.Items.Clear()
    $script:backups = Get-Backups

    foreach ($b in $script:backups) {
        [void]$list.Items.Add(("{0} | {1:N2} MB | {2}" -f $b.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), ($b.Length / 1MB), $b.FullName))
    }
})
$form.Controls.Add($refresh)

$openFolder = New-Object System.Windows.Forms.Button
$openFolder.Text = "Abrir carpeta backups"
$openFolder.Location = New-Object System.Drawing.Point(190,380)
$openFolder.Size = New-Object System.Drawing.Size(170,35)
$openFolder.Add_Click({
    if (Test-Path $BackupRoot) {
        Start-Process explorer.exe $BackupRoot
    }
})
$form.Controls.Add($openFolder)

$restore = New-Object System.Windows.Forms.Button
$restore.Text = "Restaurar seleccionado"
$restore.Location = New-Object System.Drawing.Point(430,380)
$restore.Size = New-Object System.Drawing.Size(170,35)
$restore.Add_Click({
    if ($list.SelectedIndex -lt 0) {
        Err "Selecciona un backup primero."
        return
    }

    $selected = $backups[$list.SelectedIndex]
    Restore-Backup $selected.FullName
})
$form.Controls.Add($restore)

$close = New-Object System.Windows.Forms.Button
$close.Text = "Cerrar"
$close.Location = New-Object System.Drawing.Point(630,380)
$close.Size = New-Object System.Drawing.Size(150,35)
$close.Add_Click({ $form.Close() })
$form.Controls.Add($close)

[void]$form.ShowDialog()
