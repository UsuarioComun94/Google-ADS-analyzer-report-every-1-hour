Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"
$ClientsRoot = Join-Path $BaseDir "clientes"
$LogDir = Join-Path $BaseDir "logs"

New-Item -ItemType Directory -Force -Path $ClientsRoot,$LogDir | Out-Null

function Show-Info($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Connect Ads", "OK", "Information") | Out-Null
}

function Show-ErrorBox($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Connect Ads - Error", "OK", "Error") | Out-Null
}

function Run-Cmd($file, $arguments, $workingDir) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $file
    $psi.Arguments = $arguments
    $psi.WorkingDirectory = $workingDir
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    $p.Start() | Out-Null
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    return @{
        ExitCode = $p.ExitCode
        StdOut = $stdout
        StdErr = $stderr
    }
}

function Ensure-Gitignore {
    $gitignore = Join-Path $BaseDir ".gitignore"
    if (-not (Test-Path $gitignore)) {
        New-Item -ItemType File -Path $gitignore | Out-Null
    }

    $content = Get-Content $gitignore -Raw

    $lines = @(
        "/clientes/*/secrets/",
        "/clientes/*/raw_exports/",
        "/clientes/*/logs/",
        "/clientes/*/error/",
        "/clientes/*/downloads/",
        "/clientes/*/historico/*.xlsx",
        "/clientes/*/historico/*.json"
    )

    foreach ($line in $lines) {
        if ($content -notmatch [regex]::Escape($line)) {
            Add-Content -Path $gitignore -Value $line -Encoding UTF8
        }
    }
}

function Get-HmaClients {
    if (-not (Test-Path $ClientsRoot)) {
        return @()
    }

    $clients = @()

    Get-ChildItem $ClientsRoot -Directory |
    Where-Object { $_.Name -ne "_template" } |
    ForEach-Object {
        $configPath = Join-Path $_.FullName "config\client_config.json"

        if (Test-Path $configPath) {
            try {
                $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
                $clients += [PSCustomObject]@{
                    Id = $cfg.client_id
                    Name = $cfg.client_name
                    Path = $_.FullName
                    Label = "$($cfg.client_name) [$($cfg.client_id)]"
                }
            } catch {}
        }
    }

    return $clients
}

function Select-ClientAndPlatform {
    $clients = @(Get-HmaClients)

    if ($clients.Count -eq 0) {
        Show-ErrorBox "No hay clientes creados. Primero ejecutá:`n`n.\scripts\create_hma_client.ps1 -ClientId `"cliente_001`" -ClientName `"Vics Solutions`""
        return $null
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "HMA Connect Ads"
    $form.Size = New-Object System.Drawing.Size(520,280)
    $form.StartPosition = "CenterScreen"

    $clientLabel = New-Object System.Windows.Forms.Label
    $clientLabel.Text = "Elegí el cliente:"
    $clientLabel.Location = New-Object System.Drawing.Point(20,25)
    $clientLabel.Size = New-Object System.Drawing.Size(460,22)
    $form.Controls.Add($clientLabel)

    $clientCombo = New-Object System.Windows.Forms.ComboBox
    $clientCombo.Location = New-Object System.Drawing.Point(20,50)
    $clientCombo.Size = New-Object System.Drawing.Size(460,25)
    $clientCombo.DropDownStyle = "DropDownList"

    foreach ($c in $clients) {
        [void]$clientCombo.Items.Add($c.Label)
    }

    $clientCombo.SelectedIndex = 0
    $form.Controls.Add($clientCombo)

    $platformLabel = New-Object System.Windows.Forms.Label
    $platformLabel.Text = "Elegí la plataforma:"
    $platformLabel.Location = New-Object System.Drawing.Point(20,95)
    $platformLabel.Size = New-Object System.Drawing.Size(460,22)
    $form.Controls.Add($platformLabel)

    $platformCombo = New-Object System.Windows.Forms.ComboBox
    $platformCombo.Location = New-Object System.Drawing.Point(20,120)
    $platformCombo.Size = New-Object System.Drawing.Size(460,25)
    $platformCombo.DropDownStyle = "DropDownList"
    [void]$platformCombo.Items.Add("Google Ads")
    [void]$platformCombo.Items.Add("Meta Ads")
    $platformCombo.SelectedIndex = 0
    $form.Controls.Add($platformCombo)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Continuar"
    $ok.Location = New-Object System.Drawing.Point(300,185)
    $ok.Size = New-Object System.Drawing.Size(85,30)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancelar"
    $cancel.Location = New-Object System.Drawing.Point(395,185)
    $cancel.Size = New-Object System.Drawing.Size(85,30)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)

    $form.AcceptButton = $ok
    $form.CancelButton = $cancel

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    $selectedClient = $clients[$clientCombo.SelectedIndex]

    return [PSCustomObject]@{
        Client = $selectedClient
        Platform = $platformCombo.SelectedItem.ToString()
    }
}

function Set-VisiblePlaceholder($TextBox, $Text, $IsPassword) {
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    $TextBox.AccessibleName = $Text
    $TextBox.AccessibleDescription = if ($IsPassword) { "password" } else { "" }

    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) {
        $TextBox.UseSystemPasswordChar = $false
        $TextBox.ForeColor = [System.Drawing.Color]::Gray
        $TextBox.Text = $Text
        $TextBox.Tag = "placeholder"
    }

    $TextBox.Add_Enter({
        if ($this.Tag -eq "placeholder") {
            $this.Text = ""
            $this.ForeColor = [System.Drawing.Color]::Black
            $this.Tag = ""
            if ($this.AccessibleDescription -eq "password") {
                $this.UseSystemPasswordChar = $true
            }
        }
    })

    $TextBox.Add_Leave({
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $this.UseSystemPasswordChar = $false
            $this.ForeColor = [System.Drawing.Color]::Gray
            $this.Text = $this.AccessibleName
            $this.Tag = "placeholder"
        }
    })
}

function Show-FieldsForm($title, $fields, $referenceText) {
    $height = 130 + ($fields.Count * 52)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(660,$height)
    $form.StartPosition = "CenterScreen"

    $controls = @{}
    $y = 20

    foreach ($field in $fields) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $field.Label
        $label.Location = New-Object System.Drawing.Point(20,$y)
        $label.Size = New-Object System.Drawing.Size(600,20)
        $form.Controls.Add($label)

        $box = New-Object System.Windows.Forms.TextBox
        $box.Location = New-Object System.Drawing.Point(20,($y + 22))
        $box.Size = New-Object System.Drawing.Size(600,22)

        if ($field.Default) {
            $box.Text = $field.Default
        }

        $form.Controls.Add($box)

        if ($field.Placeholder) {
            Set-VisiblePlaceholder $box $field.Placeholder $field.Password
        } elseif ($field.Password) {
            $box.UseSystemPasswordChar = $true
        }

        $controls[$field.Name] = $box
        $y += 52
    }

    $refs = New-Object System.Windows.Forms.Button
    $refs.Text = "Ver referencias"
    $refs.Location = New-Object System.Drawing.Point(20,($y + 10))
    $refs.Size = New-Object System.Drawing.Size(120,30)
    $refs.Add_Click({ Show-Info $referenceText })
    $form.Controls.Add($refs)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Guardar y probar"
    $ok.Location = New-Object System.Drawing.Point(390,($y + 10))
    $ok.Size = New-Object System.Drawing.Size(120,30)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancelar"
    $cancel.Location = New-Object System.Drawing.Point(520,($y + 10))
    $cancel.Size = New-Object System.Drawing.Size(100,30)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)

    $form.AcceptButton = $ok
    $form.CancelButton = $cancel

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    $result = @{}

    foreach ($field in $fields) {
        $value = $controls[$field.Name].Text.Trim()

        if ($controls[$field.Name].Tag -eq "placeholder") {
            $value = ""
        }

        if ($field.Required -and [string]::IsNullOrWhiteSpace($value)) {
            Show-ErrorBox "Falta completar: $($field.Label)"
            return $null
        }

        $result[$field.Name] = $value
    }

    return $result
}

function Update-Client-Config($ClientPath, $Platform, $Data) {
    $configPath = Join-Path $ClientPath "config\client_config.json"
    $cfg = Get-Content $configPath -Raw | ConvertFrom-Json

    if ($Platform -eq "google") {
        $cfg.platforms.google_ads.enabled = $true
        $cfg.platforms.google_ads.customer_id = $Data.customer_id
        $cfg.platforms.google_ads.login_customer_id = $Data.login_customer_id
    }

    if ($Platform -eq "meta") {
        $cfg.platforms.meta_ads.enabled = $true
        $cfg.platforms.meta_ads.ad_account_id = $Data.ad_account_id
    }

    $cfg | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8
}

function Write-Google-Test($ClientPath) {
    $dir = Join-Path $ClientPath "google_ads"
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $path = Join-Path $dir "test_google_ads_connection.py"

    $code = @"
from pathlib import Path
from dotenv import load_dotenv
import os

from google.ads.googleads.client import GoogleAdsClient
from google.ads.googleads.errors import GoogleAdsException

CLIENT_DIR = Path(__file__).resolve().parents[1]
CONFIG_PATH = CLIENT_DIR / "secrets" / "google-ads.yaml"
ENV_PATH = CLIENT_DIR / "secrets" / "google_ads_account.env"

load_dotenv(ENV_PATH)

CUSTOMER_ID = os.getenv("GOOGLE_ADS_CUSTOMER_ID")

def main():
    if not CUSTOMER_ID:
        raise RuntimeError("Falta GOOGLE_ADS_CUSTOMER_ID")

    client = GoogleAdsClient.load_from_storage(str(CONFIG_PATH))
    ga_service = client.get_service("GoogleAdsService")

    query = """
        SELECT
          customer.id,
          customer.descriptive_name,
          customer.currency_code,
          customer.time_zone
        FROM customer
        LIMIT 1
    """

    try:
        response = ga_service.search(customer_id=CUSTOMER_ID, query=query)
        found = False
        for row in response:
            found = True
            print("OK_GOOGLE_ADS_CONNECTION")
            print("customer_id:", row.customer.id)
            print("name:", row.customer.descriptive_name)
            print("currency:", row.customer.currency_code)
            print("timezone:", row.customer.time_zone)

        if not found:
            print("GOOGLE_ADS_CONNECTION_OK_BUT_NO_CUSTOMER_ROW")

    except GoogleAdsException as ex:
        print("GOOGLE_ADS_API_ERROR")
        print(ex)
        raise

if __name__ == "__main__":
    main()
"@

    Set-Content -Path $path -Value $code -Encoding UTF8
    return $path
}

function Write-Meta-Test($ClientPath) {
    $dir = Join-Path $ClientPath "meta_ads"
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $path = Join-Path $dir "test_meta_ads_connection.py"

    $code = @"
from pathlib import Path
import os
import requests
from dotenv import load_dotenv

CLIENT_DIR = Path(__file__).resolve().parents[1]
ENV_PATH = CLIENT_DIR / "secrets" / "meta_ads.env"

load_dotenv(ENV_PATH)

ACCESS_TOKEN = os.getenv("META_ACCESS_TOKEN")
AD_ACCOUNT_ID = os.getenv("META_AD_ACCOUNT_ID")
API_VERSION = os.getenv("META_API_VERSION", "v21.0")

def main():
    if not ACCESS_TOKEN:
        raise RuntimeError("Falta META_ACCESS_TOKEN")
    if not AD_ACCOUNT_ID:
        raise RuntimeError("Falta META_AD_ACCOUNT_ID")

    url = f"https://graph.facebook.com/{API_VERSION}/{AD_ACCOUNT_ID}"
    params = {
        "fields": "id,name,account_status,currency,timezone_name",
        "access_token": ACCESS_TOKEN,
    }

    r = requests.get(url, params=params, timeout=30)
    print("status_code:", r.status_code)
    print(r.text)
    r.raise_for_status()
    print("OK_META_ADS_CONNECTION")

if __name__ == "__main__":
    main()
"@

    Set-Content -Path $path -Value $code -Encoding UTF8
    return $path
}

function Connect-Google($Client) {
    $clientPath = $Client.Path

    $fields = @(
        @{ Name="developer_token"; Label="Developer token de Google Ads API"; Required=$true; Password=$true; Default=""; Placeholder="Ejemplo: ABCDEF1234567890" },
        @{ Name="client_id"; Label="OAuth Client ID"; Required=$true; Password=$false; Default=""; Placeholder="Ejemplo: 1234567890-abcxyz.apps.googleusercontent.com" },
        @{ Name="client_secret"; Label="OAuth Client Secret"; Required=$true; Password=$true; Default=""; Placeholder="Ejemplo: GOCSPX-xxxxxxxxxxxxxxxx" },
        @{ Name="refresh_token"; Label="OAuth Refresh Token"; Required=$true; Password=$true; Default=""; Placeholder="Ejemplo: 1//0gxxxxxxxxxxxxxxxxxxxx" },
        @{ Name="login_customer_id"; Label="Login customer ID MCC sin guiones, opcional"; Required=$false; Password=$false; Default=""; Placeholder="Ejemplo: 1234567890" },
        @{ Name="customer_id"; Label="Customer ID cliente sin guiones"; Required=$true; Password=$false; Default=""; Placeholder="Ejemplo: 9876543210" }
    )

    $refs = "Google Ads referencias:
developer_token: Google Ads Manager > API Center.
client_id: suele terminar en .apps.googleusercontent.com.
client_secret: secreto OAuth.
refresh_token: token OAuth.
login_customer_id: ID del MCC sin guiones. Opcional.
customer_id: ID de la cuenta cliente sin guiones."

    $data = Show-FieldsForm "Conectar Google Ads — $($Client.Label)" $fields $refs
    if ($null -eq $data) { return }

    Ensure-Gitignore

    $secretsDir = Join-Path $clientPath "secrets"
    New-Item -ItemType Directory -Force -Path $secretsDir | Out-Null

    $yaml = @"
developer_token: "$($data.developer_token)"
client_id: "$($data.client_id)"
client_secret: "$($data.client_secret)"
refresh_token: "$($data.refresh_token)"
login_customer_id: "$($data.login_customer_id)"
use_proto_plus: True
"@

    Set-Content -Path (Join-Path $secretsDir "google-ads.yaml") -Value $yaml -Encoding UTF8
    Set-Content -Path (Join-Path $secretsDir "google_ads_account.env") -Value "GOOGLE_ADS_CUSTOMER_ID=$($data.customer_id)" -Encoding UTF8

    Update-Client-Config $clientPath "google" $data

    $pip = Run-Cmd $Python "-m pip install google-ads python-dotenv" $BaseDir
    if ($pip.ExitCode -ne 0) {
        Show-ErrorBox "Falló pip install.`n`n$($pip.StdErr)"
        return
    }

    $testPath = Write-Google-Test $clientPath
    $test = Run-Cmd $Python "`"$testPath`"" $BaseDir

    $log = Join-Path $LogDir ("connect_ads_google_" + $Client.Id + "_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".log")
    Set-Content -Path $log -Value ($test.StdOut + "`n" + $test.StdErr) -Encoding UTF8

    if ($test.ExitCode -eq 0 -and $test.StdOut -match "OK_GOOGLE_ADS_CONNECTION|GOOGLE_ADS_CONNECTION_OK_BUT_NO_CUSTOMER_ROW") {
        Show-Info "Google Ads conectado correctamente para $($Client.Label).`n`nLog:`n$log"
    } else {
        Show-ErrorBox "Falló test Google Ads para $($Client.Label).`n`nLog:`n$log`n`n$($test.StdOut)`n$($test.StdErr)"
    }
}

function Connect-Meta($Client) {
    $clientPath = $Client.Path

    $fields = @(
        @{ Name="access_token"; Label="Meta access token"; Required=$true; Password=$true; Default=""; Placeholder="Ejemplo: EAABsbCS1iHgBOxxxxxxxxxxxxxxxx" },
        @{ Name="ad_account_id"; Label="Meta ad account ID"; Required=$true; Password=$false; Default=""; Placeholder="Ejemplo: act_123456789012345 o 123456789012345" },
        @{ Name="api_version"; Label="Meta API version"; Required=$true; Password=$false; Default="v21.0"; Placeholder="Ejemplo: v21.0" }
    )

    $refs = "Meta Ads referencias:
access_token: token con permiso ads_read.
ad_account_id: formato act_123456789 o solo número.
api_version: ejemplo v21.0."

    $data = Show-FieldsForm "Conectar Meta Ads — $($Client.Label)" $fields $refs
    if ($null -eq $data) { return }

    Ensure-Gitignore

    $secretsDir = Join-Path $clientPath "secrets"
    New-Item -ItemType Directory -Force -Path $secretsDir | Out-Null

    $adAccount = $data.ad_account_id
    if ($adAccount -notmatch "^act_") {
        $adAccount = "act_$adAccount"
    }

    $env = @"
META_ACCESS_TOKEN=$($data.access_token)
META_AD_ACCOUNT_ID=$adAccount
META_API_VERSION=$($data.api_version)
"@

    Set-Content -Path (Join-Path $secretsDir "meta_ads.env") -Value $env -Encoding UTF8

    $data.ad_account_id = $adAccount
    Update-Client-Config $clientPath "meta" $data

    $pip = Run-Cmd $Python "-m pip install requests python-dotenv" $BaseDir
    if ($pip.ExitCode -ne 0) {
        Show-ErrorBox "Falló pip install.`n`n$($pip.StdErr)"
        return
    }

    $testPath = Write-Meta-Test $clientPath
    $test = Run-Cmd $Python "`"$testPath`"" $BaseDir

    $log = Join-Path $LogDir ("connect_ads_meta_" + $Client.Id + "_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".log")
    Set-Content -Path $log -Value ($test.StdOut + "`n" + $test.StdErr) -Encoding UTF8

    if ($test.ExitCode -eq 0 -and $test.StdOut -match "OK_META_ADS_CONNECTION") {
        Show-Info "Meta Ads conectado correctamente para $($Client.Label).`n`nLog:`n$log"
    } else {
        Show-ErrorBox "Falló test Meta Ads para $($Client.Label).`n`nLog:`n$log`n`n$($test.StdOut)`n$($test.StdErr)"
    }
}

try {
    if (-not (Test-Path $Python)) {
        Show-ErrorBox "No existe Python del entorno virtual:`n$Python"
        exit 1
    }

    $selection = Select-ClientAndPlatform
    if ($null -eq $selection) {
        exit 0
    }

    if ($selection.Platform -eq "Google Ads") {
        Connect-Google $selection.Client
    } elseif ($selection.Platform -eq "Meta Ads") {
        Connect-Meta $selection.Client
    }
} catch {
    Show-ErrorBox $_.Exception.Message
    exit 1
}
