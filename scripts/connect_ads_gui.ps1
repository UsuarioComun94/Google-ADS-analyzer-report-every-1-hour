Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class TextBoxCueBanner {
    public const int EM_SETCUEBANNER = 0x1501;

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
}
"@

function Set-Placeholder($TextBox, $Text, $IsPassword) {
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


$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"
$SecretsDir = Join-Path $BaseDir "secrets"
$GoogleDir = Join-Path $BaseDir "connectors\google_ads"
$MetaDir = Join-Path $BaseDir "connectors\meta_ads"
$RawGoogleDir = Join-Path $BaseDir "raw_exports\google_ads"
$RawMetaDir = Join-Path $BaseDir "raw_exports\meta_ads"
$LogDir = Join-Path $BaseDir "logs"

New-Item -ItemType Directory -Force -Path $SecretsDir,$GoogleDir,$MetaDir,$RawGoogleDir,$RawMetaDir,$LogDir | Out-Null

function Show-Info($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Connect Ads", "OK", "Information") | Out-Null
}

function Show-ErrorBox($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Connect Ads - Error", "OK", "Error") | Out-Null
}

function Run-Cmd($file, $arguments) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $file
    $psi.Arguments = $arguments
    $psi.WorkingDirectory = $BaseDir
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
        "/secrets/",
        "/raw_exports/"
    )

    foreach ($line in $lines) {
        if ($content -notmatch [regex]::Escape($line)) {
            Add-Content -Path $gitignore -Value $line -Encoding UTF8
        }
    }
}

function Select-Platform {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "HMA Connect Ads"
    $form.Size = New-Object System.Drawing.Size(420,220)
    $form.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Elegí la plataforma a conectar:"
    $label.Location = New-Object System.Drawing.Point(20,25)
    $label.Size = New-Object System.Drawing.Size(360,25)
    $form.Controls.Add($label)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Location = New-Object System.Drawing.Point(20,60)
    $combo.Size = New-Object System.Drawing.Size(360,25)
    $combo.DropDownStyle = "DropDownList"
    [void]$combo.Items.Add("Google Ads")
    [void]$combo.Items.Add("Meta Ads")
    $combo.SelectedIndex = 0
    $form.Controls.Add($combo)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Continuar"
    $ok.Location = New-Object System.Drawing.Point(210,115)
    $ok.Size = New-Object System.Drawing.Size(80,30)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancelar"
    $cancel.Location = New-Object System.Drawing.Point(300,115)
    $cancel.Size = New-Object System.Drawing.Size(80,30)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)

    $form.AcceptButton = $ok
    $form.CancelButton = $cancel

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $combo.SelectedItem.ToString()
    }

    return $null
}

function Show-FieldsForm($title, $fields, $referenceText) {
    $height = 130 + ($fields.Count * 52)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(620,$height)
    $form.StartPosition = "CenterScreen"

    $controls = @{}
    $y = 20

    foreach ($field in $fields) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $field.Label
        $label.Location = New-Object System.Drawing.Point(20,$y)
        $label.Size = New-Object System.Drawing.Size(560,20)
        $form.Controls.Add($label)

        $box = New-Object System.Windows.Forms.TextBox
        $box.Location = New-Object System.Drawing.Point(20,($y + 22))
        $box.Size = New-Object System.Drawing.Size(560,22)

        if ($field.Password) {
            $box.UseSystemPasswordChar = $true
        }

        if ($field.Default) {
            $box.Text = $field.Default
        }

        $form.Controls.Add($box)

        if ($field.Placeholder) {
            Set-Placeholder $box $field.Placeholder $field.Password
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
    $ok.Location = New-Object System.Drawing.Point(360,($y + 10))
    $ok.Size = New-Object System.Drawing.Size(110,30)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancelar"
    $cancel.Location = New-Object System.Drawing.Point(480,($y + 10))
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

function Write-Google-Test {
    $path = Join-Path $GoogleDir "test_google_ads_connection.py"

    $code = @"
from pathlib import Path
from dotenv import load_dotenv
import os

from google.ads.googleads.client import GoogleAdsClient
from google.ads.googleads.errors import GoogleAdsException

BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_PATH = BASE_DIR / "secrets" / "google-ads.yaml"
ENV_PATH = BASE_DIR / "secrets" / "google_ads_account.env"

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

function Write-Meta-Test {
    $path = Join-Path $MetaDir "test_meta_ads_connection.py"

    $code = @"
from pathlib import Path
import os
import requests
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parents[2]
ENV_PATH = BASE_DIR / "secrets" / "meta_ads.env"

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

function Connect-Google {
    $fields = @(
        @{ Name="developer_token"; Label="Developer token de Google Ads API"; Required=$true; Password=$true; Default=""; Placeholder="Ejemplo: ABCDEF1234567890" },
        @{ Name="client_id"; Label="OAuth Client ID"; Required=$true; Password=$false; Default=""; Placeholder="Ejemplo: 1234567890-abcxyz.apps.googleusercontent.com" },
        @{ Name="client_secret"; Label="OAuth Client Secret"; Required=$true; Password=$true; Default=""; Placeholder="Ejemplo: GOCSPX-xxxxxxxxxxxxxxxx" },
        @{ Name="refresh_token"; Label="OAuth Refresh Token"; Required=$true; Password=$true; Default=""; Placeholder="Ejemplo: 1//0gxxxxxxxxxxxxxxxxxxxx" },
        @{ Name="login_customer_id"; Label="Login customer ID MCC sin guiones, opcional"; Required=$false; Password=$false; Default=""; Placeholder="Ejemplo: 1234567890" },
        @{ Name="customer_id"; Label="Customer ID cliente sin guiones"; Required=$true; Password=$false; Default=""; Placeholder="Ejemplo: 9876543210" }
    )

    $refs = "Google Ads referencias:
developer_token: se obtiene en Google Ads Manager > API Center.
client_id: suele terminar en .apps.googleusercontent.com.
client_secret: secreto OAuth del cliente.
refresh_token: token OAuth de actualización.
login_customer_id: ID del MCC sin guiones. Opcional si no usás MCC.
customer_id: ID de la cuenta cliente sin guiones."

    $data = Show-FieldsForm "Conectar Google Ads a HMA" $fields $refs
    if ($null -eq $data) { return }

    Ensure-Gitignore

    $yaml = @"
developer_token: "$($data.developer_token)"
client_id: "$($data.client_id)"
client_secret: "$($data.client_secret)"
refresh_token: "$($data.refresh_token)"
login_customer_id: "$($data.login_customer_id)"
use_proto_plus: True
"@

    Set-Content -Path (Join-Path $SecretsDir "google-ads.yaml") -Value $yaml -Encoding UTF8
    Set-Content -Path (Join-Path $SecretsDir "google_ads_account.env") -Value "GOOGLE_ADS_CUSTOMER_ID=$($data.customer_id)" -Encoding UTF8

    $pip = Run-Cmd $Python "-m pip install google-ads python-dotenv"
    if ($pip.ExitCode -ne 0) {
        Show-ErrorBox "Falló pip install.`n`n$($pip.StdErr)"
        return
    }

    $testPath = Write-Google-Test
    $test = Run-Cmd $Python "`"$testPath`""

    $log = Join-Path $LogDir ("connect_ads_google_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".log")
    Set-Content -Path $log -Value ($test.StdOut + "`n" + $test.StdErr) -Encoding UTF8

    if ($test.ExitCode -eq 0 -and $test.StdOut -match "OK_GOOGLE_ADS_CONNECTION|GOOGLE_ADS_CONNECTION_OK_BUT_NO_CUSTOMER_ROW") {
        Show-Info "Google Ads conectado correctamente.`n`nLog:`n$log"
    } else {
        Show-ErrorBox "Falló test Google Ads.`n`nLog:`n$log`n`n$($test.StdOut)`n$($test.StdErr)"
    }
}

function Connect-Meta {
    $fields = @(
        @{ Name="access_token"; Label="Meta access token"; Required=$true; Password=$true; Default=""; Placeholder="Ejemplo: EAABsbCS1iHgBOxxxxxxxxxxxxxxxx" },
        @{ Name="ad_account_id"; Label="Meta ad account ID"; Required=$true; Password=$false; Default=""; Placeholder="Ejemplo: act_123456789012345 o 123456789012345" },
        @{ Name="api_version"; Label="Meta API version"; Required=$true; Password=$false; Default="v21.0"; Placeholder="Ejemplo: v21.0" }
    )

    $refs = "Meta Ads referencias:
access_token: token con permiso ads_read.
ad_account_id: formato act_123456789.
api_version: ejemplo v21.0.
Para lectura de reportes alcanza ads_read."

    $data = Show-FieldsForm "Conectar Meta Ads a HMA" $fields $refs
    if ($null -eq $data) { return }

    Ensure-Gitignore

    $adAccount = $data.ad_account_id
    if ($adAccount -notmatch "^act_") {
        $adAccount = "act_$adAccount"
    }

    $env = @"
META_ACCESS_TOKEN=$($data.access_token)
META_AD_ACCOUNT_ID=$adAccount
META_API_VERSION=$($data.api_version)
"@

    Set-Content -Path (Join-Path $SecretsDir "meta_ads.env") -Value $env -Encoding UTF8

    $pip = Run-Cmd $Python "-m pip install requests python-dotenv"
    if ($pip.ExitCode -ne 0) {
        Show-ErrorBox "Falló pip install.`n`n$($pip.StdErr)"
        return
    }

    $testPath = Write-Meta-Test
    $test = Run-Cmd $Python "`"$testPath`""

    $log = Join-Path $LogDir ("connect_ads_meta_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".log")
    Set-Content -Path $log -Value ($test.StdOut + "`n" + $test.StdErr) -Encoding UTF8

    if ($test.ExitCode -eq 0 -and $test.StdOut -match "OK_META_ADS_CONNECTION") {
        Show-Info "Meta Ads conectado correctamente.`n`nLog:`n$log"
    } else {
        Show-ErrorBox "Falló test Meta Ads.`n`nLog:`n$log`n`n$($test.StdOut)`n$($test.StdErr)"
    }
}

try {
    if (-not (Test-Path $Python)) {
        Show-ErrorBox "No existe Python del entorno virtual:`n$Python"
        exit 1
    }

    $platform = Select-Platform
    if ($null -eq $platform) {
        exit 0
    }

    if ($platform -eq "Google Ads") {
        Connect-Google
    } elseif ($platform -eq "Meta Ads") {
        Connect-Meta
    }
} catch {
    Show-ErrorBox $_.Exception.Message
    exit 1
}
