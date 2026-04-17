# ==============================================================================
# ODT-Removal.ps1 — NinjaRMM bootstrap
# Downloads and runs the Office removal script from the repo.
# ==============================================================================

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$scriptUri = "https://raw.githubusercontent.com/discosoc/ninja-scripts/refs/heads/main/Installs/Office-Removal.ps1"

$response = Invoke-WebRequest -Uri $scriptUri -UseBasicParsing
if ($response.StatusCode -ne 200) {
    Write-Host "ERROR: Office removal script not found (HTTP $($response.StatusCode))."
    exit 1
}
Invoke-Expression $response.Content
