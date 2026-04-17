# ==============================================================================
# App Install — NinjaRMM bootstrap
# Sets script variables and downloads the full install script from the repo.
#
# Ninja script variables:
#   appname       - specify the app to be installed
#   forceinstall  - bypass already-installed check when set to any value
# ==============================================================================

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$appname      = $env:appname
$forceinstall = $env:forceinstall
$scriptUri    = "https://raw.githubusercontent.com/discosoc/ninja-scripts/refs/heads/main/Installs/$appname.ps1"

$response = Invoke-WebRequest -Uri $scriptUri -UseBasicParsing
if ($response.StatusCode -ne 200) {
    Write-Host "ERROR: appname script not found (HTTP $($response.StatusCode))."
    exit 1
}
Invoke-Expression $response.Content