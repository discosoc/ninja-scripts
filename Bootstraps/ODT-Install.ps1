# ==============================================================================
# ODT-Install.ps1 — NinjaRMM bootstrap
# Sets script variables and downloads the Office install script from the repo.
#
# Ninja script variables:
#   edition        - Office bitness: 32 or 64 (required)
#   includeproject - include Project Pro when set to any value
#   includevisio   - include Visio Pro when set to any value
# ==============================================================================

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$edition        = $env:edition
$includeproject = $env:includeproject
$includevisio   = $env:includevisio
$scriptUri      = "https://raw.githubusercontent.com/discosoc/ninja-scripts/refs/heads/main/Installs/Office-Install.ps1"

$response = Invoke-WebRequest -Uri $scriptUri -UseBasicParsing
if ($response.StatusCode -ne 200) {
    Write-Host "ERROR: Office install script not found (HTTP $($response.StatusCode))."
    exit 1
}
Invoke-Expression $response.Content
