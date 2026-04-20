# ==============================================================================
# Google-Earth-Pro.ps1
# Installs Google Earth Pro (64-bit) on Windows.
#
# Ninja script variables:
#   installtype - standard (default), upgrade, or force
# ==============================================================================

# --- Variables ---
$ProgressPreference = 'SilentlyContinue'
$workingDir    = "C:\Scripts"
$downloadUri   = "https://dl.google.com/dl/earth/client/advanced/current/googleearthprowin-7.3.7-x64.exe"
$outFile       = "$workingDir\googleearthprowin-7.3.7-x64.exe"
$detectionName = "*Google Earth Pro*"
$registryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# --- Already-installed check ---
$installedEntry = $registryPaths |
    ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue } |
    Where-Object { $_.DisplayName -like $detectionName } |
    Select-Object -First 1

if ($installtype -eq 'upgrade' -and -not $installedEntry) {
    Write-Host "Google Earth Pro is not installed — nothing to upgrade."
    exit 0
} elseif ($installedEntry -and $installtype -ne 'force' -and $installtype -ne 'upgrade') {
    Write-Host "Google Earth Pro is already installed (version $($installedEntry.DisplayVersion)). Set installtype to force to reinstall."
    exit 0
} elseif ($installedEntry -and $installtype -eq 'force') {
    Write-Host "force — reinstalling Google Earth Pro (currently $($installedEntry.DisplayVersion))."
} elseif ($installedEntry -and $installtype -eq 'upgrade') {
    Write-Host "upgrade — upgrading Google Earth Pro (currently $($installedEntry.DisplayVersion))."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Download ---
Write-Host "Downloading Google Earth Pro..."
Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
Write-Host "Download complete."

# --- Install ---
Write-Host "Installing Google Earth Pro..."
Start-Process -Wait -FilePath $outFile -ArgumentList "OMAHA=1"
Write-Host "Google Earth Pro installed successfully."
