# ==============================================================================
# HEIF.ps1
# Installs the HEIF Image Extension (Microsoft.HEIFImageExtension) on Windows.
#
# Ninja script variables:
#   installtype - standard (default), upgrade, or force
# ==============================================================================

# --- Variables ---
$ProgressPreference = 'SilentlyContinue'
$workingDir  = "C:\Scripts"
$packageName = "HEIFImageExtension"
$downloadUri = "http://tlu.dl.delivery.mp.microsoft.com/filestreamingservice/files/ce4a7a9a-bad2-4676-87f4-368e08ed2cf9?P1=1778195850&P2=404&P3=2&P4=TFi9isY3vpLdKhCLcFyCdwkiQRi6hIaVqWTujoKiVnvAxRcmzC3qHmuNXVUsr2w7Hv5sZGDnMP6ogzKDdO%2bFCg%3d%3d"
$outFile     = "$workingDir\Microsoft.HEIFImageExtension_1.2.30.0_neutral_~_8wekyb3d8bbwe.AppxBundle"

# --- Already-installed check ---
$installedPkg = Get-AppxProvisionedPackage -Online |
    Where-Object { $_.PackageName -like "*$packageName*" } |
    Select-Object -First 1

if ($installtype -eq 'upgrade' -and -not $installedPkg) {
    Write-Host "HEIF Image Extension is not installed -- nothing to upgrade."
    exit 0
} elseif ($installedPkg -and $installtype -ne 'force' -and $installtype -ne 'upgrade') {
    Write-Host "HEIF Image Extension is already installed (version $($installedPkg.Version)). Set installtype to force to reinstall."
    exit 0
} elseif ($installedPkg -and $installtype -eq 'force') {
    Write-Host "force -- reinstalling HEIF Image Extension (currently $($installedPkg.Version))."
} elseif ($installedPkg -and $installtype -eq 'upgrade') {
    Write-Host "upgrade -- upgrading HEIF Image Extension (currently $($installedPkg.Version))."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Download ---
Write-Host "Downloading HEIF Image Extension..."
try {
    Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
} catch {
    Write-Host "ERROR: Failed to download HEIF Image Extension. $_"
    exit 1
}
Write-Host "Download complete."

# --- Install ---
Write-Host "Installing HEIF Image Extension..."
try {
    Add-AppxProvisionedPackage -Online -PackagePath $outFile -SkipLicense
} catch {
    Write-Host "ERROR: HEIF Image Extension installation failed. $_"
    exit 1
}
Write-Host "HEIF Image Extension installed successfully."
