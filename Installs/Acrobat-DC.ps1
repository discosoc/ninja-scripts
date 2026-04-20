# ==============================================================================
# Acrobat-DC.ps1
# Installs Adobe Acrobat DC (64-bit) on Windows.
#
# Ninja script variables (set in NinjaRMM, passed as env vars):
#   installtype - standard (default) or force
#
# Downloaded files are retained under $workingDir.
# TODO: Future - MD5 hash check to skip redundant downloads
# ==============================================================================

# --- Variables ---
$workingDir      = "C:\Scripts"
$downloadUri     = "https://trials.adobe.com/AdobeProducts/APRO/Acrobat_HelpX/win32/Acrobat_DC_Web_x64_WWMUI.zip"
$outFile         = "$workingDir\Acrobat_DC_Web_x64_WWMUI.zip"
$installExe      = "$workingDir\Adobe Acrobat\setup.exe"
$detectionName   = "*Adobe Acrobat*"
$registryPaths   = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# --- Already-installed check ---
$installedEntry = $registryPaths |
    ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue } |
    Where-Object { $_.DisplayName -like $detectionName } |
    Select-Object -First 1

if ($installedEntry -and $installtype -ne 'force') {
    Write-Host "Adobe Acrobat DC is already installed (version $($installedEntry.DisplayVersion)). Set installtype to force to reinstall."
    exit 0
}

if ($installedEntry -and $installtype -eq 'force') {
    Write-Host "force — reinstalling Adobe Acrobat DC (currently $($installedEntry.DisplayVersion))."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Download ---
$ProgressPreference = 'SilentlyContinue'
Write-Host "Downloading Acrobat DC..."
Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
Write-Host "Download complete."

# --- Extract ---
Write-Host "Extracting archive..."
Expand-Archive -LiteralPath $outFile -DestinationPath $workingDir -Force
Write-Host "Extraction complete."

# --- Install ---
Write-Host "Installing Acrobat DC..."
Start-Process -Wait -FilePath $installExe -ArgumentList '/sAll /rs /msi EULA_ACCEPT=YES'
Write-Host "Acrobat DC installed successfully."
