# ==============================================================================
# Logitech-Options.ps1
# Installs Logitech Options on Windows.
#
# Ninja script variables:
#   installtype - standard (default) or force
# ==============================================================================

# --- Variables ---
$ProgressPreference = 'SilentlyContinue'
$workingDir    = "C:\Scripts"
$downloadUri   = "https://download01.logi.com/web/ftp/pub/techsupport/options/options_installer.exe"
$outFile       = "$workingDir\options_installer.exe"
$detectionName = "*Logitech Options*"
$registryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# --- Already-installed check ---
# Exclude Options+ to avoid a false match on the detection pattern
$installedEntry = $registryPaths |
    ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue } |
    Where-Object { $_.DisplayName -like $detectionName -and $_.DisplayName -notlike "*+*" } |
    Select-Object -First 1

if ($installedEntry -and $installtype -ne 'force') {
    Write-Host "Logitech Options is already installed (version $($installedEntry.DisplayVersion)). Set installtype to force to reinstall."
    exit 0
}

if ($installedEntry -and $installtype -eq 'force') {
    Write-Host "force — reinstalling Logitech Options (currently $($installedEntry.DisplayVersion))."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Download ---
Write-Host "Downloading Logitech Options..."
Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
Write-Host "Download complete."

# --- Install ---
Write-Host "Installing Logitech Options..."
Start-Process -Wait -FilePath $outFile -ArgumentList "/quiet /analytics no"
Write-Host "Logitech Options installed successfully."
