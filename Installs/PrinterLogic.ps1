# ==============================================================================
# PrinterLogic.ps1
# Installs the PrinterLogic Client on Windows.
#
# Ninja script variables:
#   homeurl           - PrinterLogic tenant URL (required)
#   authorizationcode - PrinterLogic authorization code (required)
#   forceinstall      - bypass already-installed check when set to any value
# ==============================================================================

# --- Ninja variables ---
$homeurl           = $env:homeurl
$authorizationcode = $env:authorizationcode

if (-not $homeurl) {
    Write-Host "ERROR: homeurl script variable is required."
    exit 1
}
if (-not $authorizationcode) {
    Write-Host "ERROR: authorizationcode script variable is required."
    exit 1
}

# --- Variables ---
$workingDir    = "C:\Scripts"
$downloadUri   = "https://afc.printercloud.com/client/setup/PrinterInstallerClient.msi"
$outFile       = "$workingDir\PrinterInstallerClient.msi"
$detectionName = "*PrinterLogic*"
$registryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# --- Already-installed check ---
$installedEntry = $registryPaths |
    ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue } |
    Where-Object { $_.DisplayName -like $detectionName } |
    Select-Object -First 1

if ($installedEntry -and -not $forceinstall) {
    Write-Host "PrinterLogic is already installed (version $($installedEntry.DisplayVersion)). Set forceinstall to override."
    exit 0
}

if ($installedEntry -and $forceinstall) {
    Write-Host "forceinstall set — reinstalling PrinterLogic (currently $($installedEntry.DisplayVersion))."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Download ---
Write-Host "Downloading PrinterLogic Client..."
Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
Write-Host "Download complete."

# --- Install ---
Write-Host "Installing PrinterLogic Client..."
Start-Process -Wait -FilePath "msiexec.exe" -ArgumentList "/i `"$outFile`" /qn HOMEURL=`"$homeurl`" AUTHORIZATION_CODE=`"$authorizationcode`""
Write-Host "PrinterLogic Client installed successfully."
