# ==============================================================================
# PrinterLogic.ps1
# Installs the PrinterLogic Client on Windows.
#
# Ninja script variables:
#   forceinstall            - bypass already-installed check when set to any value
#   printerlogichomeurl    - Org custom field - PrinterLogic tenant URL (required)
#   printerlogicauthcode   - Org custom field - PrinterLogic authorization code (required)
# ==============================================================================

Get-ChildItem env: | Sort-Object Name | ForEach-Object { Write-Host "$($_.Name) = $($_.Value)" } 

write-host $env:uefiversion

# --- Ninja variables ---
$printerlogichomeurl   = $env:printerlogichomeurl
$printerlogicauthcode  = $env:printerlogicauthcode

if (-not $printerlogichomeurl) {
    Write-Host "ERROR: printerlogichomeurl org custom field is required."
    exit 1
}
if (-not $printerlogicauthcode) {
    Write-Host "ERROR: printerlogicauthcode org custom field is required."
    exit 1
}

# --- Variables ---
$ProgressPreference = 'SilentlyContinue'
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
Start-Process -Wait -FilePath "msiexec.exe" -ArgumentList "/i `"$outFile`" /qn HOMEURL=`"$printerlogichomeurl`" AUTHORIZATION_CODE=`"$printerlogicauthcode`""
Write-Host "PrinterLogic Client installed successfully."
