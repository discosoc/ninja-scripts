# ==============================================================================
# Bluebeam.ps1
# Installs Bluebeam Revu 21 (64-bit) on Windows.
#
# Ninja script variables (set in NinjaRMM, passed as env vars):
#   installtype - standard (default) or force
#
# Downloaded files are retained under $workingDir.
# TODO: Future - MD5 hash check to skip redundant downloads
# ==============================================================================

# --- Variables ---
$workingDir         = "C:\Scripts"
$redirectUri        = "https://bluebeam.com/MSIdeployx64"
$extractDir         = "$workingDir\MSIBluebeamRevu21"
$installMsi         = "$extractDir\Bluebeam Revu x64 21.msi"
$ocrMsi             = "$extractDir\BluebeamOCR x64 21.msi"
$dotNetExe          = "$extractDir\ndp48-x86-x64-allos-enu.exe"
$vcRedistExe        = "$extractDir\vc_redist.x64.exe"
$detectionName      = "*Bluebeam Revu*"
$registryPaths      = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$ProgressPreference = 'SilentlyContinue'

# --- Already-installed check ---
$installedEntry = $registryPaths |
    ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue } |
    Where-Object { $_.DisplayName -like $detectionName } |
    Select-Object -First 1

if ($installedEntry -and $installtype -ne 'force') {
    Write-Host "Bluebeam Revu is already installed (version $($installedEntry.DisplayVersion)). Set installtype to force to reinstall."
    exit 0
}

if ($installedEntry -and $installtype -eq 'force') {
    Write-Host "force — reinstalling Bluebeam Revu (currently $($installedEntry.DisplayVersion))."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Resolve download URL ---
Write-Host "Resolving download URL..."
try {
    $req = [System.Net.WebRequest]::Create($redirectUri)
    $req.Method = "HEAD"
    $req.AllowAutoRedirect = $true
    $resp = $req.GetResponse()
    $downloadUri = $resp.ResponseUri.AbsoluteUri
    $resp.Close()
} catch {
    Write-Host "Failed to resolve download URL: $_"
    exit 1
}

if (-not $downloadUri) {
    Write-Host "Resolved URL is empty. Cannot continue."
    exit 1
}

$outFile = "$workingDir\$(Split-Path $downloadUri -Leaf)"
Write-Host "Resolved: $downloadUri"

# --- Download ---
Write-Host "Downloading Bluebeam Revu deployment package..."
Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
Write-Host "Download complete."

# --- Extract ---
Write-Host "Extracting archive..."
Expand-Archive -LiteralPath $outFile -DestinationPath $workingDir -Force
Write-Host "Extraction complete."

# --- Prerequisites ---
$dotNetRelease = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release
if ($dotNetRelease -ge 528040) {
    Write-Host ".NET Framework 4.8 already installed."
} else {
    Write-Host "Installing .NET Framework 4.8..."
    Start-Process -Wait -FilePath $dotNetExe -ArgumentList '/quiet /norestart'
    Write-Host ".NET Framework 4.8 installed."
}

$vcInstalled = $registryPaths |
    ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue } |
    Where-Object { $_.DisplayName -like "*Visual C++ 2022*" } |
    Select-Object -First 1

if ($vcInstalled) {
    Write-Host "Visual C++ 2022 Redistributable already installed."
} else {
    Write-Host "Installing Visual C++ 2022 Redistributable..."
    Start-Process -Wait -FilePath $vcRedistExe -ArgumentList '/quiet /norestart'
    Write-Host "Visual C++ 2022 Redistributable installed."
}

# --- Install Bluebeam Revu ---
Write-Host "Installing Bluebeam Revu..."
Start-Process -Wait -FilePath "msiexec.exe" -ArgumentList "/i `"$installMsi`" /qn BB_AUTO_UPDATE=0 BB_DISABLEANALYTICS=1 BB_DEFAULTVIEWER=1"
Write-Host "Bluebeam Revu installed successfully."

# --- Install OCR Module ---
Write-Host "Installing Bluebeam OCR module..."
Start-Process -Wait -FilePath "msiexec.exe" -ArgumentList "/i `"$ocrMsi`" /qn"
Write-Host "Bluebeam OCR module installed successfully."
