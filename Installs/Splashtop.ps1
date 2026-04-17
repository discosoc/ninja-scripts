# ==============================================================================
# Splashtop.ps1
# Installs the Splashtop Streamer on Windows.
#
# Ninja script variables:
#   splashtopcode - Splashtop deployment package code (required)
#   forceinstall  - bypass already-installed check when set to any value
# ==============================================================================

# --- Ninja variables ---
$splashtopcode = $env:splashtopcode

if (-not $splashtopcode) {
    Write-Host "ERROR: splashtopcode script variable is required."
    exit 1
}

# --- Variables ---
$workingDir    = "C:\Scripts"
$downloadUri   = "https://my.splashtop.com/team_deployment/download_directly/win/5A7PPZR5J3JT"
$outFile       = "$workingDir\SplashtopStreamer.msi"
$detectionName = "*Splashtop Streamer*"
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
    Write-Host "Splashtop Streamer is already installed (version $($installedEntry.DisplayVersion)). Set forceinstall to override."
    exit 0
}

if ($installedEntry -and $forceinstall) {
    Write-Host "forceinstall set — reinstalling Splashtop Streamer (currently $($installedEntry.DisplayVersion))."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Download ---
Write-Host "Downloading Splashtop Streamer..."
Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
Write-Host "Download complete."

# --- Install ---
Write-Host "Installing Splashtop Streamer..."
Start-Process -Wait -FilePath "msiexec.exe" -ArgumentList "/i `"$outFile`" /quiet /norestart USERINFO=`"decode=$splashtopcode,hidewindow=1,confirm_d=0`""
Write-Host "Splashtop Streamer installed successfully."
