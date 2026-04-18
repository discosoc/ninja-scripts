# ==============================================================================
# VLC.ps1
# Installs VLC media player (64-bit) on Windows.
#
# Ninja script variables:
#   forceinstall - bypass already-installed check when set to any value
# ==============================================================================

# --- Variables ---
$ProgressPreference = 'SilentlyContinue'
$workingDir    = "C:\Scripts"
$detectionName = "*VLC media player*"
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
    Write-Host "VLC media player is already installed (version $($installedEntry.DisplayVersion)). Set forceinstall to override."
    exit 0
}

if ($installedEntry -and $forceinstall) {
    Write-Host "forceinstall set — reinstalling VLC media player (currently $($installedEntry.DisplayVersion))."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Download ---
Write-Host "Resolving VLC download link..."
try {
    $page = Invoke-WebRequest -Uri "https://www.videolan.org/vlc/download-windows.html" -UseBasicParsing
} catch {
    Write-Host "ERROR: Failed to reach VLC download page. $_"
    exit 1
}
$downloadUri = ($page.Links | Where-Object { $_.href -like "*win64.msi" } | Select-Object -First 1).href
if (-not $downloadUri) {
    Write-Host "ERROR: Could not locate VLC download link on the download page."
    exit 1
}
$outFile = "$workingDir\$(Split-Path $downloadUri -Leaf)"
Write-Host "Downloading VLC media player..."
try {
    Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
} catch {
    Write-Host "ERROR: Failed to download VLC. $_"
    exit 1
}
Write-Host "Download complete."

# --- Install ---
Write-Host "Installing VLC media player..."
$result = Start-Process -Wait -PassThru -FilePath "msiexec.exe" -ArgumentList "/i `"$outFile`" /quiet /norestart"
if ($result.ExitCode -ne 0) {
    Write-Host "ERROR: VLC installation failed (exit code $($result.ExitCode))."
    exit 1
}
Write-Host "VLC media player installed successfully."
