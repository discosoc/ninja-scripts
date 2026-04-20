# ==============================================================================
# VLC.ps1
# Installs VLC media player (64-bit) on Windows.
#
# Ninja script variables:
#   installtype - standard (default), upgrade, or force
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

if ($installtype -eq 'upgrade' -and -not $installedEntry) {
    Write-Host "VLC media player is not installed — nothing to upgrade."
    exit 0
} elseif ($installedEntry -and $installtype -ne 'force' -and $installtype -ne 'upgrade') {
    Write-Host "VLC media player is already installed (version $($installedEntry.DisplayVersion)). Set installtype to force to reinstall."
    exit 0
} elseif ($installedEntry -and $installtype -eq 'force') {
    Write-Host "force — reinstalling VLC media player (currently $($installedEntry.DisplayVersion))."
} elseif ($installedEntry -and $installtype -eq 'upgrade') {
    Write-Host "upgrade — upgrading VLC media player (currently $($installedEntry.DisplayVersion))."
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
