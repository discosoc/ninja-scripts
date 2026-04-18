# ==============================================================================
# 7-Zip.ps1
# Installs 7-Zip (64-bit) on Windows.
#
# Ninja script variables:
#   forceinstall - bypass already-installed check when set to any value
# ==============================================================================

# --- Variables ---
$ProgressPreference = 'SilentlyContinue'
$workingDir    = "C:\Scripts"
$outFile       = "$workingDir\7zip-x64.exe"
$detectionName = "*7-Zip*"
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
    Write-Host "7-Zip is already installed (version $($installedEntry.DisplayVersion)). Set forceinstall to override."
    exit 0
}

if ($installedEntry -and $forceinstall) {
    Write-Host "forceinstall set — reinstalling 7-Zip (currently $($installedEntry.DisplayVersion))."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Download ---
Write-Host "Resolving 7-Zip download link..."
try {
    $page = Invoke-WebRequest -Uri "https://www.7-zip.org/download.html" -UseBasicParsing
} catch {
    Write-Host "ERROR: Failed to reach 7-Zip download page. $_"
    exit 1
}
$downloadUri = ($page.Links | Where-Object { $_.href -like "*releases/download*x64.exe" } | Select-Object -First 1).href
if (-not $downloadUri) {
    Write-Host "ERROR: Could not locate 7-Zip download link on the download page."
    exit 1
}
Write-Host "Downloading 7-Zip..."
try {
    Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
} catch {
    Write-Host "ERROR: Failed to download 7-Zip. $_"
    exit 1
}
Write-Host "Download complete."

# --- Install ---
Write-Host "Installing 7-Zip..."
$result = Start-Process -Wait -PassThru -FilePath $outFile -ArgumentList "/S"
if ($result.ExitCode -ne 0) {
    Write-Host "ERROR: 7-Zip installation failed (exit code $($result.ExitCode))."
    exit 1
}
Write-Host "7-Zip installed successfully."
