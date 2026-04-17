# ==============================================================================
# Dell-Command-Update.ps1
# Installs Dell Command | Update for Windows Universal on Dell systems.
# Uninstalls the classic or Windows 10 variant first if present.
#
# Ninja script variables:
#   forceinstall - bypass already-installed check when set to any value
# ==============================================================================

# --- Variables ---
$workingDir     = "C:\Scripts"
$downloadUri    = "https://dl.dell.com/FOLDER12925773M/2/Dell-Command-Update-Windows-Universal-Application_P4DJW_WIN64_5.5.0_A00_01.EXE"
$outFile        = "$workingDir\Dell-Command-Update-Universal.exe"
$uwpDisplayName = "Dell Command | Update for Windows Universal"
$legacyNames    = @("Dell Command | Update", "Dell Command | Update for Windows 10")
$registryPaths  = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# --- Collect all relevant registry entries up front ---
$allEntries = $registryPaths |
    ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue } |
    Where-Object { $_.DisplayName -like "*Dell Command*Update*" }

# --- Already-installed check (UWP target) ---
$uwpEntry = $allEntries | Where-Object { $_.DisplayName -eq $uwpDisplayName } | Select-Object -First 1

if ($uwpEntry -and -not $forceinstall) {
    Write-Host "Dell Command | Update for Windows Universal is already installed (version $($uwpEntry.DisplayVersion)). Set forceinstall to override."
    exit 0
}

if ($uwpEntry -and $forceinstall) {
    Write-Host "forceinstall set — reinstalling Dell Command | Update for Windows Universal (currently $($uwpEntry.DisplayVersion))."
}

# --- Uninstall legacy versions if present ---
$legacyEntries = $allEntries | Where-Object { $legacyNames -contains $_.DisplayName }

foreach ($entry in $legacyEntries) {
    Write-Host "Uninstalling $($entry.DisplayName) (version $($entry.DisplayVersion))..."
    if ($entry.UninstallString -match "MsiExec") {
        $guid = $entry.PSChildName
        Start-Process -Wait -FilePath "msiexec.exe" -ArgumentList "/x `"$guid`" /quiet /norestart"
    } else {
        $cmd = if ($entry.QuietUninstallString) { $entry.QuietUninstallString } else { $entry.UninstallString }
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c `"$cmd`""
    }
    Write-Host "Uninstall of $($entry.DisplayName) complete."
}

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Download ---
Write-Host "Downloading Dell Command | Update for Windows Universal..."
Invoke-WebRequest -Uri $downloadUri -OutFile $outFile -UseBasicParsing
Write-Host "Download complete."

# --- Install ---
Write-Host "Installing Dell Command | Update for Windows Universal..."
Start-Process -Wait -FilePath $outFile -ArgumentList "/s"
Write-Host "Dell Command | Update for Windows Universal installed successfully."
