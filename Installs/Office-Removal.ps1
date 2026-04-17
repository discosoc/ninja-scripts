# ==============================================================================
# Office-Removal.ps1
# Uses the Office Deployment Tool to silently remove all Office installations.
# Handles multi-language and OEM bundled installs (e.g. Dell images).
# ==============================================================================

# --- Variables ---
$ProgressPreference = 'SilentlyContinue'
$workingDir   = "C:\Scripts"
$odtUri       = "https://aka.ms/ODTsetup"
$odtInstaller = "$workingDir\ODTsetup.exe"
$configFile   = "$workingDir\OfficeRemoval.xml"

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Build removal config ---
$configXml = @"
<Configuration>
    <Remove All="TRUE"/>
    <Display Level="None" AcceptEULA="TRUE"/>
    <Property Name="AUTOACTIVATE" Value="0"/>
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
    <Property Name="SharedComputerLicensing" Value="0"/>
    <Property Name="PinIconsToTaskbar" Value="FALSE"/>
</Configuration>
"@

$configXml | Out-File -FilePath $configFile -Encoding utf8 -Force
Write-Host "Removal configuration written to $configFile"

# --- Download ODT ---
Write-Host "Downloading Office Deployment Tool..."
Invoke-WebRequest -Uri $odtUri -OutFile $odtInstaller -UseBasicParsing
Write-Host "Download complete."

# --- Extract ODT ---
Write-Host "Extracting ODT..."
Start-Process -Wait -FilePath $odtInstaller -ArgumentList "/quiet /extract:`"$workingDir`""
Write-Host "Extraction complete."

# --- Run removal ---
Write-Host "Removing Office..."
Start-Process -Wait -FilePath "$workingDir\setup.exe" -ArgumentList "/configure `"$configFile`""
Write-Host "Office removal complete."
