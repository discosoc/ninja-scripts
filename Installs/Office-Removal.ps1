# ==============================================================================
# Office-Removal.ps1
# Uses the Office Deployment Tool to silently remove all Office installations.
# Handles multi-language and OEM bundled installs (e.g. Dell images).
# ==============================================================================

# --- Variables ---
$ProgressPreference = 'SilentlyContinue'
$workingDir   = "C:\Scripts"
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
Write-Host "Resolving Office Deployment Tool download link..."
try {
    $downloadPage = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=49117" -UseBasicParsing
} catch {
    Write-Host "ERROR: Failed to reach Microsoft download page. $_"
    exit 1
}
$odtUri = ($downloadPage.Links | Where-Object { $_.href -like "*officedeploymenttool*.exe" } | Select-Object -First 1 -ExpandProperty href)
if (-not $odtUri) {
    Write-Host "ERROR: Could not locate ODT download link on Microsoft's download page."
    exit 1
}
Write-Host "Downloading Office Deployment Tool..."
try {
    Invoke-WebRequest -Uri $odtUri -OutFile $odtInstaller -UseBasicParsing
} catch {
    Write-Host "ERROR: Failed to download ODT. $_"
    exit 1
}
Write-Host "Download complete."

# --- Extract ODT ---
Write-Host "Extracting ODT..."
$result = Start-Process -Wait -PassThru -FilePath $odtInstaller -ArgumentList "/quiet /extract:`"$workingDir`""
if ($result.ExitCode -ne 0) {
    Write-Host "ERROR: ODT extraction failed (exit code $($result.ExitCode))."
    exit 1
}
Write-Host "Extraction complete."

# --- Run removal ---
Write-Host "Removing Office..."
$result = Start-Process -Wait -PassThru -FilePath "$workingDir\setup.exe" -ArgumentList "/configure `"$configFile`""
if ($result.ExitCode -ne 0) {
    Write-Host "ERROR: Office removal failed (exit code $($result.ExitCode))."
    exit 1
}
Write-Host "Office removal complete."
