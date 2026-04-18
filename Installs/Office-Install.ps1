# ==============================================================================
# Office-Install.ps1
# Downloads the Office Deployment Tool, builds a config from script variables,
# and installs Microsoft 365 Apps for Business.
#
# Ninja script variables (set via Bootstraps/Office-Install.ps1):
#   edition        - Office bitness: 32 or 64 (required)
#   includeproject - include Project Pro when set to any value
#   includevisio   - include Visio Pro when set to any value
# ==============================================================================

# --- Validate required variables ---
if (-not $edition) {
    Write-Host "ERROR: edition script variable is required (32 or 64)."
    exit 1
}
if ($edition -ne "32" -and $edition -ne "64") {
    Write-Host "ERROR: edition must be '32' or '64'."
    exit 1
}

# --- Variables ---
$ProgressPreference = 'SilentlyContinue'
$workingDir   = "C:\Scripts"
$odtInstaller = "$workingDir\ODTsetup.exe"
$configFile   = "$workingDir\OfficeConfig.xml"

# --- Ensure working directory ---
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Created working directory: $workingDir"
} else {
    Write-Host "Working directory already exists: $workingDir"
}

# --- Build optional product blocks ---
$projectBlock = ""
if ($includeproject) {
    $projectBlock = @"

    <Product ID="ProjectProRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
    </Product>
"@
}

$visioBlock = ""
if ($includevisio) {
    $visioBlock = @"

    <Product ID="VisioProRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
    </Product>
"@
}

# --- Build config XML ---
$configXml = @"
<Configuration>
  <Add OfficeClientEdition="$edition" Channel="SemiAnnual">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
    </Product>$projectBlock$visioBlock
  </Add>
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
  </AppSettings>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@

$configXml | Out-File -FilePath $configFile -Encoding utf8 -Force
Write-Host "Configuration written to $configFile"

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

# --- Run ODT ---
Write-Host "Running Office Deployment Tool..."
$result = Start-Process -Wait -PassThru -FilePath "$workingDir\setup.exe" -ArgumentList "/configure `"$configFile`""
if ($result.ExitCode -ne 0) {
    Write-Host "ERROR: Office deployment failed (exit code $($result.ExitCode))."
    exit 1
}
Write-Host "Microsoft 365 deployment complete."
