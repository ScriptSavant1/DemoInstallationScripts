# ==============================================================================
# 01_Upgrade_LREServer.ps1
# LRE Server Upgrade  |  25.1 -> 26.1
# ------------------------------------------------------------------------------
# Run this on the LRE SERVER machine as the FIRST upgrade step.
# If you have multiple LRE Servers (clustered), run this on each server
# - but STOP IIS + Backend + Alerts on ALL servers before starting any install.
#
# After this script completes it writes the Public Key to:
#   $InstallerShare\PublicKey.txt
# Host upgrade scripts (02, 03, 04) will read from that file.
#
# Usage (elevated PowerShell):
#   .\01_Upgrade_LREServer.ps1
# ==============================================================================
#Requires -RunAsAdministrator

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
. "$ScriptRoot\Common-Functions.ps1"
. "$ScriptRoot\..\config\upgrade_config.ps1"

$Global:LogFile = "$LogDir\LREServer_Upgrade_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Initialize-Log -LogFile $Global:LogFile

Write-Banner "LRE SERVER UPGRADE  |  25.1 -> 26.1  |  $(hostname)  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

# ==============================================================================
# PHASE 1 — PRE-UPGRADE CHECKS
# ==============================================================================
Write-Log "PHASE 1 — Pre-Upgrade Checks" -Level STEP

# 1.1 — Verify installer share
Write-Log "Checking installer share: $InstallerShare" -Level INFO
if (-not (Test-ShareAccessible -SharePath $InstallerShare)) {
    Write-Log "Installer share is not accessible: $InstallerShare" -Level ERROR
    Write-Log "Ensure the network share is mounted and you have read permissions." -Level ERROR
    exit 1
}
if (-not (Test-Path $SetupServerExe)) {
    Write-Log "setup_server.exe not found at: $SetupServerExe" -Level ERROR
    exit 1
}
Write-Log "Installer media found." -Level SUCCESS

# 1.2 — Verify current LRE version
Write-Log "Checking installed LRE version ..." -Level INFO
$installed = Get-LREInstalledVersion
if (-not $installed) {
    Write-Log "LRE does not appear to be installed on this machine." -Level ERROR
    exit 1
}
Write-Log "Installed: $($installed.DisplayName)  Version: $($installed.Version)" -Level INFO
if ($installed.Version -notlike "*$ExpectedCurrentVersion*") {
    Write-Log "Expected version $ExpectedCurrentVersion but found $($installed.Version)." -Level ERROR
    Write-Log "This script upgrades $ExpectedCurrentVersion to $TargetVersion only." -Level ERROR
    exit 1
}
Write-Log "Version check passed ($ExpectedCurrentVersion detected)." -Level SUCCESS

# 1.3 — Disk space
if (-not (Test-DiskSpace -Path "C:\" -MinimumFreeGB 20)) {
    exit 1
}

# 1.4 — Pending reboot
if (Test-PendingReboot) {
    Write-Log "This machine has a pending reboot. Please reboot and re-run." -Level ERROR
    exit 1
}

# ==============================================================================
# PHASE 2 — GATHER CREDENTIALS (prompt only if not set in config)
# ==============================================================================
Write-Log "PHASE 2 — Gathering Credentials" -Level STEP

$script:DbAdminPwd  = Get-SecureValue "Enter DB Admin password for '$DbAdminUser'"    $DbAdminPassword
$script:DbPwd       = Get-SecureValue "Enter DB User password for '$DbUsername'"       $DbPassword
$script:SysUserPwd  = Get-SecureValue "Enter System User password for '$SystemUserName'" $SystemUserPwd
$SecurePassphrase   = Get-SecureValue "Enter Secure Communication Passphrase (min 12 chars)" $SecurePassphrase

if ($SecurePassphrase.Length -lt 12) {
    Write-Log "Passphrase must be at least 12 alphanumeric characters." -Level ERROR
    exit 1
}

# ==============================================================================
# PHASE 3 — DB BACKUP CONFIRMATION
# ==============================================================================
Write-Log "PHASE 3 — Database Backup Confirmation" -Level STEP
Write-Host @"

  IMPORTANT — Before upgrading, ensure the following SQL Server databases
  have been backed up:

    Database Server : $DbServerHost : $DbServerPort
    Lab DB          : $LabDbName
    Admin DB        : $AdminDbName
    Site Mgmt DB    : $SiteDbName

  Contact your DBA to confirm backups are complete before continuing.

"@ -ForegroundColor Yellow
Confirm-Action "I confirm that database backups are complete. Continue?"

# ==============================================================================
# PHASE 4 — STOP SERVICES
# ==============================================================================
Write-Log "PHASE 4 — Stopping Services" -Level STEP

# Guide: Stop IIS, Backend Service, and Alerts Service before upgrade.
# If clustered, this must be done on ALL servers before installing on any.
Write-Log "NOTE: If this is a CLUSTERED environment, ensure services are stopped" -Level WARN
Write-Log "on ALL LRE Server nodes before proceeding." -Level WARN

# Stop IIS
Write-Log "Stopping IIS ..." -Level INFO
try {
    & iisreset /stop 2>&1 | ForEach-Object { Write-Log $_ -Level INFO }
    Write-Log "IIS stopped." -Level SUCCESS
} catch {
    Write-Log "iisreset /stop failed: $_" -Level WARN
}

# Stop Backend Service
$backendSvc = Resolve-ServiceName $ServiceBackend
if ($backendSvc) {
    Stop-ServiceSafely -ServiceName $backendSvc
} else {
    Write-Log "LRE Backend Service not found by name - skipping (may already be stopped)." -Level WARN
}

# Stop Alerts Service
$alertsSvc = Resolve-ServiceName $ServiceAlerts
if ($alertsSvc) {
    Stop-ServiceSafely -ServiceName $alertsSvc
} else {
    Write-Log "LRE Alerts Service not found by name - skipping." -Level WARN
}

# ==============================================================================
# PHASE 5 — INSTALL PREREQUISITES
# ==============================================================================
Write-Log "PHASE 5 — Installing Prerequisites" -Level STEP

function Install-Prereq {
    param([string]$Name, [string]$Path, [string]$Args)
    if (-not (Test-Path $Path)) {
        Write-Log "$Name installer not found at: $Path - skipping." -Level WARN
        return
    }
    Write-Log "Installing $Name ..." -Level INFO
    $code = Invoke-Installer -Executable $Path -Arguments $Args -LogFile $Global:LogFile -TimeoutMinutes 15
    if ($code -in @(0, 3010, 1641)) {
        Write-Log "$Name installed (exit: $code)." -Level SUCCESS
    } else {
        Write-Log "$Name installer returned exit code $code." -Level WARN
    }
}

Install-Prereq ".NET Framework 4.8"     $DotNet48Exe     '/LCID /q /norestart /c:"install /q"'
Install-Prereq ".NET Core Hosting 8.x"  $DotNetHostingExe '/quiet OPT_NO_RUNTIME=1 OPT_NO_SHAREDFX=1 OPT_NO_X86=1'
Install-Prereq "VC++ Redist x86"        $VCRedistX86     '/quiet /norestart'
Install-Prereq "VC++ Redist x64"        $VCRedistX64     '/quiet /norestart'

# ==============================================================================
# PHASE 6 — PREPARE UserInput.xml
# Strategy:
#   1. Copy vendor's UserInput.xml from the installer share and merge our values
#      into it by matching <Property Name="..."> nodes — safest approach.
#   2. Fall back to our own template if the share XML is not accessible.
# ==============================================================================
Write-Log "PHASE 6 — Preparing UserInput.xml" -Level STEP

if (-not (Test-Path $TempUserInputDir)) {
    New-Item -ItemType Directory -Path $TempUserInputDir -Force | Out-Null
}
$tempXml = "$TempUserInputDir\Server_UserInput_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"

# Make resolved credentials available to Get-ServerPropertyMap via script scope
$script:SysUserPwd  = $SystemUserPwd
$script:DbAdminPwd  = $DbAdminPassword
$script:DbPwd       = $DbPassword

$propertyMap = Get-ServerPropertyMap

New-UserInputXml `
    -VendorXmlPath         $ServerUserInput `
    -FallbackTemplatePath  "$ScriptRoot\..\config\LRE_Server_UserInput.xml" `
    -PropertyMap           $propertyMap `
    -OutputPath            $tempXml

# Log a sanitised summary (no passwords)
Write-Log "UserInput.xml property summary:" -Level INFO
$propertyMap.GetEnumerator() | Where-Object { $_.Key -notmatch "Pwd|Password|Key|Passphrase" } |
    Sort-Object Key | ForEach-Object { Write-Log "  $($_.Key) = $($_.Value)" -Level INFO }

# ==============================================================================
# PHASE 7 — RUN SILENT SERVER UPGRADE
# ==============================================================================
Write-Log "PHASE 7 — Running LRE Server Upgrade" -Level STEP
Confirm-Action "About to run LRE Server upgrade from $SetupServerExe. Continue?"

$installerLog = "$LogDir\LREServer_Installer_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$installerArgs = "/s USER_CONFIG_FILE_PATH=`"$tempXml`" INSTALLDIR=`"$ServerInstallDir`""

Write-Log "Installer command:" -Level INFO
Write-Log "  $SetupServerExe $installerArgs" -Level INFO

$exitCode = Invoke-Installer `
    -Executable $SetupServerExe `
    -Arguments  $installerArgs `
    -LogFile    $installerLog `
    -TimeoutMinutes 90

switch ($exitCode) {
    0    { Write-Log "LRE Server upgrade completed successfully." -Level SUCCESS }
    3010 { Write-Log "LRE Server upgrade succeeded - REBOOT REQUIRED." -Level WARN }
    1602 { Write-Log "Upgrade cancelled by user." -Level WARN; exit 1602 }
    1603 { Write-Log "Fatal error during installation (1603). Check installer log: $installerLog" -Level ERROR; exit 1603 }
    default {
        Write-Log "Unexpected exit code: $exitCode. Check log: $installerLog" -Level ERROR
        exit $exitCode
    }
}

# Handle reboot requirement
if ($exitCode -eq 3010) {
    Request-RebootIfNeeded -ExitCode $exitCode
}

# ==============================================================================
# PHASE 8 — POST-UPGRADE: VERIFY INSTALLER LOG
# ==============================================================================
Write-Log "PHASE 8 — Post-Upgrade Verification" -Level STEP

# Check vendor configuration wizard log
$configLog = "$ServerInstallDir\orchidtmp\Configuration\configurationWizardLog_pcs.txt"
if (Test-Path $configLog) {
    $logTail = Get-Content $configLog -Tail 30
    Write-Log "Last 30 lines of Configuration Wizard log:" -Level INFO
    $logTail | ForEach-Object { Write-Log "  $_" -Level INFO }
    if ($logTail -join " " -match "ERROR|FAIL|exception") {
        Write-Log "Potential errors detected in configuration log. Review: $configLog" -Level WARN
    }
} else {
    Write-Log "Configuration wizard log not found at: $configLog" -Level WARN
}

# ==============================================================================
# PHASE 9 — RESTART SERVICES
# ==============================================================================
Write-Log "PHASE 9 — Starting Services" -Level STEP

Write-Log "Starting IIS ..." -Level INFO
& iisreset /start 2>&1 | ForEach-Object { Write-Log $_ -Level INFO }

Start-Sleep -Seconds 10  # Allow IIS to fully initialize

if ($backendSvc) { Start-ServiceSafely -ServiceName $backendSvc }
if ($alertsSvc)  { Start-ServiceSafely -ServiceName $alertsSvc  }

# ==============================================================================
# PHASE 10 — CAPTURE PUBLIC KEY
# ==============================================================================
Write-Log "PHASE 10 — Capturing Public Key" -Level STEP

Write-Log "Waiting 30 seconds for LRE services to fully initialize ..." -Level INFO
Start-Sleep -Seconds 30

$publicKey = $null

# Attempt 1: REST API
try {
    Write-Log "Attempting to retrieve Public Key via REST API ..." -Level INFO
    $apiUrl  = "http://localhost/Admin/rest/v1/configuration/getPublicKey"
    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -TimeoutSec 30 -ErrorAction Stop
    if ($response.PublicKey -or $response.publicKey) {
        $publicKey = if ($response.PublicKey) { $response.PublicKey } else { $response.publicKey }
        Write-Log "Public Key retrieved via REST API." -Level SUCCESS
    }
} catch {
    Write-Log "REST API call failed: $_" -Level WARN
}

# Attempt 2: Read from pcs.config
if (-not $publicKey) {
    $pcsConfig = "$ServerInstallDir\dat\pcs.config"
    if (Test-Path $pcsConfig) {
        $configContent = Get-Content $pcsConfig -Raw
        if ($configContent -match 'PublicKey["\s=:]+([A-Za-z0-9+/=]{20,})') {
            $publicKey = $Matches[1]
            Write-Log "Public Key found in pcs.config." -Level SUCCESS
        }
    }
}

if ($publicKey) {
    # Write to network share for host scripts to consume
    try {
        Set-Content -Path $PublicKeySharePath -Value $publicKey -Encoding UTF8
        Write-Log "Public Key written to: $PublicKeySharePath" -Level SUCCESS
    } catch {
        Write-Log "Could not write Public Key to share: $_" -Level WARN
        Write-Log "You will need to manually provide the Public Key when running host upgrade scripts." -Level WARN
    }
    # Also write locally
    Set-Content -Path "$LogDir\PublicKey.txt" -Value $publicKey -Encoding UTF8
    Write-Log "Public Key also saved locally to: $LogDir\PublicKey.txt" -Level INFO
    Write-Host "`n  PUBLIC KEY (copy this for host upgrade scripts):`n  $publicKey`n" -ForegroundColor Green
} else {
    Write-Log "Could not automatically retrieve Public Key." -Level WARN
    Write-Log "ACTION REQUIRED: Retrieve the Public Key manually from:" -Level WARN
    Write-Log "  GET http://localhost/Admin/rest/v1/configuration/getPublicKey" -Level WARN
    Write-Log "  OR copy it from the Configuration Wizard 'Finish' page." -Level WARN
    Write-Log "  OR read from: $ServerInstallDir\dat\pcs.config" -Level WARN
    Write-Log "  Then update PublicKeySharePath in upgrade_config.ps1 before running host scripts." -Level WARN
}

# ==============================================================================
# PHASE 11 — CONFIG FILE CONSISTENCY CHECK
# ==============================================================================
Write-Log "PHASE 11 — Config File Consistency Check" -Level STEP

$pcsConfig = "$ServerInstallDir\dat\pcs.config"
Test-ConfigUserConsistency -ConfigFilePath $pcsConfig -ExpectedUser $SystemUserName | Out-Null

$appSettings = "$FileSystemRoot\system_config\appsettings.json"
if (Test-Path $appSettings) {
    Test-ConfigUserConsistency -ConfigFilePath $appSettings -ExpectedUser $SystemUserName | Out-Null
} else {
    Write-Log "appsettings.json not found at: $appSettings" -Level WARN
}

# ==============================================================================
# VERIFY UPGRADED VERSION
# ==============================================================================
Write-Log "Verifying new installed version ..." -Level INFO
$newInstalled = Get-LREInstalledVersion
if ($newInstalled) {
    if ($newInstalled.Version -like "*$TargetVersion*") {
        Write-Log "Version confirmed: $($newInstalled.Version)" -Level SUCCESS
    } else {
        Write-Log "Installed version after upgrade: $($newInstalled.Version) (expected $TargetVersion)" -Level WARN
    }
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Banner "LRE SERVER UPGRADE COMPLETE  |  $(hostname)"
Write-Log "Upgrade log : $Global:LogFile" -Level INFO
Write-Log "Installer log: $installerLog" -Level INFO
if ($publicKey) {
    Write-Log "Public Key  : $PublicKeySharePath" -Level INFO
}
Write-Log "" -Level INFO
Write-Log "NEXT STEPS:" -Level INFO
Write-Log "  1. Verify LRE Server is accessible via browser: https://$IisSecureHostName/LRE/" -Level INFO
Write-Log "  2. Run 02_Upgrade_Controller.ps1 on each Controller host." -Level INFO
Write-Log "  3. Run 03_Upgrade_DataProcessor.ps1 on each Data Processor host." -Level INFO
Write-Log "  4. Run 04_Upgrade_LoadGenerator.ps1 on each Load Generator." -Level INFO
Write-Log "  5. Run 05_PostUpgradeVerify.ps1 to validate all components." -Level INFO

# Secure cleanup of temp xml (contains passwords)
if (Test-Path $tempXml) {
    Remove-Item $tempXml -Force
    Write-Log "Temp UserInput.xml removed." -Level INFO
}

Stop-Transcript -ErrorAction SilentlyContinue
