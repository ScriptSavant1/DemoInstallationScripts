# ==============================================================================
# 04_Upgrade_LoadGenerator.ps1
# LRE Load Generator (OneLG) Upgrade  |  25.1 -> 26.1
# ------------------------------------------------------------------------------
# Run this on each LOAD GENERATOR machine (OneLG standalone install).
# The Load Generator purpose/location in LRE Administration is preserved.
#
# Prerequisites:
#   - 01_Upgrade_LREServer.ps1 must have completed on the Server.
#   - SetupOneLG.exe must be present at: $InstallerShare\Standalone Applications\
#
# Usage (elevated PowerShell):
#   .\04_Upgrade_LoadGenerator.ps1
# ==============================================================================
#Requires -RunAsAdministrator

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
. "$ScriptRoot\Common-Functions.ps1"
. "$ScriptRoot\..\config\upgrade_config.ps1"

$Global:LogFile = "$LogDir\LoadGenerator_Upgrade_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Initialize-Log -LogFile $Global:LogFile

Write-Banner "LOAD GENERATOR (OneLG) UPGRADE  |  25.1 -> 26.1  |  $(hostname)  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

# ==============================================================================
# PHASE 1 — PRE-UPGRADE CHECKS
# ==============================================================================
Write-Log "PHASE 1 — Pre-Upgrade Checks" -Level STEP

# 1.1 — Installer share
Write-Log "Checking installer share: $InstallerShare" -Level INFO
if (-not (Test-ShareAccessible -SharePath $InstallerShare)) {
    Write-Log "Installer share not accessible: $InstallerShare" -Level ERROR
    exit 1
}

# 1.2 — Determine install type: OneLG MSI vs setup.exe
$useSetupExe = $false
$useMsi      = $false
if (Test-Path $SetupOneLGExe) {
    $useSetupExe = $true
    Write-Log "Found: $SetupOneLGExe  (will use setup.exe method)" -Level INFO
} elseif (Test-Path $OneLGMsi) {
    $useMsi = $true
    Write-Log "Found: $OneLGMsi  (will use MSI method)" -Level INFO
} else {
    Write-Log "Neither SetupOneLG.exe nor OneLG_x64.msi found in installer share." -Level ERROR
    Write-Log "Expected locations:" -Level ERROR
    Write-Log "  $SetupOneLGExe" -Level ERROR
    Write-Log "  $OneLGMsi" -Level ERROR
    exit 1
}

# 1.3 — Verify current installed version
Write-Log "Checking installed LRE/OneLG version ..." -Level INFO
$installed = Get-LREInstalledVersion
if (-not $installed) {
    # OneLG may register under a slightly different name - broaden the check
    $installed = Get-ChildItem -Path @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ) -ErrorAction SilentlyContinue |
    ForEach-Object { Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue } |
    Where-Object { $_.DisplayName -like "*LoadRunner*" -or $_.DisplayName -like "*OneLG*" -or $_.DisplayName -like "*Performance*" } |
    Select-Object -First 1 |
    ForEach-Object { [PSCustomObject]@{ DisplayName=$_.DisplayName; Version=$_.DisplayVersion; InstallDir=$_.InstallLocation } }
}

if ($installed) {
    Write-Log "Installed: $($installed.DisplayName)  Version: $($installed.Version)" -Level INFO
    if ($installed.Version -notlike "*$ExpectedCurrentVersion*") {
        Write-Log "Expected version $ExpectedCurrentVersion but found $($installed.Version)." -Level ERROR
        exit 1
    }
    Write-Log "Version check passed." -Level SUCCESS
} else {
    Write-Log "Could not detect installed LRE/OneLG product. Proceeding cautiously ..." -Level WARN
    Confirm-Action "Installed version check inconclusive. Proceed with upgrade anyway?"
}

# 1.4 — Disk space
if (-not (Test-DiskSpace -Path "C:\" -MinimumFreeGB 10)) {
    exit 1
}

# 1.5 — Pending reboot
if (Test-PendingReboot) {
    Write-Log "Pending reboot detected. Reboot first, then re-run." -Level ERROR
    exit 1
}

# ==============================================================================
# PHASE 2 — STOP AGENT SERVICES
# ==============================================================================
Write-Log "PHASE 2 — Stopping LRE Agent Services" -Level STEP

$agentSvc  = Resolve-ServiceName $ServiceAgent
$remoteSvc = Resolve-ServiceName $ServiceRemoteMgmt

if ($agentSvc)  { Stop-ServiceSafely -ServiceName $agentSvc  }
if ($remoteSvc) { Stop-ServiceSafely -ServiceName $remoteSvc }

# ==============================================================================
# PHASE 3 — INSTALL PREREQUISITES
# ==============================================================================
Write-Log "PHASE 3 — Installing Prerequisites" -Level STEP

function Install-Prereq {
    param([string]$Name, [string]$Path, [string]$Args)
    if (-not (Test-Path $Path)) {
        Write-Log "$Name not found at $Path - skipping." -Level WARN
        return
    }
    Write-Log "Installing $Name ..." -Level INFO
    $code = Invoke-Installer -Executable $Path -Arguments $Args -LogFile $Global:LogFile -TimeoutMinutes 15
    Write-Log "$Name exit code: $code" -Level INFO
}

Install-Prereq ".NET Framework 4.8" $DotNet48Exe '/LCID /q /norestart /c:"install /q"'
Install-Prereq "VC++ Redist x86"    $VCRedistX86 '/quiet /norestart'
Install-Prereq "VC++ Redist x64"    $VCRedistX64 '/quiet /norestart'

# ==============================================================================
# PHASE 4 — RUN OneLG UPGRADE
# ==============================================================================
Write-Log "PHASE 4 — Running OneLG Upgrade" -Level STEP
Confirm-Action "About to run OneLG upgrade on $(hostname). Continue?"

$installerLog = "$LogDir\OneLG_Installer_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$exitCode     = $null

if ($useSetupExe) {
    # ---------------------------------------------------
    # Method A: setup.exe (preferred)
    # Guide:  SetupOneLG.exe -s -sp"/s" IS_RUNAS_SERVICE=1 START_LGA=1
    # ---------------------------------------------------
    $installerArgs = "-s -sp`"/s`" IS_RUNAS_SERVICE=1 START_LGA=1"
    if (-not [string]::IsNullOrWhiteSpace($OneLGInstallDir)) {
        $installerArgs += " INSTALLDIR=`"$OneLGInstallDir`""
    }
    Write-Log "Method: setup.exe" -Level INFO
    Write-Log "Command: $SetupOneLGExe $installerArgs" -Level INFO

    $exitCode = Invoke-Installer `
        -Executable $SetupOneLGExe `
        -Arguments  $installerArgs `
        -LogFile    $installerLog `
        -TimeoutMinutes 60

} elseif ($useMsi) {
    # ---------------------------------------------------
    # Method B: MSI (fallback)
    # Guide:  msiexec /i OneLG_x64.msi /qb IS_RUNAS_SERVICE=1 START_LGA="1"
    # ---------------------------------------------------
    $msiLog = "$installerLog.msi.log"
    $installerArgs = "/i `"$OneLGMsi`" /qb /l*vx `"$msiLog`" IS_RUNAS_SERVICE=1 START_LGA=1"
    if (-not [string]::IsNullOrWhiteSpace($OneLGInstallDir)) {
        $installerArgs += " INSTALLDIR=`"$OneLGInstallDir`""
    }
    Write-Log "Method: msiexec" -Level INFO
    Write-Log "Command: msiexec $installerArgs" -Level INFO

    $exitCode = Invoke-Installer `
        -Executable "msiexec.exe" `
        -Arguments  $installerArgs `
        -LogFile    $installerLog `
        -TimeoutMinutes 60
}

switch ($exitCode) {
    0    { Write-Log "OneLG upgrade completed successfully." -Level SUCCESS }
    3010 { Write-Log "OneLG upgrade succeeded - REBOOT REQUIRED." -Level WARN }
    1602 { Write-Log "Upgrade cancelled." -Level WARN; exit 1602 }
    1603 { Write-Log "Fatal install error (1603). Check: $installerLog" -Level ERROR; exit 1603 }
    default {
        Write-Log "Unexpected exit code: $exitCode. Check: $installerLog" -Level ERROR
        exit $exitCode
    }
}

if ($exitCode -eq 3010) {
    Request-RebootIfNeeded -ExitCode $exitCode
}

# ==============================================================================
# PHASE 5 — POST-UPGRADE: RESTART SERVICES
# ==============================================================================
Write-Log "PHASE 5 — Starting Services" -Level STEP

if ($agentSvc)  { Start-ServiceSafely -ServiceName $agentSvc  }
if ($remoteSvc) { Start-ServiceSafely -ServiceName $remoteSvc }

# ==============================================================================
# PHASE 6 — VERIFY AGENT SERVICE MODE
# ==============================================================================
Write-Log "PHASE 6 — Verifying Agent Configuration" -Level STEP

# The OneLG agent must run as a SERVICE (IS_RUNAS_SERVICE=1).
# Verify the service exists and is running.
$agentSvcResolved = Resolve-ServiceName $ServiceAgent
if ($agentSvcResolved) {
    $status = Get-ServiceStatus -ServiceName $agentSvcResolved
    if ($status -eq "Running") {
        Write-Log "Agent service '$agentSvcResolved' is Running." -Level SUCCESS
    } else {
        Write-Log "Agent service '$agentSvcResolved' status: $status" -Level WARN
        Write-Log "If the service is not running, the Load Generator will not be usable." -Level WARN
    }
} else {
    Write-Log "Agent service not found after upgrade." -Level WARN
    Write-Log "If LoadRunner Agent was configured to run as a process (not service)," -Level WARN
    Write-Log "it must be started manually or reconfigured via Agent Settings." -Level WARN
}

# ==============================================================================
# VERIFY VERSION
# ==============================================================================
Write-Log "Verifying new version ..." -Level INFO
$newInstalled = Get-LREInstalledVersion
if ($newInstalled) {
    if ($newInstalled.Version -like "*$TargetVersion*") {
        Write-Log "Version confirmed: $($newInstalled.Version)" -Level SUCCESS
    } else {
        Write-Log "Post-upgrade version: $($newInstalled.Version) (expected $TargetVersion)" -Level WARN
    }
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Banner "LOAD GENERATOR UPGRADE COMPLETE  |  $(hostname)"
Write-Log "Log file     : $Global:LogFile" -Level INFO
Write-Log "Installer log: $installerLog" -Level INFO
Write-Log "" -Level INFO
Write-Log "NEXT STEPS:" -Level INFO
Write-Log "  - In LRE Administration > Hosts, verify this LG shows version $TargetVersion." -Level INFO
Write-Log "  - If the host shows 'Reconfigure needed', click Reconfigure Host." -Level INFO
Write-Log "  - Run 05_PostUpgradeVerify.ps1 after all components are upgraded." -Level INFO
Write-Log "  - Note: Do NOT delete the IUSR_METRO account unless the system" -Level INFO
Write-Log "    user was configured to a different Windows account." -Level INFO

Stop-Transcript -ErrorAction SilentlyContinue
