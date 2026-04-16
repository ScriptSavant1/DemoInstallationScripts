# ==============================================================================
# Upgrade_LoadGenerator.ps1  |  LRE Load Generator (OneLG) Upgrade  |  25.1 -> 26.1
# ==============================================================================
# Self-contained - no external dependencies.
# Place this script on the load generator machine, then run.
#
# Usage:
#   .\Upgrade_LoadGenerator.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Standalone Applications\SetupOneLG.exe"
# ==============================================================================

param(
    [Parameter(Mandatory)][string]$InstallerPath
)

# Self-elevate if not Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $argList = "-NoProfile -NoExit -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`" -InstallerPath `"$InstallerPath`""
    Start-Process powershell.exe -Verb RunAs -ArgumentList $argList
    exit
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# CONSTANTS
# ==============================================================================
$InstallDir             = "D:\LREOneLG"
$LogDir                 = "D:\LRE_UpgradeLogs"
$CertSharePath          = "\\rbsres01\grpareas\LRE\LREadmin\LRE_CERTS"
$ExpectedCurrentVersion = "25.1"
$TargetVersion          = "26.1"

# Derived paths
$CACertPath  = "$InstallDir\dat\cert\verify\cacert.cer"
$CertPath    = "$InstallDir\dat\cert\cert.cer"
$GenCertExe  = "$InstallDir\bin\gen_cert.exe"

# Service name candidates
$ServiceAgentCandidates      = @("magentservice","LoadRunnerAgentService","LRAgentService","OpenText Performance Engineering Agent Service")
$ServiceRemoteMgmtCandidates = @("AlAgent","RemoteManagementAgent","al_agent","OpenText Performance Engineering Remote Management Agent")

# ==============================================================================
# EMBEDDED HELPER FUNCTIONS
# ==============================================================================

function Initialize-Log {
    param([string]$LogFile)
    if (-not (Test-Path (Split-Path $LogFile -Parent))) {
        New-Item -ItemType Directory -Path (Split-Path $LogFile -Parent) -Force | Out-Null
    }
    Start-Transcript -Path ($LogFile -replace '\.log$', '_transcript.log') -Append -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','STEP')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"
    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'STEP'    { Write-Host "`n$line" -ForegroundColor White }
    }
    if ($Global:LogFile) { Add-Content -Path $Global:LogFile -Value $line }
}

function Write-Banner {
    param([string]$Title)
    $bar = "=" * 72
    Write-Host "`n$bar" -ForegroundColor Magenta
    Write-Host "  $Title" -ForegroundColor Magenta
    Write-Host "$bar`n" -ForegroundColor Magenta
    Write-Log "=== $Title ===" -Level INFO
}

function Resolve-ServiceName {
    param([string[]]$Candidates)
    foreach ($name in $Candidates) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if ($svc) { return $svc.Name }
    }
    return $null
}

function Stop-ServiceSafely {
    param([string]$ServiceName, [int]$TimeoutSeconds = 120)
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) { Write-Log "Service '$ServiceName' not found - skipping stop." -Level WARN; return }
    if ($svc.Status -eq 'Stopped') { Write-Log "Service '$ServiceName' already stopped." -Level INFO; return }
    Write-Log "Stopping service: $ServiceName ..." -Level INFO
    try {
        Stop-Service -Name $ServiceName -Force -ErrorAction Stop
        $svc.WaitForStatus('Stopped', (New-TimeSpan -Seconds $TimeoutSeconds))
        Write-Log "Service '$ServiceName' stopped." -Level SUCCESS
    } catch {
        Write-Log "Failed to stop '$ServiceName': $_" -Level ERROR
        throw
    }
}

function Start-ServiceSafely {
    param([string]$ServiceName, [int]$TimeoutSeconds = 120)
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) { Write-Log "Service '$ServiceName' not found - skipping start." -Level WARN; return }
    if ($svc.Status -eq 'Running') { Write-Log "Service '$ServiceName' already running." -Level INFO; return }
    Write-Log "Starting service: $ServiceName ..." -Level INFO
    try {
        Start-Service -Name $ServiceName -ErrorAction Stop
        $svc.WaitForStatus('Running', (New-TimeSpan -Seconds $TimeoutSeconds))
        Write-Log "Service '$ServiceName' started." -Level SUCCESS
    } catch {
        Write-Log "Failed to start '$ServiceName': $_" -Level ERROR
        throw
    }
}

function Get-LREInstalledVersion {
    $searchTerms = @(
        "OpenText Enterprise Performance Engineering",
        "OpenText Professional Performance Engineering",
        "Micro Focus Performance Center",
        "HP Performance Center",
        "LoadRunner",
        "OneLG"
    )
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($regPath in $regPaths) {
        if (-not (Test-Path $regPath)) { continue }
        foreach ($key in (Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue)) {
            $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
            if (-not $props -or -not $props.PSObject.Properties['DisplayName']) { continue }
            foreach ($term in $searchTerms) {
                if ($props.DisplayName -like "*$term*") {
                    return [PSCustomObject]@{
                        DisplayName = $props.DisplayName
                        Version     = $props.DisplayVersion
                        InstallDir  = $props.InstallLocation
                    }
                }
            }
        }
    }
    return $null
}

function Test-DiskSpace {
    param([string]$Path = "D:\", [long]$MinimumFreeGB = 10)
    $drive = Split-Path -Qualifier $Path
    $disk  = Get-PSDrive -Name ($drive.TrimEnd(':')) -ErrorAction SilentlyContinue
    if (-not $disk) {
        $wmiDisk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$drive'" -ErrorAction SilentlyContinue
        if ($wmiDisk) { $freeGB = [math]::Round($wmiDisk.FreeSpace / 1GB, 2) }
        else { Write-Log "Cannot determine disk space for '$Path'." -Level WARN; return $true }
    } else {
        $freeGB = [math]::Round($disk.Free / 1GB, 2)
    }
    Write-Log "Disk free on '$drive': $freeGB GB  (minimum: $MinimumFreeGB GB)" -Level INFO
    if ($freeGB -lt $MinimumFreeGB) {
        Write-Log "INSUFFICIENT DISK SPACE. Free: $freeGB GB, Required: $MinimumFreeGB GB" -Level ERROR
        return $false
    }
    return $true
}

function Test-PendingReboot {
    $pendingKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
        "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
    )
    foreach ($key in $pendingKeys) { if (Test-Path $key) { return $true } }
    $ccm = Invoke-WmiMethod -Namespace "root\ccm\clientsdk" -Class "CCM_ClientUtilities" -Name "DetermineIfRebootPending" -ErrorAction SilentlyContinue
    if ($ccm -and ($ccm.RebootPending -or $ccm.IsHardRebootPending)) { return $true }
    return $false
}

function Test-ShareAccessible {
    param([string]$SharePath)
    try { return (Test-Path $SharePath -ErrorAction Stop) } catch { return $false }
}

# ==============================================================================
# LOGGING SETUP
# ==============================================================================
$Global:LogFile = "$LogDir\LoadGenerator_Upgrade_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Initialize-Log -LogFile $Global:LogFile

Write-Banner "LOAD GENERATOR (OneLG) UPGRADE  |  25.1 -> 26.1  |  $(hostname)  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Write-Log "InstallerPath: $InstallerPath" -Level INFO
Write-Log "InstallDir   : $InstallDir" -Level INFO


# ==============================================================================
# PHASE 1 - PRE-UPGRADE CHECKS
# ==============================================================================
Write-Log "PHASE 1 - Pre-Upgrade Checks" -Level STEP

Write-Log "[1.1] Verifying installer exists ..." -Level INFO
if (-not (Test-Path $InstallerPath)) {
    Write-Log "Installer not found: $InstallerPath" -Level ERROR
    exit 1
}
Write-Log "[1.1] Installer found: $InstallerPath" -Level SUCCESS

Write-Log "[1.2] Checking installed LRE/OneLG version ..." -Level INFO
$installed = Get-LREInstalledVersion
if ($installed) {
    Write-Log "Installed: $($installed.DisplayName)  Version: $($installed.Version)" -Level INFO
    if ($installed.Version -notlike "*$ExpectedCurrentVersion*") {
        Write-Log "Expected version $ExpectedCurrentVersion but found $($installed.Version)." -Level ERROR
        exit 1
    }
    Write-Log "[1.2] Version check passed ($ExpectedCurrentVersion detected)." -Level SUCCESS
} else {
    Write-Log "[1.2] Could not detect installed LRE/OneLG product - proceeding cautiously." -Level WARN
}

Write-Log "[1.3] Checking disk space ..." -Level INFO
if (-not (Test-DiskSpace -Path "D:\" -MinimumFreeGB 10)) { exit 1 }

Write-Log "[1.4] Checking pending reboot ..." -Level INFO
if (Test-PendingReboot) {
    Write-Log "Pending reboot detected. Please reboot and re-run." -Level ERROR
    exit 1
}
Write-Log "[1.4] No pending reboot." -Level SUCCESS


# ==============================================================================
# PHASE 2 - STOP AGENT SERVICES
# ==============================================================================
Write-Log "PHASE 2 - Stopping Agent Services" -Level STEP

$agentSvc  = Resolve-ServiceName $ServiceAgentCandidates
$remoteSvc = Resolve-ServiceName $ServiceRemoteMgmtCandidates

if ($agentSvc)  { Stop-ServiceSafely -ServiceName $agentSvc  } else { Write-Log "Agent service not found - skipping." -Level WARN }
if ($remoteSvc) { Stop-ServiceSafely -ServiceName $remoteSvc } else { Write-Log "Remote Management service not found - skipping." -Level WARN }


# ==============================================================================
# PHASE 3 - RUN OneLG UPGRADE
# ==============================================================================
Write-Log "PHASE 3 - Running OneLG Upgrade" -Level STEP

$installerArgs = "-s -sp`"/s`" IS_RUNAS_SERVICE=1 START_LGA=1 NVINSTALL=Y"
Write-Log "Command: $InstallerPath $installerArgs" -Level INFO
Write-Log "Started at $(Get-Date -Format 'HH:mm:ss') - this may take 20-40 minutes ..." -Level INFO

$installStart = Get-Date
try {
    $proc = Start-Process -FilePath $InstallerPath -ArgumentList $installerArgs -Wait -PassThru -NoNewWindow
    $exitCode = $proc.ExitCode
} catch {
    Write-Log "Failed to launch installer: $_" -Level ERROR
    exit 1
}

$installDuration = (Get-Date) - $installStart
Write-Log "Installer exit code: $exitCode  |  Duration: $([math]::Round($installDuration.TotalMinutes, 1)) min" -Level INFO

switch ($exitCode) {
    0    { Write-Log "OneLG upgrade completed successfully." -Level SUCCESS }
    3010 { Write-Log "OneLG upgrade succeeded - REBOOT REQUIRED (handled in Phase 7)." -Level WARN }
    1602 { Write-Log "Upgrade cancelled." -Level WARN; exit 1602 }
    1603 { Write-Log "Fatal error during installation (1603). Check installer logs." -Level ERROR; exit 1603 }
    default { Write-Log "Unexpected exit code: $exitCode" -Level ERROR; exit $exitCode }
}


# ==============================================================================
# PHASE 4 - RESTART AGENT SERVICES
# ==============================================================================
Write-Log "PHASE 4 - Starting Agent Services" -Level STEP

if ($agentSvc)  { Start-ServiceSafely -ServiceName $agentSvc  } else { Write-Log "Agent service not found - may need manual start." -Level WARN }
if ($remoteSvc) { Start-ServiceSafely -ServiceName $remoteSvc } else { Write-Log "Remote Management service not found - may need manual start." -Level WARN }


# ==============================================================================
# PHASE 5 - POST-INSTALL CERTIFICATE REPLACEMENT
# ==============================================================================
Write-Log "PHASE 5 - Post-Install Configuration" -Level STEP

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$hn = (hostname).ToLower()

Write-Log "[5.1] Replacing CA and TLS certificates ..." -Level INFO

$certUpdateOk    = $true
$caCertName      = "lre_dev_cacert.cer"
$caCertSource    = Join-Path $CertSharePath $caCertName
$expectedCertName = "lre_dev_${hn}_lg.cer"

if (Test-ShareAccessible -SharePath $CertSharePath) {
    Write-Log "[5.1] CA cert source: `"$caCertSource`"" -Level INFO
    if (Test-Path $caCertSource) {
        $caCertDestDir = Split-Path $CACertPath -Parent
        if (-not (Test-Path $caCertDestDir)) { New-Item -ItemType Directory -Path $caCertDestDir -Force | Out-Null }
        if (Test-Path $CACertPath) {
            Copy-Item -Path $CACertPath -Destination "${CACertPath}.bak_${ts}" -Force
            Write-Log "[5.1] Backed up existing CA cert to `"${CACertPath}.bak_${ts}`"" -Level INFO
        }
        Copy-Item -Path $caCertSource -Destination $CACertPath -Force
        Write-Log "[5.1] Copied `"$caCertName`" to `"$CACertPath`"" -Level INFO
    } else {
        Write-Log "[5.1] CA cert not found: $caCertSource" -Level WARN
        $certUpdateOk = $false
    }

    Write-Log "[5.1] TLS cert: looking for `"$expectedCertName`"" -Level INFO
    $lgCertFile = Get-ChildItem -Path $CertSharePath -Filter $expectedCertName -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $lgCertFile) {
        Write-Log "[5.1] Expected cert '$expectedCertName' not found, trying wildcard *${hn}*.cer ..." -Level WARN
        $lgCertFile = Get-ChildItem -Path $CertSharePath -Filter "*${hn}*.cer" -ErrorAction SilentlyContinue | Select-Object -First 1
    }

    if ($lgCertFile) {
        if (Test-Path $CertPath) {
            Copy-Item -Path $CertPath -Destination "${CertPath}.bak_${ts}" -Force
            Write-Log "[5.1] Backed up existing TLS cert to `"${CertPath}.bak_${ts}`"" -Level INFO
        }
        Copy-Item -Path $lgCertFile.FullName -Destination $CertPath -Force
        Write-Log "[5.1] Copied `"$($lgCertFile.Name)`" to `"$CertPath`"" -Level INFO
    } else {
        Write-Log "[5.1] TLS cert not found in $CertSharePath for hostname '$hn'." -Level WARN
        $certUpdateOk = $false
    }

    if ($certUpdateOk -and (Test-Path $GenCertExe)) {
        Write-Log "[5.1] Running gen_cert.exe -verify ..." -Level INFO
        try {
            $verifyOutput = & $GenCertExe -verify 2>&1
            $verifyOutput | ForEach-Object { Write-Log "[5.1]   $_" -Level INFO }
            if ($verifyOutput -join " " -match "success|verified") {
                Write-Log "[5.1] SUCCESS: Certificate replacement and verification passed." -Level SUCCESS
            } else {
                Write-Log "[5.1] Certificate verification - review output above." -Level WARN
            }
        } catch {
            Write-Log "[5.1] gen_cert.exe -verify failed: $_" -Level WARN
        }
    }
} else {
    Write-Log "[5.1] Certificate share not accessible: $CertSharePath - skipping cert replacement." -Level WARN
}


# ==============================================================================
# PHASE 6 - VERIFY
# ==============================================================================
Write-Log "PHASE 6 - Verification" -Level STEP

$agentSvcResolved = Resolve-ServiceName $ServiceAgentCandidates
if ($agentSvcResolved) {
    $svc = Get-Service -Name $agentSvcResolved -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq 'Running') {
        Write-Log "  Agent service '$agentSvcResolved' is Running (service mode confirmed)." -Level SUCCESS
    } else {
        Write-Log "  Agent service '$agentSvcResolved' status: $($svc.Status) - verify manually." -Level WARN
    }
} else {
    Write-Log "  Agent service not found after upgrade. Verify agent is running in service mode." -Level WARN
}

$newInstalled = Get-LREInstalledVersion
if ($newInstalled) {
    if ($newInstalled.Version -like "*$TargetVersion*") {
        Write-Log "  Version confirmed: $($newInstalled.Version)" -Level SUCCESS
    } else {
        Write-Log "  Installed version: $($newInstalled.Version) (expected $TargetVersion)" -Level WARN
    }
}


# ==============================================================================
# PHASE 7 - SUMMARY
# ==============================================================================
Write-Banner "LOAD GENERATOR UPGRADE COMPLETE  |  $(hostname)"
Write-Log "Upgrade log: $Global:LogFile" -Level INFO
Write-Log "" -Level INFO
Write-Log "NEXT STEPS:" -Level WARN
Write-Log "  1. In LRE Administration > Hosts, verify this LG shows version $TargetVersion." -Level INFO
Write-Log "  2. If 'Reconfigure needed', click Reconfigure Host in Administration." -Level INFO
Write-Log "  3. In LRE Admin Portal, set Enable SSL = True for this Load Generator." -Level INFO
Write-Log "  4. Verify Load Generator connectivity from the LRE Admin Portal." -Level INFO

if ($exitCode -eq 3010) {
    Write-Log "Installer requires a reboot (exit code 3010). Rebooting in 30 seconds ..." -Level WARN
    Start-Sleep -Seconds 30
    Restart-Computer -Force
}

Write-Host "`nUpgrade complete. Press any key to close ..." -ForegroundColor Green
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
Stop-Transcript -ErrorAction SilentlyContinue
