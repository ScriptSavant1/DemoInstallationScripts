# ==============================================================================
# Upgrade_Host.ps1  |  LRE Host Upgrade  |  25.1 -> 26.1
# ==============================================================================
# Self-contained - no external dependencies except the UserInput.xml.
# Place this script + LRE_Host_*_UserInput.xml + PublicKey.txt on the host,
# then run.
#
# Usage:
#   .\Upgrade_Host.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Setup\En\setup_host.exe" -ComponentType Controller
#   .\Upgrade_Host.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Setup\En\setup_host.exe" -ComponentType DataProcessor
# ==============================================================================

param(
    [Parameter(Mandatory)][string]$InstallerPath,
    [Parameter(Mandatory)][ValidateSet("Controller","DataProcessor")][string]$ComponentType
)

# Self-elevate if not Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $argList = "-NoProfile -NoExit -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`" -InstallerPath `"$InstallerPath`" -ComponentType $ComponentType"
    Start-Process powershell.exe -Verb RunAs -ArgumentList $argList
    exit
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# CONSTANTS
# ==============================================================================
$InstallDir             = "D:\LREHost"
$LogDir                 = "D:\LRE_UpgradeLogs"
$CertSharePath          = "\\rbsres01\grpareas\LRE\LREadmin\LRE_CERTS"
$ExpectedCurrentVersion = "25.1"
$TargetVersion          = "26.1"

# Derived paths
$LTOPBinDir    = "$InstallDir\bin\LTOPbin"
$HttpsConfigDir = "$InstallDir\conf\httpsconfigfiles"
$CACertPath    = "$InstallDir\dat\cert\verify\cacert.cer"
$CertPath      = "$InstallDir\dat\cert\cert.cer"
$GenCertExe    = "$InstallDir\bin\gen_cert.exe"
$LtsConfig     = "$InstallDir\dat\lts.config"
$AgentSettingsExe = "$InstallDir\bin\lr_agent_settings.exe"
$TempUserInputDir = "$LogDir\UserInput"

# Service name candidates
$ServiceAgentCandidates       = @("magentservice","LoadRunnerAgentService","LRAgentService","OpenText Performance Engineering Agent Service")
$ServiceRemoteMgmtCandidates  = @("AlAgent","RemoteManagementAgent","al_agent","OpenText Performance Engineering Remote Management Agent")
$ServiceLoadTestCandidates    = @("OpenText Performance Engineering Load Testing","LTOPService","LoadTestingService")

$ScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent

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
        "HP Performance Center"
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
$Global:LogFile = "$LogDir\LREHost_${ComponentType}_Upgrade_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Initialize-Log -LogFile $Global:LogFile

Write-Banner "LRE HOST UPGRADE  |  25.1 -> 26.1  |  $ComponentType  |  $(hostname)  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Write-Log "ComponentType: $ComponentType" -Level INFO
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

Write-Log "[1.2] Checking installed LRE version ..." -Level INFO
$installed = Get-LREInstalledVersion
if (-not $installed) {
    Write-Log "LRE Host does not appear to be installed on this machine." -Level ERROR
    exit 1
}
Write-Log "Installed: $($installed.DisplayName)  Version: $($installed.Version)" -Level INFO
if ($installed.Version -notlike "*$ExpectedCurrentVersion*") {
    Write-Log "Expected version $ExpectedCurrentVersion but found $($installed.Version)." -Level ERROR
    exit 1
}
Write-Log "[1.2] Version check passed ($ExpectedCurrentVersion detected)." -Level SUCCESS

Write-Log "[1.3] Checking disk space ..." -Level INFO
if (-not (Test-DiskSpace -Path "D:\" -MinimumFreeGB 10)) { exit 1 }

Write-Log "[1.4] Checking pending reboot ..." -Level INFO
if (Test-PendingReboot) {
    Write-Log "Pending reboot detected. Please reboot and re-run." -Level ERROR
    exit 1
}
Write-Log "[1.4] No pending reboot." -Level SUCCESS


# ==============================================================================
# PHASE 2 - RESOLVE PUBLIC KEY
# ==============================================================================
Write-Log "PHASE 2 - Resolving Public Key" -Level STEP

$publicKey = $null
$publicKeyFile = Join-Path $ScriptRoot "PublicKey.txt"

if (Test-Path $publicKeyFile) {
    $publicKey = (Get-Content $publicKeyFile -Raw -Encoding UTF8).Trim()
    Write-Log "[2.1] Public Key loaded from: $publicKeyFile" -Level SUCCESS
}

if ([string]::IsNullOrWhiteSpace($publicKey)) {
    Write-Log "[2.1] PublicKey.txt not found in script folder." -Level WARN
    Write-Host "`n  Public Key was not found at: $publicKeyFile" -ForegroundColor Yellow
    Write-Host "  Find it in one of:" -ForegroundColor Yellow
    Write-Host "    1. Server upgrade log: D:\LRE_UpgradeLogs\PublicKey.txt  (on the LRE Server)" -ForegroundColor Yellow
    Write-Host "    2. Server config file: D:\LREServer\dat\pcs.config" -ForegroundColor Yellow
    Write-Host "    3. REST API on server: GET http://<server>/Admin/rest/v1/configuration/getPublicKey" -ForegroundColor Yellow
    Write-Host "`n  Enter the Public Key (paste and press Enter): " -ForegroundColor Yellow -NoNewline
    $publicKey = Read-Host
}

if ([string]::IsNullOrWhiteSpace($publicKey)) {
    Write-Log "Public Key is required but was not provided. Aborting." -Level ERROR
    exit 1
}
Write-Log "[2.1] Public Key is set." -Level SUCCESS


# ==============================================================================
# PHASE 3 - STOP AGENT SERVICES
# ==============================================================================
Write-Log "PHASE 3 - Stopping Agent Services" -Level STEP

$agentSvc  = Resolve-ServiceName $ServiceAgentCandidates
$remoteSvc = Resolve-ServiceName $ServiceRemoteMgmtCandidates

if ($agentSvc)  { Stop-ServiceSafely -ServiceName $agentSvc  } else { Write-Log "Agent service not found - skipping." -Level WARN }
if ($remoteSvc) { Stop-ServiceSafely -ServiceName $remoteSvc } else { Write-Log "Remote Management service not found - skipping." -Level WARN }


# ==============================================================================
# PHASE 4 - PREPARE UserInput.xml
# ==============================================================================
Write-Log "PHASE 4 - Preparing UserInput.xml" -Level STEP

if (-not (Test-Path $TempUserInputDir)) {
    New-Item -ItemType Directory -Path $TempUserInputDir -Force | Out-Null
}

$xmlFile = Get-ChildItem -Path $ScriptRoot -Filter "LRE_Host_*_UserInput.xml" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $xmlFile) {
    Write-Log "No LRE_Host_*_UserInput.xml found in: $ScriptRoot" -Level ERROR
    exit 1
}
Write-Log "[4.1] Found UserInput.xml: $($xmlFile.FullName)" -Level SUCCESS

$tempXml = "$TempUserInputDir\Host_${ComponentType}_UserInput_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
Copy-Item -Path $xmlFile.FullName -Destination $tempXml -Force
Write-Log "[4.2] Copied to temp: $tempXml" -Level INFO

$hn = (hostname).ToLower()
$xmlContent = Get-Content $tempXml -Raw -Encoding UTF8

if ($xmlContent -match '\{\{HOSTNAME\}\}') {
    $xmlContent = $xmlContent -replace '\{\{HOSTNAME\}\}', $hn
    Write-Log "[4.3] Replaced {{HOSTNAME}} with '$hn'" -Level INFO
}
if ($xmlContent -match '\{\{PUBLIC_KEY\}\}') {
    $xmlContent = $xmlContent -replace '\{\{PUBLIC_KEY\}\}', $publicKey
    Write-Log "[4.3] Replaced {{PUBLIC_KEY}} in UserInput.xml" -Level INFO
}

Set-Content -Path $tempXml -Value $xmlContent -Encoding UTF8
Write-Log "[4.3] UserInput.xml prepared." -Level SUCCESS


# ==============================================================================
# PHASE 5 - RUN SILENT HOST UPGRADE
# ==============================================================================
Write-Log "PHASE 5 - Running LRE Host Upgrade" -Level STEP

$installerArgs = "/s INSTALLDIR=`"$InstallDir`" USER_CONFIG_FILE_PATH=`"$tempXml`" LRASPCHOST=1 INSTALL_JMETER=1 INSTALL_GATLING=1 NVINSTALL=Y"
Write-Log "Command: $InstallerPath $installerArgs" -Level INFO
Write-Log "Started at $(Get-Date -Format 'HH:mm:ss') - this may take 20-40 minutes ..." -Level INFO

$installStart = Get-Date
try {
    $proc = Start-Process -FilePath $InstallerPath -ArgumentList $installerArgs -Wait -PassThru -NoNewWindow
    $exitCode = $proc.ExitCode
} catch {
    Write-Log "Failed to launch installer: $_" -Level ERROR
    if (Test-Path $tempXml) { Remove-Item $tempXml -Force -ErrorAction SilentlyContinue }
    exit 1
}

$installDuration = (Get-Date) - $installStart
Write-Log "Installer exit code: $exitCode  |  Duration: $([math]::Round($installDuration.TotalMinutes, 1)) min" -Level INFO

if (Test-Path $tempXml) {
    Remove-Item $tempXml -Force -ErrorAction SilentlyContinue
    Write-Log "Temp UserInput.xml deleted." -Level INFO
}

switch ($exitCode) {
    0    { Write-Log "$ComponentType upgrade completed successfully." -Level SUCCESS }
    3010 { Write-Log "$ComponentType upgrade succeeded - REBOOT REQUIRED (handled in Phase 9)." -Level WARN }
    1602 { Write-Log "Upgrade cancelled by user." -Level WARN; exit 1602 }
    1603 { Write-Log "Fatal error during installation (1603). Check installer logs." -Level ERROR; exit 1603 }
    default { Write-Log "Unexpected exit code: $exitCode" -Level ERROR; exit $exitCode }
}


# ==============================================================================
# PHASE 6 - RESTART AGENT SERVICES
# ==============================================================================
Write-Log "PHASE 6 - Starting Agent Services" -Level STEP

if ($agentSvc)  { Start-ServiceSafely -ServiceName $agentSvc  } else { Write-Log "Agent service not found - may need manual start." -Level WARN }
if ($remoteSvc) { Start-ServiceSafely -ServiceName $remoteSvc } else { Write-Log "Remote Management service not found - may need manual start." -Level WARN }


# ==============================================================================
# PHASE 7 - POST-INSTALL CONFIGURATION
# ==============================================================================
Write-Log "PHASE 7 - Post-Install Configuration" -Level STEP

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'

# --- 7.1 LTOPSvc.exe.config swap ---
Write-Log "[7.1] Configure TLS - LTOPSvc.exe.config swap ..." -Level INFO
$ltopConfig    = Join-Path $LTOPBinDir "LTOPSvc.exe.config"
$ltopConfigSsl = Join-Path $HttpsConfigDir "LTOPSvc.exe.config-for_ssl"

if (Test-Path $ltopConfig) {
    $ltopConfigBak = "${ltopConfig}.bak_${ts}"
    Copy-Item -Path $ltopConfig -Destination $ltopConfigBak -Force
    Write-Log "[7.1] Backed up `"LTOPSvc.exe.config`" to `"$ltopConfigBak`"" -Level INFO
    if (Test-Path $ltopConfigSsl) {
        Copy-Item -Path $ltopConfigSsl -Destination $ltopConfig -Force
        Write-Log "[7.1] Copied `"LTOPSvc.exe.config-for_ssl`" from `"$HttpsConfigDir`" to `"$ltopConfig`"" -Level INFO
        Write-Log "[7.1] SUCCESS: LTOPSvc.exe.config replaced with SSL version." -Level SUCCESS
    } else {
        Write-Log "[7.1] LTOPSvc.exe.config-for_ssl not found at: $ltopConfigSsl - manual edit needed." -Level WARN
    }
} else {
    Write-Log "[7.1] LTOPSvc.exe.config not found: $ltopConfig - skipping." -Level WARN
}

Write-Log "[7.1] Restarting Load Testing service to apply TLS config ..." -Level INFO
$loadTestSvc = Resolve-ServiceName $ServiceLoadTestCandidates
if ($loadTestSvc) {
    Stop-ServiceSafely  -ServiceName $loadTestSvc
    Start-ServiceSafely -ServiceName $loadTestSvc
    Write-Log "[7.1] Load Testing service restarted." -Level SUCCESS
} else {
    Write-Log "[7.1] Load Testing service not found - may need manual restart." -Level WARN
}

# --- 7.2 Certificate replacement (always lre_dev_ prefix for Host) ---
Write-Log "[7.2] Replacing CA and TLS certificates ..." -Level INFO

$certUpdateOk = $true
$caCertName   = "lre_dev_cacert.cer"
$caCertSource = Join-Path $CertSharePath $caCertName
$certTypeSuffix = if ($ComponentType -eq "Controller") { "controller" } else { "dp" }
$expectedCertName = "lre_dev_${hn}_${certTypeSuffix}.cer"

if (Test-ShareAccessible -SharePath $CertSharePath) {
    Write-Log "[7.2] CA cert source: `"$caCertSource`"" -Level INFO
    if (Test-Path $caCertSource) {
        $caCertDestDir = Split-Path $CACertPath -Parent
        if (-not (Test-Path $caCertDestDir)) { New-Item -ItemType Directory -Path $caCertDestDir -Force | Out-Null }
        if (Test-Path $CACertPath) {
            Copy-Item -Path $CACertPath -Destination "${CACertPath}.bak_${ts}" -Force
            Write-Log "[7.2] Backed up existing CA cert to `"${CACertPath}.bak_${ts}`"" -Level INFO
        }
        Copy-Item -Path $caCertSource -Destination $CACertPath -Force
        Write-Log "[7.2] Copied `"$caCertName`" to `"$CACertPath`"" -Level INFO
    } else {
        Write-Log "[7.2] CA cert not found: $caCertSource" -Level WARN
        $certUpdateOk = $false
    }

    Write-Log "[7.2] TLS cert: ComponentType=$ComponentType -> looking for `"$expectedCertName`"" -Level INFO
    $hostCertFile = Get-ChildItem -Path $CertSharePath -Filter $expectedCertName -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $hostCertFile) {
        Write-Log "[7.2] Expected cert '$expectedCertName' not found, trying wildcard *${hn}*.cer ..." -Level WARN
        $hostCertFile = Get-ChildItem -Path $CertSharePath -Filter "*${hn}*.cer" -ErrorAction SilentlyContinue | Select-Object -First 1
    }

    if ($hostCertFile) {
        if (Test-Path $CertPath) {
            Copy-Item -Path $CertPath -Destination "${CertPath}.bak_${ts}" -Force
            Write-Log "[7.2] Backed up existing TLS cert to `"${CertPath}.bak_${ts}`"" -Level INFO
        }
        Copy-Item -Path $hostCertFile.FullName -Destination $CertPath -Force
        Write-Log "[7.2] Copied `"$($hostCertFile.Name)`" to `"$CertPath`"" -Level INFO
    } else {
        Write-Log "[7.2] TLS cert not found in $CertSharePath for hostname '$hn'." -Level WARN
        $certUpdateOk = $false
    }

    if ($certUpdateOk -and (Test-Path $GenCertExe)) {
        Write-Log "[7.2] Running gen_cert.exe -verify ..." -Level INFO
        try {
            $verifyOutput = & $GenCertExe -verify 2>&1
            $verifyOutput | ForEach-Object { Write-Log "[7.2]   $_" -Level INFO }
            if ($verifyOutput -join " " -match "success|verified") {
                Write-Log "[7.2] SUCCESS: Certificate replacement and verification passed." -Level SUCCESS
            } else {
                Write-Log "[7.2] Certificate verification - review output above." -Level WARN
            }
        } catch {
            Write-Log "[7.2] gen_cert.exe -verify failed: $_" -Level WARN
        }
    }
} else {
    Write-Log "[7.2] Certificate share not accessible: $CertSharePath - skipping cert replacement." -Level WARN
}

# --- 7.3 Agent settings ---
Write-Log "[7.3] Running lr_agent_settings.exe -check_client_cert 1 -restart_agent" -Level INFO
if (Test-Path $AgentSettingsExe) {
    try {
        $agentOutput = & $AgentSettingsExe -check_client_cert 1 -restart_agent 2>&1
        $agentOutput | ForEach-Object { Write-Log "[7.3]   $_" -Level INFO }
        Write-Log "[7.3] SUCCESS: Agent settings updated." -Level SUCCESS
    } catch {
        Write-Log "[7.3] lr_agent_settings.exe failed: $_" -Level WARN
    }
} else {
    Write-Log "[7.3] lr_agent_settings.exe not found: $AgentSettingsExe - skipping." -Level WARN
}


# ==============================================================================
# PHASE 8 - VERIFY
# ==============================================================================
Write-Log "PHASE 8 - Verification" -Level STEP

if (Test-Path $LtsConfig) {
    $ltsContent = Get-Content $LtsConfig -Raw -ErrorAction SilentlyContinue
    Write-Log "  lts.config found: $LtsConfig" -Level INFO
    if ($ltsContent -match 'IsSecured\s*=\s*"true"|ltopIsSecured\s*=\s*"true"') {
        Write-Log "  lts.config: TLS/secured flag is true." -Level SUCCESS
    } else {
        Write-Log "  lts.config: TLS/secured flag not detected - verify manually." -Level WARN
    }
} else {
    Write-Log "  lts.config not found: $LtsConfig" -Level WARN
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
# PHASE 9 - SUMMARY
# ==============================================================================
Write-Banner "LRE HOST UPGRADE COMPLETE  |  $ComponentType  |  $(hostname)"
Write-Log "Upgrade log: $Global:LogFile" -Level INFO
Write-Log "" -Level INFO
Write-Log "NEXT STEPS:" -Level WARN
Write-Log "  1. In LRE Administration > Hosts, verify this host shows version $TargetVersion." -Level INFO
Write-Log "  2. If 'Reconfigure needed', click Reconfigure Host in Administration." -Level INFO
Write-Log "  3. Verify TLS/SSL connectivity from the LRE Admin Portal." -Level INFO

if ($exitCode -eq 3010) {
    Write-Log "Installer requires a reboot (exit code 3010). Rebooting in 30 seconds ..." -Level WARN
    Start-Sleep -Seconds 30
    Restart-Computer -Force
}

Write-Host "`nUpgrade complete. Press any key to close ..." -ForegroundColor Green
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
Stop-Transcript -ErrorAction SilentlyContinue
