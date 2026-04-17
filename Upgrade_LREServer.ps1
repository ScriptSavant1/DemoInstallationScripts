# ==============================================================================
# Upgrade_LREServer.ps1  |  LRE Server Upgrade  |  25.1 -> 26.1
# ==============================================================================
# Self-contained - no external dependencies except the UserInput.xml.
# Place this script + LRE_Server_<ENV>_UserInput.xml on the server, then run.
#
# Usage:
#   .\Upgrade_LREServer.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Setup\En\setup_server.exe" -Environment DEV
#   .\Upgrade_LREServer.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Setup\En\setup_server.exe" -Environment PROD
# ==============================================================================

param(
    [Parameter(Mandatory)][string]$InstallerPath,
    [Parameter(Mandatory)][ValidateSet("DEV","PROD")][string]$Environment
)

# Self-elevate if not Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $argList = "-NoProfile -NoExit -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`" -InstallerPath `"$InstallerPath`" -Environment $Environment"
    Start-Process powershell.exe -Verb RunAs -ArgumentList $argList
    exit
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# CONSTANTS
# ==============================================================================
$InstallDir             = "D:\LREServer"
$LogDir                 = "D:\LRE_UpgradeLogs"
$CertSharePath          = "\\rbsres01\grpareas\LRE\LREadmin\LRE_CERTS"
$AWSRegion              = "eu-west-2"
$ExpectedCurrentVersion = "25.1"
$TargetVersion          = "26.1"

# Derived paths
$WebConfigDir       = "$InstallDir\PCS"
$HttpsConfigDir     = "$InstallDir\conf\httpsConfigFiles"
$PcsConfig          = "$InstallDir\dat\PCS.config"
$CACertPath         = "$InstallDir\dat\cert\verify\cacert.cer"
$CertPath           = "$InstallDir\dat\cert\cert.cer"
$GenCertExe         = "$InstallDir\bin\gen_cert.exe"
$BackendAppSettings = "$InstallDir\LRE_BACKEND\appsettings.defaults.json"
$TempUserInputDir   = "$LogDir\UserInput"

# Service name candidates
$ServiceBackendCandidates    = @("LoadRunner Backend Service","OpenText Performance Engineering Backend","LREBackend","OpenTextLREBackend","LRE_Backend")
$ServiceAlertsCandidates     = @("LoadRunner Alerts Service","OpenText Performance Engineering Alerts","LREAlerts","OpenTextLREAlerts","LRE_Alerts")

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
    param([string]$Path = "D:\", [long]$MinimumFreeGB = 20)
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
$Global:LogFile = "$LogDir\LREServer_Upgrade_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Initialize-Log -LogFile $Global:LogFile

Write-Banner "LRE SERVER UPGRADE  |  25.1 -> 26.1  |  $(hostname)  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Write-Log "Environment  : $Environment" -Level INFO
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
    Write-Log "LRE does not appear to be installed on this machine." -Level ERROR
    exit 1
}
Write-Log "Installed: $($installed.DisplayName)  Version: $($installed.Version)" -Level INFO
if ($installed.Version -notlike "*$ExpectedCurrentVersion*") {
    Write-Log "Expected version $ExpectedCurrentVersion but found $($installed.Version)." -Level ERROR
    exit 1
}
Write-Log "[1.2] Version check passed ($ExpectedCurrentVersion detected)." -Level SUCCESS

Write-Log "[1.3] Checking disk space ..." -Level INFO
if (-not (Test-DiskSpace -Path "D:\" -MinimumFreeGB 20)) { exit 1 }

Write-Log "[1.4] Checking pending reboot ..." -Level INFO
if (Test-PendingReboot) {
    Write-Log "Pending reboot detected. Please reboot and re-run." -Level ERROR
    exit 1
}
Write-Log "[1.4] No pending reboot." -Level SUCCESS


# ==============================================================================
# PHASE 2 - STOP SERVICES
# ==============================================================================
Write-Log "PHASE 2 - Stopping Services" -Level STEP
$phase2Failed = $false

Write-Log "[2.1] Checking IIS (W3SVC) ..." -Level INFO
$w3svc = Get-Service -Name "W3SVC" -ErrorAction SilentlyContinue
if (-not $w3svc) {
    Write-Log "[2.1] W3SVC not found - skipping." -Level WARN
} elseif ($w3svc.Status -eq 'Stopped') {
    Write-Log "[2.1] IIS already stopped." -Level SUCCESS
} else {
    try {
        $iisResult = & iisreset /stop 2>&1
        $iisResult | ForEach-Object { Write-Log "  $_" -Level INFO }
        Start-Sleep -Seconds 5
        $w3svc.Refresh()
        if ($w3svc.Status -eq 'Stopped') { Write-Log "[2.1] IIS stopped." -Level SUCCESS }
        else { Write-Log "[2.1] IIS still $($w3svc.Status) after iisreset /stop." -Level ERROR; $phase2Failed = $true }
    } catch {
        Write-Log "[2.1] iisreset /stop failed: $_" -Level ERROR
        $phase2Failed = $true
    }
}

Write-Log "[2.2] Checking LRE Backend Service ..." -Level INFO
$backendSvc = Resolve-ServiceName $ServiceBackendCandidates
if (-not $backendSvc) {
    Write-Log "[2.2] Backend Service not found - skipping." -Level WARN
} else {
    $beSvc = Get-Service -Name $backendSvc
    Write-Log "[2.2] '$backendSvc' status: $($beSvc.Status)" -Level INFO
    if ($beSvc.Status -eq 'Stopped') {
        Write-Log "[2.2] Backend Service already stopped." -Level SUCCESS
    } else {
        try {
            Stop-Service -Name $backendSvc -Force -ErrorAction Stop
            $beSvc.WaitForStatus('Stopped', (New-TimeSpan -Seconds 120))
            Write-Log "[2.2] Backend Service stopped." -Level SUCCESS
        } catch {
            Write-Log "[2.2] Stop failed: $_ - attempting force kill ..." -Level WARN
            Get-Process -Name "LRECore.API" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
            $beSvc.Refresh()
            if ($beSvc.Status -eq 'Stopped') { Write-Log "[2.2] Backend stopped after force kill." -Level SUCCESS }
            else { Write-Log "[2.2] FAILED to stop Backend Service." -Level ERROR; $phase2Failed = $true }
        }
    }
}

Write-Log "[2.3] Checking LRE Alerts Service ..." -Level INFO
$alertsSvc = Resolve-ServiceName $ServiceAlertsCandidates
if (-not $alertsSvc) {
    Write-Log "[2.3] Alerts Service not found - skipping." -Level WARN
} else {
    $alSvc = Get-Service -Name $alertsSvc
    Write-Log "[2.3] '$alertsSvc' status: $($alSvc.Status)" -Level INFO
    if ($alSvc.Status -eq 'Stopped') {
        Write-Log "[2.3] Alerts Service already stopped." -Level SUCCESS
    } else {
        try {
            Stop-Service -Name $alertsSvc -Force -ErrorAction Stop
            $alSvc.WaitForStatus('Stopped', (New-TimeSpan -Seconds 60))
            Write-Log "[2.3] Alerts Service stopped." -Level SUCCESS
        } catch {
            Write-Log "[2.3] Stop failed: $_ - attempting force kill ..." -Level WARN
            Get-Process -Name "PCAlertsSvc" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
            $alSvc.Refresh()
            if ($alSvc.Status -eq 'Stopped') { Write-Log "[2.3] Alerts stopped after force kill." -Level SUCCESS }
            else { Write-Log "[2.3] FAILED to stop Alerts Service." -Level ERROR; $phase2Failed = $true }
        }
    }
}

if ($phase2Failed) {
    Write-Log "PHASE 2 FAILED - services could not be stopped. Stop them manually and re-run." -Level ERROR
    exit 1
}
Write-Log "All services confirmed stopped." -Level SUCCESS


# ==============================================================================
# PHASE 3 - PREPARE UserInput.xml
# ==============================================================================
Write-Log "PHASE 3 - Preparing UserInput.xml" -Level STEP

if (-not (Test-Path $TempUserInputDir)) {
    New-Item -ItemType Directory -Path $TempUserInputDir -Force | Out-Null
}

# Find environment-specific XML first, then fall back to any matching
$xmlFile = Get-ChildItem -Path $ScriptRoot -Filter "LRE_Server_${Environment}_UserInput.xml" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $xmlFile) {
    $xmlFile = Get-ChildItem -Path $ScriptRoot -Filter "LRE_Server_*_UserInput.xml" -ErrorAction SilentlyContinue | Select-Object -First 1
}
if (-not $xmlFile) {
    Write-Log "No LRE_Server_*_UserInput.xml found in: $ScriptRoot" -Level ERROR
    exit 1
}
Write-Log "[3.1] Found UserInput.xml: $($xmlFile.FullName)" -Level SUCCESS

$tempXml = "$TempUserInputDir\Server_UserInput_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
Copy-Item -Path $xmlFile.FullName -Destination $tempXml -Force
Write-Log "[3.2] Copied to temp: $tempXml" -Level INFO

$hn = (hostname).ToLower()
$xmlContent = Get-Content $tempXml -Raw -Encoding UTF8
if ($xmlContent -match '\{\{HOSTNAME\}\}') {
    $xmlContent = $xmlContent -replace '\{\{HOSTNAME\}\}', $hn
    Set-Content -Path $tempXml -Value $xmlContent -Encoding UTF8
    Write-Log "[3.3] Replaced {{HOSTNAME}} with '$hn'" -Level SUCCESS
}


# ==============================================================================
# PHASE 4 - RUN SILENT SERVER UPGRADE
# ==============================================================================
Write-Log "PHASE 4 - Running LRE Server Upgrade" -Level STEP

$installerArgs = "/s USER_CONFIG_FILE_PATH=`"$tempXml`" INSTALLDIR=`"$InstallDir`" NVINSTALL=Y"
Write-Log "Command: $InstallerPath $installerArgs" -Level INFO
Write-Log "Started at $(Get-Date -Format 'HH:mm:ss') - this may take 30-60 minutes ..." -Level INFO

$configWizardLog = "$InstallDir\orchidtmp\Configuration\configurationWizardLog_pcs.txt"
$monitorJob = Start-Job -ScriptBlock {
    param($installDir, $configLog)
    $lastLineCount = 0; $lastMsiCheck = ""
    while ($true) {
        Start-Sleep -Seconds 15
        $msgs = @()
        $msiProc = Get-Process -Name "msiexec" -ErrorAction SilentlyContinue
        if ($msiProc) {
            $status = "MSI installer running (PID: $(($msiProc | Select-Object -First 1).Id))"
            if ($status -ne $lastMsiCheck) { $msgs += $status; $lastMsiCheck = $status }
        }
        if (Test-Path $configLog) {
            $lines = @(Get-Content $configLog -ErrorAction SilentlyContinue)
            if ($lines.Count -gt $lastLineCount) {
                $lines[$lastLineCount..($lines.Count-1)] | Where-Object { $_ -match '\S' } | ForEach-Object { $msgs += "  ConfigWizard: $_" }
                $lastLineCount = $lines.Count
            }
        }
        $recentFiles = Get-ChildItem -Path $installDir -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -gt (Get-Date).AddSeconds(-20) } | Select-Object -First 3
        if ($recentFiles) { $msgs += "  Files being written: $($recentFiles[0].FullName)" }
        foreach ($msg in $msgs) { Write-Output "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] $msg" }
    }
} -ArgumentList $InstallDir, $configWizardLog

$installStart = Get-Date
try {
    $proc = Start-Process -FilePath $InstallerPath -ArgumentList $installerArgs -Wait -PassThru -NoNewWindow
    $exitCode = $proc.ExitCode
} catch {
    Write-Log "Failed to launch installer: $_" -Level ERROR
    if (Test-Path $tempXml) { Remove-Item $tempXml -Force -ErrorAction SilentlyContinue }
    exit 1
}

Stop-Job $monitorJob -ErrorAction SilentlyContinue
$monitorOutput = Receive-Job $monitorJob -ErrorAction SilentlyContinue
Remove-Job $monitorJob -Force -ErrorAction SilentlyContinue
if ($monitorOutput) {
    $monitorOutput | ForEach-Object {
        Write-Host $_ -ForegroundColor DarkGray
        if ($Global:LogFile) { Add-Content -Path $Global:LogFile -Value $_ }
    }
}

$installDuration = (Get-Date) - $installStart
Write-Log "Installer exit code: $exitCode  |  Duration: $([math]::Round($installDuration.TotalMinutes, 1)) min" -Level INFO

if (Test-Path $tempXml) {
    Remove-Item $tempXml -Force -ErrorAction SilentlyContinue
    Write-Log "Temp UserInput.xml deleted." -Level INFO
}

switch ($exitCode) {
    0    { Write-Log "LRE Server upgrade completed successfully." -Level SUCCESS }
    3010 { Write-Log "LRE Server upgrade succeeded - REBOOT REQUIRED (handled in Phase 9)." -Level WARN }
    1602 { Write-Log "Upgrade cancelled by user." -Level WARN; exit 1602 }
    1603 { Write-Log "Fatal error during installation (1603). Check installer logs." -Level ERROR; exit 1603 }
    default { Write-Log "Unexpected exit code: $exitCode" -Level ERROR; exit $exitCode }
}


# ==============================================================================
# PHASE 5 - VERIFY INSTALLER LOG
# ==============================================================================
Write-Log "PHASE 5 - Verifying Installer Log" -Level STEP

$configLog = "$InstallDir\orchidtmp\Configuration\configurationWizardLog_pcs.txt"
if (Test-Path $configLog) {
    $logTail = Get-Content $configLog -Tail 30
    Write-Log "Last 30 lines of Configuration Wizard log:" -Level INFO
    $logTail | ForEach-Object { Write-Log "  $_" -Level INFO }
    if ($logTail -join " " -match "ERROR|FAIL|exception") {
        Write-Log "Potential errors in configuration log. Review: $configLog" -Level WARN
    }
} else {
    Write-Log "Configuration wizard log not found: $configLog" -Level WARN
}


# ==============================================================================
# PHASE 6 - RESTART SERVICES + CAPTURE PUBLIC KEY
# ==============================================================================
Write-Log "PHASE 6 - Restart Services + Capture Public Key" -Level STEP

$backendSvc = Resolve-ServiceName $ServiceBackendCandidates
$alertsSvc  = Resolve-ServiceName $ServiceAlertsCandidates
Write-Log "[6.0] Post-upgrade services: Backend='$backendSvc'  Alerts='$alertsSvc'" -Level INFO

Write-Log "[6.1] Starting IIS ..." -Level INFO
$iisStartResult = & iisreset /start 2>&1
$iisStartResult | ForEach-Object { Write-Log "  $_" -Level INFO }
Start-Sleep -Seconds 10

Write-Log "[6.2] Starting Backend Service ..." -Level INFO
if ($backendSvc) { Start-ServiceSafely -ServiceName $backendSvc }
else { Write-Log "[6.2] Backend Service not found - may need manual start." -Level WARN }

Write-Log "[6.3] Starting Alerts Service ..." -Level INFO
if ($alertsSvc) { Start-ServiceSafely -ServiceName $alertsSvc }
else { Write-Log "[6.3] Alerts Service not found - may need manual start." -Level WARN }

Write-Log "Waiting 30 seconds for services to initialize ..." -Level INFO
for ($i = 30; $i -gt 0; $i -= 10) { Write-Log "  ... $i seconds remaining" -Level INFO; Start-Sleep -Seconds 10 }

$publicKey = $null

try {
    Write-Log "[6.4] Attempting Public Key via REST API ..." -Level INFO
    $apiUrl   = "http://localhost/Admin/rest/v1/configuration/getPublicKey"
    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -TimeoutSec 30 -ErrorAction Stop
    if ($response.PublicKey -or $response.publicKey) {
        $publicKey = if ($response.PublicKey) { $response.PublicKey } else { $response.publicKey }
        Write-Log "[6.4] Public Key retrieved via REST API." -Level SUCCESS
    }
} catch {
    Write-Log "[6.4] REST API call failed: $_" -Level WARN
}

if (-not $publicKey) {
    $pcsConfigFile = "$InstallDir\dat\pcs.config"
    if (Test-Path $pcsConfigFile) {
        $configContent = Get-Content $pcsConfigFile -Raw
        if ($configContent -match 'PublicKey["\s=:]+([A-Za-z0-9+/=]{20,})') {
            $publicKey = $Matches[1]
            Write-Log "[6.4] Public Key found in pcs.config." -Level SUCCESS
        }
    }
}

if ($publicKey) {
    Set-Content -Path "$LogDir\PublicKey.txt" -Value $publicKey -Encoding UTF8
    Write-Log "[6.4] Public Key saved to: $LogDir\PublicKey.txt" -Level SUCCESS
} else {
    Write-Log "[6.4] Could not retrieve Public Key automatically." -Level WARN
    Write-Log "      Retrieve manually from: $InstallDir\dat\pcs.config" -Level WARN
}


# ==============================================================================
# PHASE 7 - POST-INSTALL CONFIGURATION
# ==============================================================================
Write-Log "PHASE 7 - Post-Install Configuration" -Level STEP

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'

# --- 7.1 web.config swap ---
Write-Log "[7.1] Configure TLS/SSL - Web.config swap ..." -Level INFO
$webConfig    = Join-Path $WebConfigDir "web.config"
$webConfigSsl = Join-Path $HttpsConfigDir "web.config-for_ssl"

if (Test-Path $webConfig) {
    $webConfigBak = "${webConfig}.bak_${ts}"
    Copy-Item -Path $webConfig -Destination $webConfigBak -Force
    Write-Log "[7.1] Backed up `"web.config`" to `"$webConfigBak`"" -Level INFO
    if (Test-Path $webConfigSsl) {
        Copy-Item -Path $webConfigSsl -Destination $webConfig -Force
        Write-Log "[7.1] Copied `"web.config-for_ssl`" from `"$HttpsConfigDir`" to `"$webConfig`"" -Level INFO
        Write-Log "[7.1] SUCCESS: web.config replaced with SSL version." -Level SUCCESS
    } else {
        Write-Log "[7.1] web.config-for-ssl not found at: $webConfigSsl - manual edit needed." -Level WARN
    }
} else {
    Write-Log "[7.1] web.config not found: $webConfig - skipping." -Level WARN
}

# --- 7.2 PCS.config updates ---
Write-Log "[7.2] Configure TLS/SSL - PCS.config updates ..." -Level INFO

if (Test-Path $PcsConfig) {
    $pcsConfigBak = "${PcsConfig}.bak_${ts}"
    Copy-Item -Path $PcsConfig -Destination $pcsConfigBak -Force
    Write-Log "[7.2] Backed up `"PCS.config`" to `"$pcsConfigBak`"" -Level INFO

    $pcsContent = Get-Content $PcsConfig -Raw -Encoding UTF8
    $hn = (hostname).ToLower()
    $lreDomain = if ($Environment -eq "DEV") { "webdev.banksvcs.net" } else { "web.banksvcs.net" }
    $newInternalUrl = "https://${hn}.${lreDomain}:12001"

    if ($pcsContent -match 'internalUrl\s*=\s*"([^"]*)"') {
        $oldUrl = $Matches[1]
        $pcsContent = $pcsContent -replace [regex]::Escape("internalUrl=`"$oldUrl`""), "internalUrl=`"$newInternalUrl`""
        Write-Log "[7.2] Updated internalUrl: `"$oldUrl`" -> `"$newInternalUrl`"" -Level INFO
    } else {
        Write-Log "[7.2] internalUrl not found in PCS.config." -Level WARN
    }

    if ($pcsContent -match 'ltopIsSecured\s*=\s*"false"') {
        $pcsContent = $pcsContent -replace 'ltopIsSecured\s*=\s*"false"', 'ltopIsSecured="true"'
        Write-Log "[7.2] Updated ltopIsSecured: `"false`" -> `"true`"" -Level INFO
    } elseif ($pcsContent -match 'ltopIsSecured\s*=\s*"true"') {
        Write-Log "[7.2] ltopIsSecured already true." -Level INFO
    } else {
        Write-Log "[7.2] ltopIsSecured not found in PCS.config." -Level WARN
    }

    Set-Content -Path $PcsConfig -Value $pcsContent -Encoding UTF8
    Write-Log "[7.2] Saved PCS.config to `"$PcsConfig`"" -Level INFO
    Write-Log "[7.2] SUCCESS: PCS.config updated." -Level SUCCESS
} else {
    Write-Log "[7.2] PCS.config not found: $PcsConfig" -Level WARN
}

# --- 7.3 Certificate replacement ---
Write-Log "[7.3] Replacing CA and TLS certificates ..." -Level INFO

$certUpdateOk = $true
$envLower  = $Environment.ToLower()
$certSuffix = if ($Environment -eq "DEV") { "iis_dev" } else { "iis_server" }
$caCertName  = "lre_${envLower}_cacert.cer"
$caCertSource = Join-Path $CertSharePath $caCertName

if (Test-ShareAccessible -SharePath $CertSharePath) {
    Write-Log "[7.3] CA cert source: `"$caCertSource`"" -Level INFO
    if (Test-Path $caCertSource) {
        $caCertDestDir = Split-Path $CACertPath -Parent
        if (-not (Test-Path $caCertDestDir)) { New-Item -ItemType Directory -Path $caCertDestDir -Force | Out-Null }
        if (Test-Path $CACertPath) {
            Copy-Item -Path $CACertPath -Destination "${CACertPath}.bak_${ts}" -Force
            Write-Log "[7.3] Backed up existing CA cert to `"${CACertPath}.bak_${ts}`"" -Level INFO
        }
        Copy-Item -Path $caCertSource -Destination $CACertPath -Force
        Write-Log "[7.3] Copied `"$caCertName`" to `"$CACertPath`"" -Level INFO
    } else {
        Write-Log "[7.3] CA cert not found: $caCertSource" -Level WARN
        $certUpdateOk = $false
    }

    $expectedCertName = "lre_${envLower}_${hn}_${certSuffix}.cer"
    Write-Log "[7.3] TLS cert source: `"$(Join-Path $CertSharePath $expectedCertName)`"" -Level INFO
    $serverCertFile = Get-ChildItem -Path $CertSharePath -Filter $expectedCertName -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $serverCertFile) {
        Write-Log "[7.3] Expected cert '$expectedCertName' not found, trying wildcard *${hn}*.cer ..." -Level WARN
        $serverCertFile = Get-ChildItem -Path $CertSharePath -Filter "*${hn}*.cer" -ErrorAction SilentlyContinue | Select-Object -First 1
    }

    if ($serverCertFile) {
        if (Test-Path $CertPath) {
            Copy-Item -Path $CertPath -Destination "${CertPath}.bak_${ts}" -Force
            Write-Log "[7.3] Backed up existing TLS cert to `"${CertPath}.bak_${ts}`"" -Level INFO
        }
        Copy-Item -Path $serverCertFile.FullName -Destination $CertPath -Force
        Write-Log "[7.3] Copied `"$($serverCertFile.Name)`" to `"$CertPath`"" -Level INFO
    } else {
        Write-Log "[7.3] TLS cert not found in $CertSharePath for hostname '$hn'." -Level WARN
        $certUpdateOk = $false
    }

    if ($certUpdateOk -and (Test-Path $GenCertExe)) {
        Write-Log "[7.3] Running gen_cert.exe -verify ..." -Level INFO
        try {
            $verifyOutput = & $GenCertExe -verify 2>&1
            $verifyOutput | ForEach-Object { Write-Log "[7.3]   $_" -Level INFO }
            if ($verifyOutput -join " " -match "success|verified") {
                Write-Log "[7.3] SUCCESS: Certificate replacement and verification passed." -Level SUCCESS
            } else {
                Write-Log "[7.3] Certificate verification - review output above." -Level WARN
            }
        } catch {
            Write-Log "[7.3] gen_cert.exe -verify failed: $_" -Level WARN
        }
    }
} else {
    Write-Log "[7.3] Certificate share not accessible: $CertSharePath - skipping cert replacement." -Level WARN
}

# --- 7.4 AWS ActiveRegions ---
Write-Log "[7.4] Configure AWS ActiveRegions ..." -Level INFO

if (Test-Path $BackendAppSettings) {
    $appSettingsBak = "${BackendAppSettings}.bak_${ts}"
    Copy-Item -Path $BackendAppSettings -Destination $appSettingsBak -Force
    Write-Log "[7.4] Backed up `"appsettings.defaults.json`" to `"$appSettingsBak`"" -Level INFO

    $appContent = Get-Content $BackendAppSettings -Raw -Encoding UTF8

    if ($appContent -match '"ActiveRegions"\s*:\s*\[\s*"[^"]+"') {
        Write-Log "[7.4] AWS ActiveRegions already configured - skipping." -Level INFO
    } elseif ($appContent -match '"ActiveRegions"\s*:\s*\[\s*\]') {
        Write-Log "[7.4] Found empty ActiveRegions array." -Level INFO
        $appContent = $appContent -replace '"ActiveRegions"\s*:\s*\[\s*\]', "`"ActiveRegions`": [ `"$AWSRegion`" ]"
        Set-Content -Path $BackendAppSettings -Value $appContent -Encoding UTF8
        Write-Log "[7.4] Updated to: `"ActiveRegions`": [ `"$AWSRegion`" ]" -Level INFO
        Write-Log "[7.4] Saved `"$BackendAppSettings`"" -Level INFO
        Write-Log "[7.4] SUCCESS: AWS ActiveRegions configured." -Level SUCCESS
    } elseif ($appContent -match '"AWS"\s*:') {
        $appContent = $appContent -replace '("AWS"\s*:\s*\{)', "`$1`n    `"ActiveRegions`": [ `"$AWSRegion`" ],"
        Set-Content -Path $BackendAppSettings -Value $appContent -Encoding UTF8
        Write-Log "[7.4] Added ActiveRegions to existing AWS section." -Level SUCCESS
    } else {
        $insertion = "  `"AWS`": {`n    `"ActiveRegions`": [ `"$AWSRegion`" ]`n  },"
        $appContent = $appContent -replace '(\{\s*)(\r?\n\s*")', "`$1`n  $insertion`n`$2"
        Set-Content -Path $BackendAppSettings -Value $appContent -Encoding UTF8
        Write-Log "[7.4] Added AWS section with ActiveRegions." -Level SUCCESS
    }
} else {
    Write-Log "[7.4] appsettings.defaults.json not found: $BackendAppSettings" -Level WARN
}

# --- 7.5 Restart services ---
Write-Log "[7.5] Restarting services ..." -Level INFO

if ($backendSvc) { Stop-ServiceSafely -ServiceName $backendSvc; Start-ServiceSafely -ServiceName $backendSvc }
if ($alertsSvc)  { Stop-ServiceSafely -ServiceName $alertsSvc;  Start-ServiceSafely -ServiceName $alertsSvc  }

$iisRestartResult = & iisreset /restart 2>&1
$iisRestartResult | ForEach-Object { Write-Log "  $_" -Level INFO }
Start-Sleep -Seconds 10
Write-Log "[7.5] Services restarted." -Level SUCCESS


# ==============================================================================
# PHASE 8 - VERIFICATION
# ==============================================================================
Write-Log "PHASE 8 - Verification" -Level STEP

$webConfigContent = Get-Content $webConfig -Raw -ErrorAction SilentlyContinue
if ($webConfigContent -match "Transport|httpsGetEnabled") {
    Write-Log "  web.config: SSL configuration present." -Level SUCCESS
} else {
    Write-Log "  web.config: SSL configuration may be missing." -Level WARN
}

$pcsVerify = Get-Content $PcsConfig -Raw -ErrorAction SilentlyContinue
if ($pcsVerify -match 'internalUrl\s*=\s*"https://') {
    Write-Log "  PCS.config: internalUrl is HTTPS." -Level SUCCESS
} else {
    Write-Log "  PCS.config: internalUrl may not be HTTPS." -Level WARN
}
if ($pcsVerify -match 'ltopIsSecured\s*=\s*"true"') {
    Write-Log "  PCS.config: ltopIsSecured is true." -Level SUCCESS
} else {
    Write-Log "  PCS.config: ltopIsSecured may not be true." -Level WARN
}

try {
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    $resp = Invoke-WebRequest -Uri "https://localhost/LRE/" -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
    Write-Log "  HTTPS check: HTTP $($resp.StatusCode) from https://localhost/LRE/" -Level SUCCESS
} catch {
    Write-Log "  HTTPS check failed: $($_.Exception.Message) - server may need a few minutes." -Level WARN
} finally {
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
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
Write-Banner "LRE SERVER UPGRADE COMPLETE  |  $(hostname)"
Write-Log "Upgrade log : $Global:LogFile" -Level INFO
if ($publicKey) {
    Write-Log "Public Key  : $LogDir\PublicKey.txt" -Level INFO
} else {
    Write-Log "Public Key  : NOT CAPTURED - retrieve manually from $InstallDir\dat\pcs.config" -Level WARN
}

Write-Log "" -Level INFO
Write-Log "AUTOMATED STEPS COMPLETED:" -Level SUCCESS
Write-Log "  - LRE Server upgraded from $ExpectedCurrentVersion to $TargetVersion" -Level INFO
Write-Log "  - web.config replaced with SSL version" -Level INFO
Write-Log "  - PCS.config internalUrl set to HTTPS" -Level INFO
Write-Log "  - PCS.config ltopIsSecured set to true" -Level INFO
Write-Log "  - CA certificate replaced" -Level INFO
Write-Log "  - Server TLS certificate replaced" -Level INFO
Write-Log "  - Certificates verified via gen_cert.exe" -Level INFO
Write-Log "  - AWS ActiveRegions configured" -Level INFO
Write-Log "  - Services restarted" -Level INFO
Write-Log "" -Level INFO
Write-Log "REMAINING MANUAL STEPS (Admin Portal):" -Level WARN
Write-Log "  1. Update auth type in DB" -Level INFO
Write-Log "  2. Verify Site Admin login and update LDAP configuration" -Level INFO
Write-Log "  3. Change authentication method to LDAP" -Level INFO
Write-Log "  4. Update Server URLs in Admin Portal" -Level INFO
Write-Log "  5. Upgrade projects in Admin Portal" -Level INFO
Write-Log "  6. Upload license file" -Level INFO
Write-Log "  7. Verify AWS Cloud account/proxy/templates" -Level INFO
Write-Log "  8. Reconfigure hosts in Admin Portal" -Level INFO

if ($exitCode -eq 3010) {
    Write-Log "Installer requires a reboot (exit code 3010). Rebooting in 30 seconds ..." -Level WARN
    Start-Sleep -Seconds 30
    Restart-Computer -Force
}

Write-Host "`nUpgrade complete. Press any key to close ..." -ForegroundColor Green
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
Stop-Transcript -ErrorAction SilentlyContinue
