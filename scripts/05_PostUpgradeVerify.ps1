# ==============================================================================
# 05_PostUpgradeVerify.ps1
# LRE Post-Upgrade Verification
# ------------------------------------------------------------------------------
# Run this on EACH machine after all upgrade scripts have completed.
# Produces a comprehensive health report without making any changes.
#
# Usage (elevated PowerShell):
#   .\05_PostUpgradeVerify.ps1
#
# To specify component type explicitly (for clearer reporting):
#   .\05_PostUpgradeVerify.ps1 -ComponentType Server
#   .\05_PostUpgradeVerify.ps1 -ComponentType Controller
#   .\05_PostUpgradeVerify.ps1 -ComponentType DataProcessor
#   .\05_PostUpgradeVerify.ps1 -ComponentType LoadGenerator
# ==============================================================================
#Requires -RunAsAdministrator

param(
    [ValidateSet("Server","Controller","DataProcessor","LoadGenerator","Auto")]
    [string]$ComponentType = "Auto"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'   # Non-blocking - report all issues

$ScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
. "$ScriptRoot\Common-Functions.ps1"
. "$ScriptRoot\..\config\upgrade_config.ps1"

$Global:LogFile = "$LogDir\PostUpgradeVerify_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Initialize-Log -LogFile $Global:LogFile

Write-Banner "POST-UPGRADE VERIFICATION  |  $(hostname)  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
function Add-Result {
    param([string]$Check, [bool]$Passed, [string]$Detail)
    $status = if ($Passed) { "PASS" } else { "FAIL" }
    $level  = if ($Passed) { "SUCCESS" } else { "ERROR" }
    Write-Log "[$status] $Check - $Detail" -Level $level
    $results.Add([PSCustomObject]@{ Check=$Check; Status=$status; Detail=$Detail })
}

# ==============================================================================
# AUTO-DETECT COMPONENT TYPE
# ==============================================================================
if ($ComponentType -eq "Auto") {
    $detected = Get-LREInstalledVersion
    if ($detected) {
        if ($detected.DisplayName -like "*Enterprise Performance Engineering*" -and
            $detected.DisplayName -notlike "*Host*") {
            $ComponentType = "Server"
        } elseif ($detected.DisplayName -like "*Host*" -or
                  $detected.DisplayName -like "*Professional Performance*") {
            # Can't distinguish Controller/DP/LG-Host by install alone - use "Controller" as default label
            $ComponentType = "Controller"
        } elseif ($detected.DisplayName -like "*OneLG*" -or $detected.DisplayName -like "*LoadRunner*") {
            $ComponentType = "LoadGenerator"
        }
    }
    Write-Log "Auto-detected component type: $ComponentType" -Level INFO
}

Write-Log "Verifying as: $ComponentType" -Level INFO

# ==============================================================================
# CHECK 1 — Installed Version = 26.1
# ==============================================================================
Write-Log "--- Check: Installed Version" -Level STEP
$installed = Get-LREInstalledVersion
if ($installed) {
    $versionOk = $installed.Version -like "*$TargetVersion*"
    Add-Result "Version is $TargetVersion" $versionOk "$($installed.DisplayName) v$($installed.Version)"
} else {
    # Broaden search for OneLG
    $allProducts = Get-ChildItem -Path @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ) -ErrorAction SilentlyContinue |
    ForEach-Object { Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue } |
    Where-Object { $_.DisplayName -like "*LoadRunner*" -or $_.DisplayName -like "*Performance*" -or $_.DisplayName -like "*OneLG*" }

    if ($allProducts) {
        $p = $allProducts | Select-Object -First 1
        $versionOk = $p.DisplayVersion -like "*$TargetVersion*"
        Add-Result "Version is $TargetVersion" $versionOk "$($p.DisplayName) v$($p.DisplayVersion)"
    } else {
        Add-Result "Version is $TargetVersion" $false "LRE product not found in registry"
    }
}

# ==============================================================================
# CHECK 2 — Services Are Running
# ==============================================================================
Write-Log "--- Check: Services" -Level STEP

if ($ComponentType -eq "Server") {
    # IIS
    $iisStatus = Get-ServiceStatus -ServiceName $ServiceIIS
    Add-Result "Service: IIS (W3SVC)" ($iisStatus -eq "Running") "Status: $iisStatus"

    # Backend
    $backendSvc = Resolve-ServiceName $ServiceBackend
    if ($backendSvc) {
        $st = Get-ServiceStatus -ServiceName $backendSvc
        Add-Result "Service: Backend ($backendSvc)" ($st -eq "Running") "Status: $st"
    } else {
        Add-Result "Service: Backend" $false "Service not found (check $ServiceBackend)"
    }

    # Alerts
    $alertsSvc = Resolve-ServiceName $ServiceAlerts
    if ($alertsSvc) {
        $st = Get-ServiceStatus -ServiceName $alertsSvc
        Add-Result "Service: Alerts ($alertsSvc)" ($st -eq "Running") "Status: $st"
    } else {
        Add-Result "Service: Alerts" $false "Service not found (check $ServiceAlerts)"
    }
}

if ($ComponentType -in @("Controller","DataProcessor","LoadGenerator")) {
    # LoadRunner Agent
    $agentSvc = Resolve-ServiceName $ServiceAgent
    if ($agentSvc) {
        $st = Get-ServiceStatus -ServiceName $agentSvc
        Add-Result "Service: LoadRunner Agent ($agentSvc)" ($st -eq "Running") "Status: $st"
    } else {
        Add-Result "Service: LoadRunner Agent" $false "Service not found - check service name in upgrade_config.ps1"
    }

    # Remote Management Agent
    $remoteSvc = Resolve-ServiceName $ServiceRemoteMgmt
    if ($remoteSvc) {
        $st = Get-ServiceStatus -ServiceName $remoteSvc
        Add-Result "Service: Remote Mgmt Agent ($remoteSvc)" ($st -eq "Running") "Status: $st"
    } else {
        Add-Result "Service: Remote Mgmt Agent" $false "Service not found - check service name in upgrade_config.ps1"
    }
}

# ==============================================================================
# CHECK 3 — Web Accessibility (Server only)
# ==============================================================================
if ($ComponentType -eq "Server") {
    Write-Log "--- Check: LRE Web Application" -Level STEP
    $urls = @(
        "https://$IisSecureHostName/LRE/",
        "http://localhost/LRE/"
    )
    $webOk = $false
    foreach ($url in $urls) {
        try {
            Write-Log "Testing: $url" -Level INFO
            # Ignore SSL errors for self-signed certs in test environment
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
            $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
            if ($response.StatusCode -in @(200, 302, 301)) {
                $webOk = $true
                Write-Log "HTTP $($response.StatusCode) from $url" -Level SUCCESS
                break
            }
        } catch {
            Write-Log "Request to $url failed: $($_.Exception.Message)" -Level WARN
        }
    }
    Add-Result "LRE Web App Accessible" $webOk "Tested: $($urls -join ' | ')"
}

# ==============================================================================
# CHECK 4 — Config File Consistency
# ==============================================================================
Write-Log "--- Check: Config File Service User Consistency" -Level STEP

if ($ComponentType -eq "Server") {
    $pcsConfig = "$ServerInstallDir\dat\pcs.config"
    if (Test-Path $pcsConfig) {
        $content = Get-Content $pcsConfig -Raw
        $userOk  = $content -match [regex]::Escape($SystemUserName)
        Add-Result "pcs.config user consistency" $userOk "Expected user '$SystemUserName' in $pcsConfig"
    } else {
        Add-Result "pcs.config exists" $false "Not found: $pcsConfig"
    }

    $appSettings = "$FileSystemRoot\system_config\appsettings.json"
    if (Test-Path $appSettings) {
        $content = Get-Content $appSettings -Raw
        $userOk  = $content -match [regex]::Escape($SystemUserName)
        Add-Result "appsettings.json user consistency" $userOk "Expected user '$SystemUserName' in $appSettings"
    } else {
        Add-Result "appsettings.json exists" $false "Not found: $appSettings"
    }
}

if ($ComponentType -in @("Controller","DataProcessor")) {
    $ltsConfig = "$HostInstallDir\dat\lts.config"
    if (Test-Path $ltsConfig) {
        $content = Get-Content $ltsConfig -Raw
        $userOk  = $content -match [regex]::Escape($SystemUserName)
        Add-Result "lts.config user consistency" $userOk "Expected user '$SystemUserName' in $ltsConfig"
    } else {
        Add-Result "lts.config exists" $false "Not found: $ltsConfig"
    }
}

if ($ComponentType -eq "LoadGenerator") {
    $oneLGConfig = "$OneLGInstallDir\dat\lts.config"
    if (-not (Test-Path $oneLGConfig)) {
        $oneLGConfig = "$OneLGInstallDir\dat\br_lnch_server.cfg"
    }
    if (Test-Path $oneLGConfig) {
        Add-Result "OneLG config file exists" $true $oneLGConfig
    } else {
        Add-Result "OneLG config file exists" $false "Expected at: $oneLGConfig"
    }
}

# ==============================================================================
# CHECK 5 — No Pending Reboot
# ==============================================================================
Write-Log "--- Check: Pending Reboot" -Level STEP
$pendingReboot = Test-PendingReboot
Add-Result "No Pending Reboot" (-not $pendingReboot) `
    $(if ($pendingReboot) { "REBOOT REQUIRED - reboot before putting machine back in service" } else { "No pending reboot" })

# ==============================================================================
# CHECK 6 — Temp UserInput.xml Cleaned Up (security)
# ==============================================================================
Write-Log "--- Check: Temp Files Cleaned Up" -Level STEP
$tempXmlFiles = Get-ChildItem -Path $TempUserInputDir -Filter "*.xml" -ErrorAction SilentlyContinue
$tempClean = ($null -eq $tempXmlFiles -or $tempXmlFiles.Count -eq 0)
Add-Result "Temp UserInput.xml files cleaned up" $tempClean `
    $(if ($tempClean) { "No temp XML files found" } else { "WARNING: $($tempXmlFiles.Count) temp XML file(s) contain credentials - delete them: $TempUserInputDir" })

# ==============================================================================
# CHECK 7 — Event Log (recent errors)
# ==============================================================================
Write-Log "--- Check: Recent Event Log Errors" -Level STEP
try {
    $since = (Get-Date).AddHours(-2)
    $errors = Get-EventLog -LogName Application -EntryType Error -After $since -ErrorAction SilentlyContinue |
              Where-Object { $_.Source -like "*Performance*" -or $_.Source -like "*LRE*" -or
                             $_.Source -like "*LoadRunner*" -or $_.Source -like "*OpenText*" } |
              Select-Object -First 5
    if ($errors -and $errors.Count -gt 0) {
        Write-Log "Recent application errors found:" -Level WARN
        $errors | ForEach-Object { Write-Log "  [$($_.TimeGenerated)] $($_.Source): $($_.Message.Substring(0,[Math]::Min(200,$_.Message.Length)))" -Level WARN }
        Add-Result "No Recent LRE Event Log Errors" $false "$($errors.Count) error(s) in last 2 hours - review Event Viewer"
    } else {
        Add-Result "No Recent LRE Event Log Errors" $true "No LRE-related errors in last 2 hours"
    }
} catch {
    Add-Result "Event Log Check" $true "Could not read event log: $_"
}

# ==============================================================================
# SUMMARY
# ==============================================================================
Write-Banner "VERIFICATION SUMMARY  |  $ComponentType  |  $(hostname)"

$passCount = ($results | Where-Object { $_.Status -eq 'PASS' }).Count
$failCount = ($results | Where-Object { $_.Status -eq 'FAIL' }).Count

$results | Format-Table -AutoSize -Property Check, Status, Detail | Out-String | Write-Host

if ($failCount -eq 0) {
    Write-Host "`n  RESULT: ALL CHECKS PASSED ($passCount/$($results.Count))" -ForegroundColor Green
    Write-Host "  $(hostname) [$ComponentType] is successfully upgraded to $TargetVersion.`n" -ForegroundColor Green
    Write-Log "RESULT: ALL CHECKS PASSED. Upgrade verified for $ComponentType on $(hostname)." -Level SUCCESS
} else {
    Write-Host "`n  RESULT: $failCount CHECK(S) FAILED" -ForegroundColor Red
    Write-Host "  Review the FAIL items above and resolve before putting this machine back in service.`n" -ForegroundColor Red
    Write-Log "RESULT: $failCount FAILED for $ComponentType on $(hostname)." -Level ERROR
}

Write-Host "  Full log: $Global:LogFile`n" -ForegroundColor Cyan
Stop-Transcript -ErrorAction SilentlyContinue
