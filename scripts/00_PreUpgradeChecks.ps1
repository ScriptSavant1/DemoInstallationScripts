# ==============================================================================
# 00_PreUpgradeChecks.ps1
# LRE Upgrade Pre-flight Diagnostic
# ------------------------------------------------------------------------------
# Run this on EVERY machine (Server, Controller, DP, LG) BEFORE starting any
# upgrade. Produces a pass/fail report without making any changes.
#
# Usage:
#   Right-click -> "Run with PowerShell" (as Administrator)
#   OR from an elevated PowerShell prompt:
#   .\00_PreUpgradeChecks.ps1
# ==============================================================================
#Requires -RunAsAdministrator

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Dot-source shared helpers and config ---
$ScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
. "$ScriptRoot\Common-Functions.ps1"
. "$ScriptRoot\..\config\upgrade_config.ps1"

# --- Init log ---
$Global:LogFile = "$LogDir\PreUpgradeChecks_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Initialize-Log -LogFile $Global:LogFile

Write-Banner "LRE Pre-Upgrade Checks  |  Machine: $(hostname)  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

function Add-Result {
    param([string]$Check, [bool]$Passed, [string]$Detail)
    $status = if ($Passed) { "PASS" } else { "FAIL" }
    $level  = if ($Passed) { "SUCCESS" } else { "ERROR" }
    Write-Log "[$status] $Check - $Detail" -Level $level
    $results.Add([PSCustomObject]@{
        Check  = $Check
        Status = $status
        Detail = $Detail
    })
}

# ==============================================================================
# CHECK 1 — Running as Administrator
# ==============================================================================
Write-Log "--- Check: Administrator Privileges" -Level STEP
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
Add-Result "Administrator Privileges" $isAdmin $(if ($isAdmin) { "Running as administrator" } else { "NOT running as administrator - re-run elevated" })

# ==============================================================================
# CHECK 2 — OS Version (Windows)
# ==============================================================================
Write-Log "--- Check: Operating System" -Level STEP
$os = Get-WmiObject Win32_OperatingSystem
$osInfo = "$($os.Caption) - Build $($os.BuildNumber)"
$isWin  = $os.Caption -like "*Windows*"
Add-Result "Operating System" $isWin $osInfo

# ==============================================================================
# CHECK 3 — System Locale (must be English for smooth install)
# ==============================================================================
Write-Log "--- Check: System Locale" -Level STEP
$locale = (Get-WinSystemLocale).Name
$localeOk = $locale -like "en-*"
Add-Result "System Locale" $localeOk "Locale: $locale $(if(-not $localeOk){'(non-English locales may cause issues - see guide section: Windows system locale considerations)'})"

# ==============================================================================
# CHECK 4 — Pending Reboot
# ==============================================================================
Write-Log "--- Check: Pending Reboot" -Level STEP
$pendingReboot = Test-PendingReboot
Add-Result "No Pending Reboot" (-not $pendingReboot) $(if ($pendingReboot) { "PENDING REBOOT DETECTED - reboot before upgrading" } else { "No pending reboot" })

# ==============================================================================
# CHECK 5 — Installed LRE Version
# ==============================================================================
Write-Log "--- Check: Installed LRE Version" -Level STEP
$installed = Get-LREInstalledVersion
if ($installed) {
    $versionMatch = $installed.Version -like "*$ExpectedCurrentVersion*"
    Add-Result "LRE Version is $ExpectedCurrentVersion" $versionMatch "Found: $($installed.DisplayName) v$($installed.Version)"
    Write-Log "  Install Dir: $($installed.InstallDir)" -Level INFO
} else {
    Add-Result "LRE Version is $ExpectedCurrentVersion" $false "LRE does not appear to be installed on this machine"
}

# ==============================================================================
# CHECK 6 — Disk Space (minimum 20 GB free on C:\)
# ==============================================================================
Write-Log "--- Check: Disk Space" -Level STEP
$diskOk = Test-DiskSpace -Path "C:\" -MinimumFreeGB 20
Add-Result "Disk Space (>=20 GB free)" $diskOk "See log for detail"

# ==============================================================================
# CHECK 7 — .NET Framework 4.8
# ==============================================================================
Write-Log "--- Check: .NET Framework 4.8" -Level STEP
$dotnetKey  = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
$dotnetRelease = (Get-ItemProperty -Path $dotnetKey -ErrorAction SilentlyContinue).Release
# Release number 528040+ = .NET 4.8
$dotnetOk = $dotnetRelease -ge 528040
Add-Result ".NET Framework 4.8" $dotnetOk "Registry Release value: $dotnetRelease $(if($dotnetOk){'(>= 528040 = 4.8)'}else{'(< 528040 - install required)'})"

# ==============================================================================
# CHECK 8 — .NET Hosting 8.x
# ==============================================================================
Write-Log "--- Check: .NET Core Hosting" -Level STEP
$hostingPaths = @(
    "HKLM:\SOFTWARE\dotnet\Setup\InstalledVersions\x64\hostfxr",
    "HKLM:\SOFTWARE\Microsoft\ASP.NET Core\Shared Framework\v8.0"
)
$hostingFound = $false
foreach ($p in $hostingPaths) {
    if (Test-Path $p) { $hostingFound = $true; break }
}
# Also check via dotnet command
if (-not $hostingFound) {
    try {
        $dotnetOutput = & dotnet --list-runtimes 2>$null
        if ($dotnetOutput -match "Microsoft.AspNetCore.App 8\.") { $hostingFound = $true }
    } catch {}
}
Add-Result ".NET Core Hosting 8.x" $hostingFound $(if ($hostingFound) { "Found" } else { "Not found - will be installed by upgrade script" })

# ==============================================================================
# CHECK 9 — IIS Installed (Server machines only)
# ==============================================================================
Write-Log "--- Check: IIS Installation" -Level STEP
$iisFeature = Get-WindowsOptionalFeature -Online -FeatureName "IIS-WebServerRole" -ErrorAction SilentlyContinue
$iisService = Get-Service -Name "W3SVC" -ErrorAction SilentlyContinue
$iisOk = ($iisFeature -and $iisFeature.State -eq 'Enabled') -or ($null -ne $iisService)
Add-Result "IIS Installed" $iisOk $(if ($iisOk) { "IIS is installed (W3SVC service found)" } else { "IIS not found - required for LRE Server machines" })

# ==============================================================================
# CHECK 10 — LRE Services Status
# ==============================================================================
Write-Log "--- Check: LRE Services" -Level STEP
$serviceMap = @{
    "IIS (W3SVC)"              = $ServiceIIS
    "LRE Backend Service"      = (Resolve-ServiceName $ServiceBackend)
    "LRE Alerts Service"       = (Resolve-ServiceName $ServiceAlerts)
    "LoadRunner Agent Service" = (Resolve-ServiceName $ServiceAgent)
    "Remote Mgmt Agent"        = (Resolve-ServiceName $ServiceRemoteMgmt)
}
foreach ($label in $serviceMap.Keys) {
    $svcName = $serviceMap[$label]
    if ($svcName) {
        $status = Get-ServiceStatus -ServiceName $svcName
        # For pre-checks we just report status, not pass/fail
        Add-Result "Service: $label" $true "[$status] ($svcName)"
    } else {
        Add-Result "Service: $label" $true "[NotFound] - may not apply to this machine type"
    }
}

# ==============================================================================
# CHECK 11 — Installer Share Accessible
# ==============================================================================
Write-Log "--- Check: Installer Share" -Level STEP
$shareOk = Test-ShareAccessible -SharePath $InstallerShare
Add-Result "Installer Share Accessible" $shareOk "Share: $InstallerShare"
if ($shareOk) {
    $serverExeOk = Test-Path $SetupServerExe
    $hostExeOk   = Test-Path $SetupHostExe
    $oneLGExeOk  = Test-Path $SetupOneLGExe
    Add-Result "setup_server.exe found" $serverExeOk $SetupServerExe
    Add-Result "setup_host.exe found"   $hostExeOk   $SetupHostExe
    Add-Result "SetupOneLG.exe found"   $oneLGExeOk  $SetupOneLGExe
}

# ==============================================================================
# CHECK 12 — No Conflicting Processes (installer not already running)
# ==============================================================================
Write-Log "--- Check: No Installer Processes Running" -Level STEP
$installerProcs = Get-Process -Name "setup","setup_server","setup_host","msiexec" -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -ne "msiexec" -or $_.MainWindowTitle -ne "" }
$noConflict = ($null -eq $installerProcs -or $installerProcs.Count -eq 0)
Add-Result "No Installer Processes Running" $noConflict $(
    if ($noConflict) { "Clean" }
    else { "Running: $(($installerProcs | Select-Object -Expand Name) -join ', ')" }
)

# ==============================================================================
# SUMMARY REPORT
# ==============================================================================
Write-Banner "Pre-Upgrade Check Summary  |  $(hostname)"

$passCount = ($results | Where-Object { $_.Status -eq 'PASS' }).Count
$failCount = ($results | Where-Object { $_.Status -eq 'FAIL' }).Count

$results | Format-Table -AutoSize -Property Check, Status, Detail | Out-String | Write-Host

Write-Host ""
if ($failCount -eq 0) {
    Write-Host "  RESULT: ALL CHECKS PASSED ($passCount/$($results.Count)) - Machine is ready for upgrade." `
        -ForegroundColor Green
    Write-Log "RESULT: ALL CHECKS PASSED - Machine is ready for upgrade." -Level SUCCESS
} else {
    Write-Host "  RESULT: $failCount CHECK(S) FAILED - Resolve issues above before upgrading." `
        -ForegroundColor Red
    Write-Log "RESULT: $failCount FAILED - resolve before upgrading." -Level ERROR
}

Write-Host "`n  Full log: $Global:LogFile`n" -ForegroundColor Cyan
Stop-Transcript -ErrorAction SilentlyContinue
