# ==============================================================================
# Shared-HostUpgrade.ps1
# Shared upgrade logic for LRE Host machines.
# Used by: 02_Upgrade_Controller.ps1 and 03_Upgrade_DataProcessor.ps1
# DO NOT run this file directly.
# ==============================================================================

function Invoke-HostUpgrade {
    param(
        [ValidateSet("Controller","DataProcessor")]
        [string]$ComponentType
    )

    # ============================================================
    # PHASE 1 — PRE-UPGRADE CHECKS
    # ============================================================
    Write-Log "PHASE 1 — Pre-Upgrade Checks" -Level STEP

    # 1.1 — Installer share
    if (-not (Test-ShareAccessible -SharePath $InstallerShare)) {
        Write-Log "Installer share not accessible: $InstallerShare" -Level ERROR
        exit 1
    }
    if (-not (Test-Path $SetupHostExe)) {
        Write-Log "setup_host.exe not found at: $SetupHostExe" -Level ERROR
        exit 1
    }
    Write-Log "Installer media found." -Level SUCCESS

    # 1.2 — Installed LRE version
    Write-Log "Checking installed LRE version ..." -Level INFO
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
    Write-Log "Version check passed." -Level SUCCESS

    # 1.3 — Disk space
    if (-not (Test-DiskSpace -Path "C:\" -MinimumFreeGB 10)) {
        exit 1
    }

    # 1.4 — Pending reboot
    if (Test-PendingReboot) {
        Write-Log "Pending reboot detected. Reboot first, then re-run." -Level ERROR
        exit 1
    }

    # ============================================================
    # PHASE 2 — RESOLVE PUBLIC KEY
    # ============================================================
    Write-Log "PHASE 2 — Resolving Public Key" -Level STEP

    $publicKey = $PublicKeyValue  # from upgrade_config.ps1 (may be empty)

    if ([string]::IsNullOrWhiteSpace($publicKey)) {
        # Try reading from the network share (written by 01_Upgrade_LREServer.ps1)
        if (Test-Path $PublicKeySharePath) {
            $publicKey = (Get-Content $PublicKeySharePath -Raw).Trim()
            Write-Log "Public Key loaded from: $PublicKeySharePath" -Level SUCCESS
        }
    }

    if ([string]::IsNullOrWhiteSpace($publicKey)) {
        Write-Log "Public Key not found automatically." -Level WARN
        Write-Host "`n  The Public Key was not found at: $PublicKeySharePath" -ForegroundColor Yellow
        Write-Host "  You can find it in:" -ForegroundColor Yellow
        Write-Host "    1. The LRE Server config file: $ServerInstallDir\dat\pcs.config" -ForegroundColor Yellow
        Write-Host "    2. The Server upgrade log: $LogDir\PublicKey.txt  (on the server machine)" -ForegroundColor Yellow
        Write-Host "    3. REST API on server: GET http://<server>/Admin/rest/v1/configuration/getPublicKey" -ForegroundColor Yellow
        Write-Host "`n  Enter the Public Key (paste and press Enter): " -ForegroundColor Yellow -NoNewline
        $publicKey = Read-Host
        if ([string]::IsNullOrWhiteSpace($publicKey)) {
            Write-Log "Public Key is required but was not provided. Aborting." -Level ERROR
            exit 1
        }
    }
    Write-Log "Public Key is set." -Level INFO

    # ============================================================
    # PHASE 3 — GATHER CREDENTIALS
    # ============================================================
    Write-Log "PHASE 3 — Gathering Credentials" -Level STEP

    $script:SysUserPwd = Get-SecureValue "Enter System User password for '$SystemUserName'" $SystemUserPwd

    # ============================================================
    # PHASE 4 — STOP AGENT SERVICES
    # ============================================================
    Write-Log "PHASE 4 — Stopping LRE Agent Services" -Level STEP

    $agentSvc   = Resolve-ServiceName $ServiceAgent
    $remoteSvc  = Resolve-ServiceName $ServiceRemoteMgmt

    if ($agentSvc)  { Stop-ServiceSafely -ServiceName $agentSvc  }
    if ($remoteSvc) { Stop-ServiceSafely -ServiceName $remoteSvc }

    # ============================================================
    # PHASE 5 — INSTALL PREREQUISITES
    # ============================================================
    Write-Log "PHASE 5 — Installing Prerequisites" -Level STEP

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

    Install-Prereq ".NET Framework 4.8"    $DotNet48Exe    '/LCID /q /norestart /c:"install /q"'
    Install-Prereq "VC++ Redist x86"       $VCRedistX86    '/quiet /norestart'
    Install-Prereq "VC++ Redist x64"       $VCRedistX64    '/quiet /norestart'

    # ============================================================
    # PHASE 6 — PREPARE UserInput.xml
    # Strategy: copy vendor XML from share and merge our values;
    # fall back to local template if share is unreachable.
    # ============================================================
    Write-Log "PHASE 6 — Preparing UserInput.xml" -Level STEP

    if (-not (Test-Path $TempUserInputDir)) {
        New-Item -ItemType Directory -Path $TempUserInputDir -Force | Out-Null
    }
    $tempXml = "$TempUserInputDir\Host_${ComponentType}_UserInput_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"

    $propertyMap = Get-HostPropertyMap -PublicKey $publicKey

    New-UserInputXml `
        -VendorXmlPath        $HostUserInput `
        -FallbackTemplatePath "$ScriptRoot\..\config\LRE_Host_UserInput.xml" `
        -PropertyMap          $propertyMap `
        -OutputPath           $tempXml

    # Log sanitised summary
    Write-Log "UserInput.xml property summary:" -Level INFO
    $propertyMap.GetEnumerator() | Where-Object { $_.Key -notmatch "Pwd|Password|Key|Passphrase" } |
        Sort-Object Key | ForEach-Object { Write-Log "  $($_.Key) = $($_.Value)" -Level INFO }

    # ============================================================
    # PHASE 7 — RUN SILENT HOST UPGRADE
    # ============================================================
    Write-Log "PHASE 7 — Running LRE Host Upgrade" -Level STEP
    Confirm-Action "About to run LRE Host upgrade on $(hostname) as [$ComponentType]. Continue?"

    $installerLog  = "$LogDir\${ComponentType}_Installer_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $installerArgs = "/s LRASPCHOST=1 USER_CONFIG_FILE_PATH=`"$tempXml`" INSTALLDIR=`"$HostInstallDir`""

    Write-Log "Installer: $SetupHostExe" -Level INFO
    Write-Log "Arguments: $installerArgs" -Level INFO

    $exitCode = Invoke-Installer `
        -Executable $SetupHostExe `
        -Arguments  $installerArgs `
        -LogFile    $installerLog `
        -TimeoutMinutes 60

    switch ($exitCode) {
        0    { Write-Log "$ComponentType upgrade completed successfully." -Level SUCCESS }
        3010 { Write-Log "$ComponentType upgrade succeeded - REBOOT REQUIRED." -Level WARN }
        1602 { Write-Log "Upgrade cancelled by user." -Level WARN; exit 1602 }
        1603 { Write-Log "Fatal install error (1603). Check: $installerLog" -Level ERROR; exit 1603 }
        default {
            Write-Log "Unexpected exit code: $exitCode. Check: $installerLog" -Level ERROR
            exit $exitCode
        }
    }

    if ($exitCode -eq 3010) {
        Request-RebootIfNeeded -ExitCode $exitCode
    }

    # ============================================================
    # PHASE 8 — POST-UPGRADE: VERIFY CONFIG LOG
    # ============================================================
    Write-Log "PHASE 8 — Verifying Configuration" -Level STEP

    $configLog = "$HostInstallDir\orchidtmp\Configuration\configurationWizardLog_pcs.txt"
    if (Test-Path $configLog) {
        $logTail = Get-Content $configLog -Tail 20
        $logTail | ForEach-Object { Write-Log "  $_" -Level INFO }
    } else {
        Write-Log "Config wizard log not found - this may be normal for host upgrades." -Level INFO
    }

    # Check lts.config consistency
    $ltsConfig = "$HostInstallDir\dat\lts.config"
    Test-ConfigUserConsistency -ConfigFilePath $ltsConfig -ExpectedUser $SystemUserName | Out-Null

    # ============================================================
    # PHASE 9 — RESTART SERVICES
    # ============================================================
    Write-Log "PHASE 9 — Starting Services" -Level STEP

    if ($agentSvc)  { Start-ServiceSafely -ServiceName $agentSvc  }
    if ($remoteSvc) { Start-ServiceSafely -ServiceName $remoteSvc }

    # ============================================================
    # VERIFY VERSION
    # ============================================================
    $newInstalled = Get-LREInstalledVersion
    if ($newInstalled -and ($newInstalled.Version -like "*$TargetVersion*")) {
        Write-Log "Version confirmed: $($newInstalled.Version)" -Level SUCCESS
    } else {
        Write-Log "Post-upgrade version: $($newInstalled.Version)  (expected $TargetVersion)" -Level WARN
    }

    # ============================================================
    # SUMMARY
    # ============================================================
    Write-Banner "$ComponentType HOST UPGRADE COMPLETE  |  $(hostname)"
    Write-Log "Log file: $Global:LogFile" -Level INFO
    Write-Log "Installer log: $installerLog" -Level INFO
    Write-Log "" -Level INFO
    Write-Log "NEXT STEPS:" -Level INFO
    Write-Log "  - In LRE Administration > Hosts, verify this host shows version $TargetVersion." -Level INFO
    Write-Log "  - If the host shows 'Reconfigure needed', click Reconfigure Host in Administration." -Level INFO
    Write-Log "  - Run 05_PostUpgradeVerify.ps1 after all components are upgraded." -Level INFO

    # Secure cleanup
    if (Test-Path $tempXml) {
        Remove-Item $tempXml -Force
        Write-Log "Temp UserInput.xml removed." -Level INFO
    }

    Stop-Transcript -ErrorAction SilentlyContinue
}
