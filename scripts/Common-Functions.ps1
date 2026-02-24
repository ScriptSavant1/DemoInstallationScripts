# ==============================================================================
# Common-Functions.ps1
# Shared helper functions dot-sourced by every LRE upgrade script.
# DO NOT run this file directly.
# ==============================================================================

# ------------------------------------------------------------------------------
# LOGGING
# ------------------------------------------------------------------------------
function Initialize-Log {
    param([string]$LogFile)
    if (-not (Test-Path (Split-Path $LogFile -Parent))) {
        New-Item -ItemType Directory -Path (Split-Path $LogFile -Parent) -Force | Out-Null
    }
    # Start PowerShell transcript alongside the custom log
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

    # Write to console with color
    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'STEP'    { Write-Host "`n$line" -ForegroundColor White }
    }

    # Append to log file
    if ($Global:LogFile) {
        Add-Content -Path $Global:LogFile -Value $line
    }
}

function Write-Banner {
    param([string]$Title)
    $bar = "=" * 72
    Write-Host "`n$bar" -ForegroundColor Magenta
    Write-Host "  $Title" -ForegroundColor Magenta
    Write-Host "$bar`n" -ForegroundColor Magenta
    Write-Log "=== $Title ===" -Level INFO
}

# ------------------------------------------------------------------------------
# SERVICE HELPERS
# ------------------------------------------------------------------------------

# Resolve a service name from a list of candidates (handles vendor renames)
function Resolve-ServiceName {
    param([string[]]$Candidates)
    foreach ($name in $Candidates) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if ($svc) { return $svc.Name }
    }
    return $null
}

function Stop-ServiceSafely {
    param(
        [string]$ServiceName,
        [int]$TimeoutSeconds = 120
    )
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Log "Service '$ServiceName' not found - skipping stop." -Level WARN
        return
    }
    if ($svc.Status -eq 'Stopped') {
        Write-Log "Service '$ServiceName' is already stopped." -Level INFO
        return
    }
    Write-Log "Stopping service: $ServiceName ..." -Level INFO
    try {
        Stop-Service -Name $ServiceName -Force -ErrorAction Stop
        $svc.WaitForStatus('Stopped', (New-TimeSpan -Seconds $TimeoutSeconds))
        Write-Log "Service '$ServiceName' stopped successfully." -Level SUCCESS
    } catch {
        Write-Log "Failed to stop service '$ServiceName': $_" -Level ERROR
        throw
    }
}

function Start-ServiceSafely {
    param(
        [string]$ServiceName,
        [int]$TimeoutSeconds = 120
    )
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Log "Service '$ServiceName' not found - skipping start." -Level WARN
        return
    }
    if ($svc.Status -eq 'Running') {
        Write-Log "Service '$ServiceName' is already running." -Level INFO
        return
    }
    Write-Log "Starting service: $ServiceName ..." -Level INFO
    try {
        Start-Service -Name $ServiceName -ErrorAction Stop
        $svc.WaitForStatus('Running', (New-TimeSpan -Seconds $TimeoutSeconds))
        Write-Log "Service '$ServiceName' started successfully." -Level SUCCESS
    } catch {
        Write-Log "Failed to start service '$ServiceName': $_" -Level ERROR
        throw
    }
}

function Get-ServiceStatus {
    param([string]$ServiceName)
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) { return "NotFound" }
    return $svc.Status.ToString()
}

# ------------------------------------------------------------------------------
# INSTALLED VERSION CHECK
# ------------------------------------------------------------------------------
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
        if (Test-Path $regPath) {
            $keys = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
            foreach ($key in $keys) {
                $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                foreach ($term in $searchTerms) {
                    if ($props.DisplayName -like "*$term*") {
                        return [PSCustomObject]@{
                            DisplayName = $props.DisplayName
                            Version     = $props.DisplayVersion
                            InstallDir  = $props.InstallLocation
                            UninstallString = $props.UninstallString
                        }
                    }
                }
            }
        }
    }
    return $null
}

# ------------------------------------------------------------------------------
# DISK SPACE CHECK
# ------------------------------------------------------------------------------
function Test-DiskSpace {
    param(
        [string]$Path = "C:\",
        [long]$MinimumFreeGB = 20
    )
    $drive = Split-Path -Qualifier $Path
    $disk  = Get-PSDrive -Name ($drive.TrimEnd(':')) -ErrorAction SilentlyContinue
    if (-not $disk) {
        # Fallback via WMI
        $wmiDisk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$drive'" -ErrorAction SilentlyContinue
        if ($wmiDisk) {
            $freeGB = [math]::Round($wmiDisk.FreeSpace / 1GB, 2)
        } else {
            Write-Log "Cannot determine disk space for '$Path'." -Level WARN
            return $true  # Don't block if we can't check
        }
    } else {
        $freeGB = [math]::Round($disk.Free / 1GB, 2)
    }
    Write-Log "Disk free on '$drive': $freeGB GB  (minimum required: $MinimumFreeGB GB)" -Level INFO
    if ($freeGB -lt $MinimumFreeGB) {
        Write-Log "INSUFFICIENT DISK SPACE. Free: $freeGB GB, Required: $MinimumFreeGB GB" -Level ERROR
        return $false
    }
    return $true
}

# ------------------------------------------------------------------------------
# PENDING REBOOT CHECK
# ------------------------------------------------------------------------------
function Test-PendingReboot {
    $pendingKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
        "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
    )
    foreach ($key in $pendingKeys) {
        if (Test-Path $key) {
            return $true
        }
    }
    # Also check via WMI
    $ccm = Invoke-WmiMethod -Namespace "root\ccm\clientsdk" `
           -Class "CCM_ClientUtilities" -Name "DetermineIfRebootPending" -ErrorAction SilentlyContinue
    if ($ccm -and ($ccm.RebootPending -or $ccm.IsHardRebootPending)) {
        return $true
    }
    return $false
}

# ------------------------------------------------------------------------------
# NETWORK SHARE CONNECTIVITY
# ------------------------------------------------------------------------------
function Test-ShareAccessible {
    param([string]$SharePath)
    try {
        $result = Test-Path $SharePath -ErrorAction Stop
        return $result
    } catch {
        return $false
    }
}

# ------------------------------------------------------------------------------
# USERINPUT.XML — VENDOR MERGE + FALLBACK TEMPLATE
#
# Strategy (in order of preference):
#   1. Copy vendor's UserInput.xml from the installer share and update values
#      by matching <Property Name="KEY"> nodes — safest: preserves every
#      installer-specific property we don't know about.
#   2. If vendor XML is not accessible, fall back to our own template and do
#      simple %%PLACEHOLDER%% token substitution.
# ------------------------------------------------------------------------------
function New-UserInputXml {
    param(
        [string]$VendorXmlPath,           # e.g. \\share\Setup\Install\Server\UserInput.xml
        [string]$FallbackTemplatePath,    # e.g. ..\config\LRE_Server_UserInput.xml
        [hashtable]$PropertyMap,          # key = Property Name attribute, value = desired value
        [string]$OutputPath
    )

    # --- Ensure output directory exists ---
    $outDir = Split-Path $OutputPath -Parent
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    # ---------------------------------------------------------------
    # PATH 1: Vendor XML from installer share (preferred)
    # ---------------------------------------------------------------
    if (-not [string]::IsNullOrWhiteSpace($VendorXmlPath) -and (Test-Path $VendorXmlPath)) {
        Write-Log "Loading vendor UserInput.xml from installer share: $VendorXmlPath" -Level INFO

        try {
            [xml]$xml = Get-Content $VendorXmlPath -Raw -Encoding UTF8

            $updatedCount = 0
            $addedCount   = 0

            foreach ($propName in $PropertyMap.Keys) {
                $value = $PropertyMap[$propName]

                # Try <Property Name="propName"> structure (most common in LRE)
                $node = $xml.SelectSingleNode("//Property[@Name='$propName']")

                if ($null -ne $node) {
                    $node.InnerText = $value
                    $updatedCount++
                } else {
                    # Try <propName> element structure (some older versions)
                    $node = $xml.SelectSingleNode("//$propName")
                    if ($null -ne $node) {
                        $node.InnerText = $value
                        $updatedCount++
                    } else {
                        # Property not present in vendor XML — add it to the root element
                        $newProp = $xml.CreateElement("Property")
                        $newProp.SetAttribute("Name", $propName)
                        $newProp.InnerText = $value
                        $xml.DocumentElement.AppendChild($newProp) | Out-Null
                        $addedCount++
                        Write-Log "  Property '$propName' not in vendor XML — added." -Level WARN
                    }
                }
            }

            $xml.Save($OutputPath)
            Write-Log "UserInput.xml saved: $OutputPath  (updated: $updatedCount, added: $addedCount)" -Level SUCCESS
            return
        } catch {
            Write-Log "Failed to process vendor UserInput.xml: $_" -Level WARN
            Write-Log "Falling back to local template ..." -Level WARN
        }
    } else {
        Write-Log "Vendor UserInput.xml not found at: $VendorXmlPath" -Level WARN
        Write-Log "Falling back to local template ..." -Level WARN
    }

    # ---------------------------------------------------------------
    # PATH 2: Fallback — our own template with %%PLACEHOLDER%% tokens
    # ---------------------------------------------------------------
    if ([string]::IsNullOrWhiteSpace($FallbackTemplatePath) -or -not (Test-Path $FallbackTemplatePath)) {
        throw "Neither vendor UserInput.xml nor fallback template is accessible.`n  Vendor path : $VendorXmlPath`n  Fallback    : $FallbackTemplatePath"
    }

    Write-Log "Using fallback template: $FallbackTemplatePath" -Level WARN

    # Build a reverse-lookup: %%KEY%% -> value  (keys are the Property Name values)
    # The fallback template uses %%UPPER_SNAKE_CASE%% tokens mapped from the same hashtable
    $content = Get-Content $FallbackTemplatePath -Raw -Encoding UTF8
    foreach ($propName in $PropertyMap.Keys) {
        # Convert Property Name -> template token  (e.g. "DbServerHost" -> "%%DB_SERVER_HOST%%")
        # The mapping is embedded in the template itself so just do a direct named-token pass.
        $content = $content -replace [regex]::Escape("%%$propName%%"), $PropertyMap[$propName]
    }

    Set-Content -Path $OutputPath -Value $content -Encoding UTF8
    Write-Log "UserInput.xml written from fallback template: $OutputPath" -Level SUCCESS
}

# Helper: build the Property Name -> value map using the token names that match
# both the vendor XML Property Name attributes AND our %%PLACEHOLDER%% tokens.
function Get-ServerPropertyMap {
    return @{
        UseDefaultUserSetting = $UseDefaultUser
        DomainName            = $DomainName
        SystemUserName        = $SystemUserName
        SystemUserPwd         = $script:SysUserPwd
        LW_CRYPTO_INIT_STRING = $CryptoKey
        FileSystemRoot        = $FileSystemRoot
        DbType                = $DbType
        DbServerHost          = $DbServerHost
        DbServerPort          = $DbServerPort
        DbAdminUser           = $DbAdminUser
        DbAdminPassword       = $script:DbAdminPwd
        DbUsername            = $DbUsername
        DbPassword            = $script:DbPwd
        LabDbName             = $LabDbName
        AdminDbName           = $AdminDbName
        SiteDbName            = $SiteDbName
        IIS_WEB_SITE_NAME     = $IisWebSiteName
        IisSecureConfiguration= $IisSecureConfig
        IisSecureHostName     = $IisSecureHostName
        IisSecurePort         = $IisSecurePort
        ImportCertificate     = $ImportCertificate
        CertificateStore      = $CertificateStore
        CertificateName       = $CertificateName
        CertificateFilePath   = $CertificateFilePath
        CertificatePassword   = $CertificatePassword
    }
}

function Get-HostPropertyMap {
    param([string]$PublicKey)
    return @{
        UseDefaultUserSetting = $UseDefaultUser
        DomainName            = $DomainName
        SystemUserName        = $SystemUserName
        SystemUserPwd         = $script:SysUserPwd
        LW_CRYPTO_INIT_STRING = $CryptoKey
        PublicKey             = $PublicKey
    }
}

# ------------------------------------------------------------------------------
# INSTALLER EXECUTION
# ------------------------------------------------------------------------------
function Invoke-Installer {
    param(
        [string]$Executable,
        [string]$Arguments,
        [string]$LogFile,
        [int]$TimeoutMinutes = 60
    )
    Write-Log "Launching: $Executable $Arguments" -Level STEP
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName  = $Executable
    $startInfo.Arguments = $Arguments
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError  = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo

    # Capture output asynchronously
    $stdout = [System.Text.StringBuilder]::new()
    $stderr = [System.Text.StringBuilder]::new()
    $process.add_OutputDataReceived({ param($s,$e)
        if ($e.Data) {
            [void]$stdout.AppendLine($e.Data)
            Write-Host "  | $($e.Data)" -ForegroundColor DarkGray
        }
    })
    $process.add_ErrorDataReceived({ param($s,$e)
        if ($e.Data) {
            [void]$stderr.AppendLine($e.Data)
            Write-Host "  ! $($e.Data)" -ForegroundColor Yellow
        }
    })

    $process.Start() | Out-Null
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()

    $timeout = [TimeSpan]::FromMinutes($TimeoutMinutes)
    $finished = $process.WaitForExit([int]$timeout.TotalMilliseconds)

    if (-not $finished) {
        $process.Kill()
        throw "Installer timed out after $TimeoutMinutes minutes."
    }

    $exitCode = $process.ExitCode
    Write-Log "Installer exit code: $exitCode" -Level INFO

    # Log output to file
    if ($LogFile) {
        Add-Content -Path $LogFile -Value "--- INSTALLER STDOUT ---"
        Add-Content -Path $LogFile -Value $stdout.ToString()
        Add-Content -Path $LogFile -Value "--- INSTALLER STDERR ---"
        Add-Content -Path $LogFile -Value $stderr.ToString()
    }

    return $exitCode
}

# ------------------------------------------------------------------------------
# SECURE PASSWORD PROMPT
# ------------------------------------------------------------------------------
function Get-SecureValue {
    param(
        [string]$PromptMessage,
        [string]$ExistingValue
    )
    if (-not [string]::IsNullOrWhiteSpace($ExistingValue)) {
        return $ExistingValue
    }
    $secure = Read-Host -Prompt $PromptMessage -AsSecureString
    $plain  = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                  [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
    return $plain
}

# ------------------------------------------------------------------------------
# CONFIG FILE CONSISTENCY CHECK  (pcs.config / lts.config)
# ------------------------------------------------------------------------------
function Test-ConfigUserConsistency {
    param(
        [string]$ConfigFilePath,
        [string]$ExpectedUser
    )
    if (-not (Test-Path $ConfigFilePath)) {
        Write-Log "Config file not found: $ConfigFilePath" -Level WARN
        return $false
    }
    $content = Get-Content $ConfigFilePath -Raw
    if ($content -match $ExpectedUser) {
        Write-Log "Config consistency OK - user '$ExpectedUser' found in: $ConfigFilePath" -Level SUCCESS
        return $true
    } else {
        Write-Log "WARNING: user '$ExpectedUser' NOT found in: $ConfigFilePath" -Level WARN
        Write-Log "Check Upgrade Tips in the guide - service user must be case-sensitively consistent." -Level WARN
        return $false
    }
}

# ------------------------------------------------------------------------------
# CONFIRM PROMPT
# ------------------------------------------------------------------------------
function Confirm-Action {
    param([string]$Message)
    Write-Host "`n$Message" -ForegroundColor Yellow
    Write-Host "Press [Y] to continue, any other key to abort: " -ForegroundColor Yellow -NoNewline
    $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Write-Host $key.Character
    if ($key.Character -ne 'Y' -and $key.Character -ne 'y') {
        Write-Log "Upgrade aborted by user." -Level WARN
        Stop-Transcript -ErrorAction SilentlyContinue
        exit 1
    }
}

# ------------------------------------------------------------------------------
# REBOOT HELPER
# ------------------------------------------------------------------------------
function Request-RebootIfNeeded {
    param([int]$ExitCode)
    # Exit code 3010 = success but reboot required
    if ($ExitCode -eq 3010) {
        Write-Log "Installer requires a reboot (exit code 3010)." -Level WARN
        Write-Host "`nThe machine needs to restart to complete the installation." -ForegroundColor Yellow
        Write-Host "After reboot, re-run this script if any post-upgrade steps remain." -ForegroundColor Yellow
        Confirm-Action "Reboot now?"
        Restart-Computer -Force
    }
}
