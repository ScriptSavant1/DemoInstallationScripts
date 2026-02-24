# ==============================================================================
# LRE Upgrade Configuration  |  25.1 -> 26.1
# ==============================================================================
# INSTRUCTIONS:
#   1. Fill in every <-- UPDATE THIS section below.
#   2. This file is dot-sourced by every upgrade script.
#      Do NOT run it directly.
#   3. Passwords left blank will be securely prompted at runtime.
# ==============================================================================

# ------------------------------------------------------------------------------
# INSTALLER MEDIA
# ------------------------------------------------------------------------------
# UNC path to the root of the 26.1 installer share.
# Must be accessible from every machine that runs an upgrade script.
$Global:InstallerShare = "\\fileserver\LRE_26.1"          # <-- UPDATE THIS

# Relative paths within the share (do not change unless vendor changes layout)
$Global:SetupServerExe  = "$InstallerShare\Setup\En\setup_server.exe"
$Global:SetupHostExe    = "$InstallerShare\Setup\En\setup_host.exe"
$Global:SetupOneLGExe   = "$InstallerShare\Standalone Applications\SetupOneLG.exe"
$Global:ServerMsi       = "$InstallerShare\Setup\Install\Server\LRE_Server.msi"
$Global:HostMsi         = "$InstallerShare\Setup\Install\Host\LoadRunner_x64.msi"
$Global:OneLGMsi        = "$InstallerShare\Standalone Applications\OneLG_x64.msi"
$Global:ServerUserInput = "$InstallerShare\Setup\Install\Server\UserInput.xml"
$Global:HostUserInput   = "$InstallerShare\Setup\Install\Host\UserInput.xml"

# Prerequisite installers (relative paths within the share)
$Global:DotNet48Exe     = "$InstallerShare\Setup\Common\dotnet48\ndp48-x86-x64-allos-enu.exe"
$Global:DotNetHostingExe= "$InstallerShare\Setup\Common\dotnet_hosting\dotnet-hosting-8.0.17-win.exe"
$Global:VCRedistX86     = "$InstallerShare\Setup\Common\vc2022_redist_x86\vc_redist.x86.exe"
$Global:VCRedistX64     = "$InstallerShare\Setup\Common\vc2022_redist_x64\vc_redist.x64.exe"

# ------------------------------------------------------------------------------
# LOCAL PATHS (on each target machine)
# ------------------------------------------------------------------------------
$Global:LogDir              = "C:\LRE_UpgradeLogs"
$Global:TempUserInputDir    = "C:\LRE_UpgradeLogs\UserInput"

$Global:ServerInstallDir    = "C:\Program Files\OpenText\LRE"           # <-- UPDATE IF DIFFERENT
$Global:HostInstallDir      = "C:\Program Files\OpenText\Performance Center Host" # <-- UPDATE IF DIFFERENT
$Global:OneLGInstallDir     = "C:\Program Files\OpenText\OneLG"         # <-- UPDATE IF DIFFERENT

# ------------------------------------------------------------------------------
# SYSTEM USER
# Use IUSR_METRO defaults unless a custom domain user was configured.
# IMPORTANT: must be identical (case-sensitive) on ALL nodes.
# ------------------------------------------------------------------------------
$Global:UseDefaultUser      = "true"       # "true" = use IUSR_METRO; "false" = use custom
$Global:DomainName          = "."          # "." for local account; "MYDOMAIN" for domain
$Global:SystemUserName      = "IUSR_METRO" # <-- UPDATE IF USING CUSTOM USER
$Global:SystemUserPwd       = ""           # Leave blank to be prompted; or set value here

# ------------------------------------------------------------------------------
# ENCRYPTION KEY
# Must be IDENTICAL across every LRE Server, Controller, Data Processor,
# and Load Generator node. If blank, the existing value is preserved.
# ------------------------------------------------------------------------------
$Global:CryptoKey           = ""           # LW_CRYPTO_INIT_STRING  <-- UPDATE THIS

# ------------------------------------------------------------------------------
# DATABASE  (Microsoft SQL Server)
# ------------------------------------------------------------------------------
$Global:DbType              = "MS-SQL"
$Global:DbServerHost        = "sql-server-01"   # <-- UPDATE THIS
$Global:DbServerPort        = "1433"             # <-- UPDATE IF DIFFERENT
$Global:DbAdminUser         = "sa"               # <-- UPDATE THIS
$Global:DbAdminPassword     = ""                 # Leave blank to be prompted
$Global:DbUsername          = "lre_user"         # <-- UPDATE THIS (app DB user)
$Global:DbPassword          = ""                 # Leave blank to be prompted
$Global:LabDbName           = "LRE_LAB"          # <-- UPDATE THIS
$Global:AdminDbName         = "LRE_ADMIN"        # <-- UPDATE THIS
$Global:SiteDbName          = "LRE_SITE"         # <-- UPDATE THIS

# ------------------------------------------------------------------------------
# IIS / TLS CERTIFICATE
# ------------------------------------------------------------------------------
$Global:IisWebSiteName      = "Default Web Site"
$Global:IisSecureConfig     = "true"
$Global:IisSecureHostName   = "lre.yourdomain.com"  # <-- UPDATE THIS (LRE server FQDN)
$Global:IisSecurePort       = "443"
$Global:ImportCertificate   = "false"               # "true" = import from file; "false" = use existing store cert
$Global:CertificateStore    = "My"
$Global:CertificateName     = ""                    # <-- UPDATE THIS (cert CN or thumbprint)
$Global:CertificateFilePath = ""                    # Only needed if ImportCertificate = "true"
$Global:CertificatePassword = ""                    # Only needed if ImportCertificate = "true"

# ------------------------------------------------------------------------------
# LRE REPOSITORY
# Fully qualified path. For clustered environments use a UNC path.
# ------------------------------------------------------------------------------
$Global:FileSystemRoot      = "C:\LRE_Repository"  # <-- UPDATE THIS

# ------------------------------------------------------------------------------
# SECURITY PASSPHRASE
# Minimum 12 alphanumeric characters. Must be IDENTICAL across ALL nodes.
# ------------------------------------------------------------------------------
$Global:SecurePassphrase    = ""           # <-- UPDATE THIS

# ------------------------------------------------------------------------------
# PUBLIC KEY
# Populated automatically by 01_Upgrade_LREServer.ps1.
# Can also be set manually if running host scripts independently.
# Path on the network share where the server script writes the public key.
# ------------------------------------------------------------------------------
$Global:PublicKeySharePath  = "$InstallerShare\PublicKey.txt"
$Global:PublicKeyValue      = ""           # Set manually if needed

# ------------------------------------------------------------------------------
# VERSION CONSTANTS (do not change)
# ------------------------------------------------------------------------------
$Global:ExpectedCurrentVersion  = "25.1"
$Global:TargetVersion           = "26.1"

# ------------------------------------------------------------------------------
# SERVICE NAMES
# These are the Windows service names (not display names).
# Adjust if your environment uses non-default service names.
# ------------------------------------------------------------------------------
$Global:ServiceIIS          = "W3SVC"
$Global:ServiceBackend      = @("LREBackend","OpenTextLREBackend","LRE_Backend")   # checked in order
$Global:ServiceAlerts       = @("LREAlerts","OpenTextLREAlerts","LRE_Alerts")
$Global:ServiceAgent        = @("magentservice","LoadRunnerAgentService","LRAgentService")
$Global:ServiceRemoteMgmt   = @("AlAgent","RemoteManagementAgent","al_agent")
