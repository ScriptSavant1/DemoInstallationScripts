# LRE Upgrade User Guide
## OpenText Enterprise Performance Engineering — 25.1 → 26.1
### Step-by-Step Instructions for Each Component

---

## Table of Contents

1. [Before You Begin](#1-before-you-begin)
2. [Understanding the Upgrade Order](#2-understanding-the-upgrade-order)
3. [Step 1 — Configure upgrade_config.ps1](#3-step-1--configure-upgrade_configps1)
4. [Step 2 — Run Pre-Upgrade Checks (All Machines)](#4-step-2--run-pre-upgrade-checks-all-machines)
5. [Step 3 — Database Backup](#5-step-3--database-backup)
6. [Step 4 — Upgrade the LRE Server](#6-step-4--upgrade-the-lre-server)
7. [Step 5 — Upgrade Controller Host(s)](#7-step-5--upgrade-controller-hosts)
8. [Step 6 — Upgrade Data Processor Host(s)](#8-step-6--upgrade-data-processor-hosts)
9. [Step 7 — Upgrade Load Generator(s)](#9-step-7--upgrade-load-generators)
10. [Step 8 — Post-Upgrade Verification](#10-step-8--post-upgrade-verification)
11. [Troubleshooting](#11-troubleshooting)
12. [Rollback Procedure](#12-rollback-procedure)
13. [Configuration Reference](#13-configuration-reference)

---

## 1. Before You Begin

### What You Need

| Requirement | Detail |
|------------|--------|
| LRE installer media | 26.1 installer package placed on a network share |
| Share accessible from all machines | All target machines can reach the UNC path |
| Local Administrator access | RDP into each machine as a local admin |
| SQL Server DB backups | DBA must back up all three LRE databases first |
| Maintenance window | Plan for downtime — all LRE components will be offline |
| PowerShell 5.1+ | Pre-installed on Windows Server 2016 and later |

### Network Share Layout Expected

The scripts expect the 26.1 installer to be laid out exactly as it comes from
the vendor. After extracting the installer package to your share, the structure
should look like this:

```
\\fileserver\LRE_26.1\
├── Setup\
│   ├── En\
│   │   ├── setup_server.exe          ← LRE Server silent installer
│   │   └── setup_host.exe            ← Host silent installer
│   ├── Install\
│   │   ├── Server\
│   │   │   ├── LRE_Server.msi
│   │   │   └── UserInput.xml         ← Vendor template (scripts read this)
│   │   └── Host\
│   │       ├── LoadRunner_x64.msi
│   │       └── UserInput.xml         ← Vendor template (scripts read this)
│   └── Common\
│       ├── dotnet48\
│       ├── dotnet_hosting\
│       ├── vc2022_redist_x86\
│       └── vc2022_redist_x64\
└── Standalone Applications\
    └── SetupOneLG.exe                ← OneLG Load Generator installer
```

> **Important:** If your share layout differs from the above, update the path
> variables in `config/upgrade_config.ps1` (the `$Global:Setup*` variables).

### How Execution Works

Each script is run **locally on the target machine** via RDP:

```
Your machine (RDP) ──► LRE Server machine    → run 01_Upgrade_LREServer.ps1
Your machine (RDP) ──► Controller machine    → run 02_Upgrade_Controller.ps1
Your machine (RDP) ──► Data Processor machine→ run 03_Upgrade_DataProcessor.ps1
Your machine (RDP) ──► Load Generator machine→ run 04_Upgrade_LoadGenerator.ps1
```

The scripts themselves pull the installer from the network share. You do not
need to pre-copy any files to the target machines.

---

## 2. Understanding the Upgrade Order

**This order is mandatory.** Do not upgrade hosts before the server.

```
┌─────────────────────────────────────────────────────┐
│  STEP 0  Run 00_PreUpgradeChecks.ps1 on ALL machines │
│          (no changes made — diagnostic only)         │
└─────────────────────────────────┬───────────────────┘
                                  │
┌─────────────────────────────────▼───────────────────┐
│  STEP 1  DBA takes DB backups                        │
└─────────────────────────────────┬───────────────────┘
                                  │
┌─────────────────────────────────▼───────────────────┐
│  STEP 2  Upgrade LRE Server                          │
│          01_Upgrade_LREServer.ps1                    │
│          → Captures Public Key → writes to share     │
└─────────────────────────────────┬───────────────────┘
                                  │
              ┌───────────────────┼───────────────────┐
              │                   │                   │
┌─────────────▼──────┐ ┌─────────▼──────┐ ┌─────────▼──────┐
│ STEP 3             │ │ STEP 4         │ │ STEP 5         │
│ Upgrade Controller │ │ Upgrade DP     │ │ Upgrade LG(s)  │
│ 02_Upgrade_...ps1  │ │ 03_Upgrade_..  │ │ 04_Upgrade_..  │
└─────────────┬──────┘ └─────────┬──────┘ └─────────┬──────┘
              │                   │                   │
              └───────────────────┼───────────────────┘
                                  │
┌─────────────────────────────────▼───────────────────┐
│  STEP 6  Run 05_PostUpgradeVerify.ps1 on ALL machines│
└─────────────────────────────────────────────────────┘
```

> Steps 3, 4, and 5 (Controller, DP, Load Generators) can be done in any order
> or in parallel across machines — but ALL must come after Step 2 (Server).

---

## 3. Step 1 — Configure upgrade_config.ps1

**Do this once before running anything else.**

### 3.1 Open the file

On any machine (or your local workstation), open:

```
config\upgrade_config.ps1
```

### 3.2 Set the installer share path

```powershell
$Global:InstallerShare = "\\fileserver\LRE_26.1"
```

Replace `\\fileserver\LRE_26.1` with the actual UNC path where you placed
the 26.1 installer. Every target machine must be able to reach this path.

### 3.3 Set the database details

```powershell
$Global:DbServerHost  = "sql-server-01"     # hostname or IP of your SQL Server
$Global:DbServerPort  = "1433"              # default SQL port
$Global:DbAdminUser   = "sa"                # SQL admin account (needs dbcreator role)
$Global:DbAdminPassword = ""               # leave blank → will prompt at runtime
$Global:DbUsername    = "lre_user"          # LRE application DB account
$Global:DbPassword    = ""                 # leave blank → will prompt at runtime
$Global:LabDbName     = "LRE_LAB"           # your existing Lab DB schema name
$Global:AdminDbName   = "LRE_ADMIN"         # your existing Admin DB schema name
$Global:SiteDbName    = "LRE_SITE"          # your existing Site Management schema name
```

> **Tip:** You can find the exact DB schema names by checking the existing
> 25.1 LRE Administration > DB configuration page, or by querying SQL Server.

### 3.4 Set the system user

If you kept the LRE default system user (most installations):

```powershell
$Global:UseDefaultUser  = "true"
$Global:SystemUserName  = "IUSR_METRO"
$Global:SystemUserPwd   = ""              # leave blank → prompted at runtime
```

If you configured a custom domain or local user:

```powershell
$Global:UseDefaultUser  = "false"
$Global:DomainName      = "MYDOMAIN"      # or "." for local account
$Global:SystemUserName  = "lre_service"
$Global:SystemUserPwd   = ""
```

> **Critical:** The system user name is **case-sensitive** and must be
> identical across the LRE Server, all Hosts, and the repository config.

### 3.5 Set the security passphrase

```powershell
$Global:SecurePassphrase = "MyPassphrase2024!"
```

- Minimum 12 alphanumeric characters
- Must be **identical on every node** (Server, Controller, DP, LG)
- This is the existing passphrase from your 25.1 environment

> **Where to find it:** On the existing 25.1 server, check
> `<installdir>\dat\pcs.config` — look for `CommunicationSecurity`.

### 3.6 Set the encryption key

```powershell
$Global:CryptoKey = "your_existing_encryption_key"
```

Find the existing value in your 25.1 installation:
- Server: `<installdir>\dat\pcs.config` → `LW_CRYPTO_INIT_STRING`
- Leave blank only if your environment uses the auto-generated default.

### 3.7 Set the IIS/SSL settings (Server only)

```powershell
$Global:IisSecureHostName   = "lre.yourdomain.com"   # FQDN of LRE server
$Global:IisSecurePort       = "443"
$Global:CertificateName     = "lre.yourdomain.com"   # cert CN or thumbprint
$Global:CertificateStore    = "My"                    # Windows cert store name
```

If you need to import a certificate from a file (`.pfx`):

```powershell
$Global:ImportCertificate   = "true"
$Global:CertificateFilePath = "C:\certs\lre.pfx"
$Global:CertificatePassword = ""    # leave blank → prompted at runtime
```

### 3.8 Set the repository path

```powershell
$Global:FileSystemRoot = "C:\LRE_Repository"
```

This must match the **existing** repository path from your 25.1 installation.
For clustered environments use a UNC path (e.g. `\\nas\LRE_Repository`).

### 3.9 Copy config to all machines

Copy the entire `LRE-Automates-Scripts` folder to each target machine — or
keep it on the network share and access it via UNC path from each machine.

---

## 4. Step 2 — Run Pre-Upgrade Checks (All Machines)

Run this on **every machine** before making any changes. It makes no
modifications — it is a read-only diagnostic.

### How to run

1. RDP into the machine as local Administrator
2. Open PowerShell as Administrator (right-click → Run as Administrator)
3. Navigate to the scripts folder:

```powershell
cd "C:\LRE-Automates-Scripts\scripts"
# Allow script execution (one-time per machine):
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
.\00_PreUpgradeChecks.ps1
```

### What it checks

| Check | What it looks for |
|-------|------------------|
| Administrator privileges | Script must run elevated |
| Operating system | Windows confirmed |
| System locale | English locale (non-English can cause install issues) |
| Pending reboot | Must be clear before upgrading |
| Installed LRE version | Must be 25.1 |
| Disk space | Minimum 20 GB free on C:\ |
| .NET Framework 4.8 | Required by LRE |
| .NET Core Hosting 8.x | Required by LRE Backend service |
| IIS installed | Required on Server machines |
| LRE services status | Reports current state of all LRE services |
| Installer share accessible | Network share reachable from this machine |
| Installer executables found | setup_server.exe, setup_host.exe, SetupOneLG.exe |
| No installer processes running | No previous install in progress |

### Expected result

```
  RESULT: ALL CHECKS PASSED (12/12) - Machine is ready for upgrade.
  Full log: C:\LRE_UpgradeLogs\PreUpgradeChecks_<hostname>_<date>.log
```

### If checks fail

| Failure | Action |
|---------|--------|
| `Pending Reboot` | Reboot the machine and re-run checks |
| `Disk Space` | Free at least 20 GB on C:\ |
| `Installer Share` | Verify the UNC path is correct and accessible |
| `LRE version not 25.1` | Confirm you are on the right machine |
| `IIS not installed` | Install IIS before upgrading the Server machine |
| `.NET 4.8` | Will be installed automatically by the upgrade script |

---

## 5. Step 3 — Database Backup

**This is a manual DBA step. Do not skip it.**

Provide your DBA with the following information:

```
SQL Server Host : <your DbServerHost value>
Databases to back up:
  1. <LabDbName>    (Lab database)
  2. <AdminDbName>  (Admin database)
  3. <SiteDbName>   (Site Management database)

Backup type: Full backup
Storage: Ensure backups are stored off the target SQL server
```

Wait for DBA confirmation that all three backups are complete before
proceeding to Step 4.

> The Server upgrade script will display a confirmation prompt reminding you
> to confirm backups before it proceeds.

---

## 6. Step 4 — Upgrade the LRE Server

Run this **first**, before any host upgrades.

### 6.1 RDP into the LRE Server machine as Administrator

### 6.2 Copy scripts to the machine (if not already there)

```
\\fileserver\LRE-Automates-Scripts\  →  C:\LRE-Automates-Scripts\
```

Or access the scripts directly from the network share.

### 6.3 Open elevated PowerShell and run

```powershell
cd "C:\LRE-Automates-Scripts\scripts"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
.\01_Upgrade_LREServer.ps1
```

### 6.4 What the script does — phase by phase

| Phase | What happens |
|-------|-------------|
| Pre-checks | Verifies version is 25.1, disk space, share access |
| Credentials | Prompts for DB admin password, DB user password, system user password, passphrase |
| DB backup confirm | Shows DB names, waits for your [Y] confirmation |
| Stop services | Stops IIS, LRE Backend Service, LRE Alerts Service |
| Install prerequisites | .NET 4.8, .NET Core Hosting 8.x, VC++ Redistributable x86/x64 |
| Prepare UserInput.xml | Copies vendor XML from share, merges config values, saves to temp |
| Run upgrade | Executes `setup_server.exe /s USER_CONFIG_FILE_PATH="<temp.xml>"` |
| Restart services | Starts IIS, Backend, Alerts services |
| Capture Public Key | Reads Public Key from REST API → saves to share and local log |
| Consistency check | Verifies service user is consistent in pcs.config and appsettings.json |
| Cleanup | Deletes temp UserInput.xml (contains credentials) |

### 6.5 Prompts you will see

```
Enter DB Admin password for 'sa': ****
Enter DB User password for 'lre_user': ****
Enter System User password for 'IUSR_METRO': ****
Enter Secure Communication Passphrase (min 12 chars): ****

IMPORTANT — databases must be backed up...
Press [Y] to continue, any other key to abort: Y

About to run LRE Server upgrade from \\server\LRE_26.1\Setup\En\setup_server.exe
Press [Y] to continue, any other key to abort: Y
```

### 6.6 Verify the Public Key was captured

After the script finishes, confirm:

```
C:\LRE_UpgradeLogs\PublicKey.txt        ← local copy
\\fileserver\LRE_26.1\PublicKey.txt     ← share copy (read by host scripts)
```

If the Public Key was not captured automatically, see
[Troubleshooting — Public Key not captured](#public-key-not-captured).

### 6.7 Verify the LRE Server is accessible

Open a browser on the server machine and navigate to:

```
https://lre.yourdomain.com/LRE/
```

You should see the LRE login page. Log in with your administrator credentials
to confirm the server is running correctly before proceeding to hosts.

### 6.8 Clustered LRE Server environment

If you have **two or more LRE Server nodes**:

1. Before running the script on **any** node, manually stop services on **all** nodes:
   ```powershell
   iisreset /stop
   Stop-Service -Name "<BackendServiceName>" -Force
   Stop-Service -Name "<AlertsServiceName>"  -Force
   ```
2. Run `01_Upgrade_LREServer.ps1` on **node 1** first.
3. When node 1 upgrade is complete, run the script on **node 2**.

---

## 7. Step 5 — Upgrade Controller Host(s)

Run this on each machine assigned as a **Controller** in LRE Administration.

### 7.1 Prerequisite

Confirm that Step 4 (LRE Server upgrade) is complete and the Public Key file
exists at `\\fileserver\LRE_26.1\PublicKey.txt`.

### 7.2 RDP into the Controller machine as Administrator

### 7.3 Run the script

```powershell
cd "C:\LRE-Automates-Scripts\scripts"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
.\02_Upgrade_Controller.ps1
```

### 7.4 What the script does — phase by phase

| Phase | What happens |
|-------|-------------|
| Pre-checks | Version is 25.1, disk space, share access |
| Resolve Public Key | Reads from `\\share\PublicKey.txt` automatically |
| Credentials | Prompts for system user password |
| Stop services | Stops LoadRunner Agent Service, Remote Management Agent |
| Install prerequisites | .NET 4.8, VC++ Redistributable |
| Prepare UserInput.xml | Merges Public Key + credentials into vendor XML |
| Run upgrade | Executes `setup_host.exe /s LRASPCHOST=1 USER_CONFIG_FILE_PATH="..."` |
| Restart services | Starts Agent + Remote Mgmt Agent |
| Consistency check | Verifies service user in lts.config |
| Cleanup | Deletes temp UserInput.xml |

### 7.5 After upgrade

In **LRE Administration**:
1. Go to **Maintenance > Hosts**
2. Find this Controller host
3. If it shows **"Reconfigure needed"**, select it and click **Reconfigure Host**
4. Verify the displayed version shows **26.1**

> The Controller **role/purpose** assignment is preserved — you do not need
> to reassign it after the upgrade.

---

## 8. Step 6 — Upgrade Data Processor Host(s)

Identical process to the Controller upgrade, using a different script.

### 8.1 Prerequisite

Step 4 (LRE Server upgrade) must be complete.

### 8.2 RDP into the Data Processor machine as Administrator

### 8.3 Run the script

```powershell
cd "C:\LRE-Automates-Scripts\scripts"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
.\03_Upgrade_DataProcessor.ps1
```

### 8.4 What the script does

Same phases as the Controller upgrade (the underlying installer is identical).
The component label in all logs reads "DataProcessor" for clarity.

### 8.5 After upgrade

Same post-upgrade steps as Controller:
- Check **LRE Administration > Maintenance > Hosts**
- Reconfigure if needed
- Verify version shows 26.1

---

## 9. Step 7 — Upgrade Load Generator(s)

Load Generators are upgraded using the **OneLG standalone installer**, which is
lighter than the full host installer.

### 9.1 Prerequisite

Step 4 (LRE Server upgrade) must be complete. The Public Key is **not**
required for OneLG machines — they do not need it in their installer.

### 9.2 RDP into each Load Generator machine as Administrator

### 9.3 Run the script

```powershell
cd "C:\LRE-Automates-Scripts\scripts"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
.\04_Upgrade_LoadGenerator.ps1
```

### 9.4 What the script does — phase by phase

| Phase | What happens |
|-------|-------------|
| Pre-checks | Detects install type (OneLG vs full Host), version check, disk space |
| Stop services | Stops LoadRunner Agent Service, Remote Management Agent |
| Install prerequisites | .NET 4.8, VC++ Redistributable |
| Run upgrade | Executes `SetupOneLG.exe -s -sp"/s" IS_RUNAS_SERVICE=1 START_LGA=1` |
| Restart services | Starts Agent + Remote Mgmt Agent |
| Verify agent mode | Confirms agent is running as a Windows service |

### 9.5 Important note about IUSR_METRO

> **Do NOT delete the `IUSR_METRO` local account** on Load Generator machines
> unless you configured a custom system user. The OneLG agent service runs
> under this account by default.

### 9.6 After upgrade

In **LRE Administration**:
1. Go to **Maintenance > Hosts**
2. Find each Load Generator
3. Reconfigure if needed
4. Verify version shows 26.1

### 9.7 Multiple Load Generators

The script runs on one machine at a time. For multiple LGs, RDP into each
and repeat Step 9.2–9.6. Since LGs are independent, you can do them
concurrently across separate RDP sessions.

---

## 10. Step 8 — Post-Upgrade Verification

Run this on **every machine** after all upgrade scripts have completed.

### 10.1 Run the script

```powershell
cd "C:\LRE-Automates-Scripts\scripts"
.\05_PostUpgradeVerify.ps1
```

The script auto-detects the component type. You can also specify it:

```powershell
.\05_PostUpgradeVerify.ps1 -ComponentType Server
.\05_PostUpgradeVerify.ps1 -ComponentType Controller
.\05_PostUpgradeVerify.ps1 -ComponentType DataProcessor
.\05_PostUpgradeVerify.ps1 -ComponentType LoadGenerator
```

### 10.2 What it verifies

| Check | Server | Controller | DP | LG |
|-------|--------|-----------|----|----|
| Installed version = 26.1 | ✓ | ✓ | ✓ | ✓ |
| IIS running | ✓ | | | |
| LRE Backend Service running | ✓ | | | |
| LRE Alerts Service running | ✓ | | | |
| LoadRunner Agent Service running | | ✓ | ✓ | ✓ |
| Remote Mgmt Agent running | | ✓ | ✓ | ✓ |
| LRE web app accessible (HTTP check) | ✓ | | | |
| pcs.config service user consistent | ✓ | | | |
| lts.config service user consistent | | ✓ | ✓ | |
| appsettings.json consistent | ✓ | | | |
| No pending reboot | ✓ | ✓ | ✓ | ✓ |
| Temp credential files cleaned up | ✓ | ✓ | ✓ | ✓ |
| Recent Event Log errors | ✓ | ✓ | ✓ | ✓ |

### 10.3 Expected result

```
  RESULT: ALL CHECKS PASSED (13/13)
  <hostname> [Server] is successfully upgraded to 26.1.
```

### 10.4 Final check in LRE Administration

1. Log into LRE: `https://lre.yourdomain.com/LRE/`
2. Go to **Administration > Maintenance > Hosts**
3. Verify all hosts show version **26.1**
4. Check that Controller, Data Processor, and Load Generator hosts are
   all in **Active** status
5. Run a **smoke test** — create a test run with a simple Vuser script to
   confirm end-to-end functionality

---

## 11. Troubleshooting

### Public Key not captured

**Symptom:** Server script warns "Could not automatically retrieve Public Key"

**Resolution:**
1. Open a browser on the LRE Server and go to:
   ```
   http://localhost/Admin/rest/v1/configuration/getPublicKey
   ```
2. Copy the `PublicKey` value from the JSON response
3. Create the file manually:
   ```powershell
   Set-Content "\\fileserver\LRE_26.1\PublicKey.txt" -Value "<paste_key_here>"
   ```
4. Then run the host scripts — they will read from this file

Alternatively, set `$Global:PublicKeyValue` directly in `upgrade_config.ps1`.

---

### Installer exits with code 1603 (Fatal error)

**Symptom:** Script reports "Fatal install error (1603)"

**Resolution:**
1. Open the installer log: `C:\LRE_UpgradeLogs\LREServer_Installer_<date>.log`
2. Search for `Error` or `FAILED` to find the root cause
3. Common causes:
   - IIS not installed (Server machines)
   - A required service could not be stopped
   - Repository path does not exist or has wrong permissions
   - The system user password is incorrect

---

### Installer exits with code 3010 (Reboot required)

**Symptom:** Script shows "REBOOT REQUIRED" and asks to reboot

**Action:** Allow the reboot. After the machine restarts:
1. Log back in as Administrator
2. Re-run the same script — it will detect the upgrade is already installed
   and skip directly to post-upgrade steps

---

### Host shows "Reconfigure needed" in Administration

This is normal after an upgrade. In LRE Administration:
1. Go to **Maintenance > Hosts**
2. Select the affected host
3. Click **Reconfigure Host**
4. Wait for the reconfiguration to complete

---

### Service user mismatch warning

**Symptom:** Post-upgrade check warns service user not found in config files

**Resolution:**
Check the exact case of the user name across all three files:
```powershell
Select-String -Path "$ServerInstallDir\dat\pcs.config"                  -Pattern "IUSR"
Select-String -Path "$HostInstallDir\dat\lts.config"                    -Pattern "IUSR"
Select-String -Path "$FileSystemRoot\system_config\appsettings.json"    -Pattern "IUSR"
```

If cases differ (e.g. `IUSR_METRO` vs `iusr_metro`), update them to match.
See the vendor guide section **Upgrade Tips** for the exact files to edit.

---

### setup_server.exe not found on share

**Symptom:** Pre-check fails — "setup_server.exe not found"

**Resolution:**
1. Confirm the installer package was fully extracted to the share
2. Check the exact path set in `upgrade_config.ps1`:
   ```powershell
   $Global:SetupServerExe = "$InstallerShare\Setup\En\setup_server.exe"
   ```
3. If the vendor layout differs, update the path variable to match

---

### OneLG installer not found

**Symptom:** LG script reports "Neither SetupOneLG.exe nor OneLG_x64.msi found"

**Resolution:**
Check if the standalone application is in a different location on your share:
```powershell
Get-ChildItem "\\fileserver\LRE_26.1" -Recurse -Filter "SetupOneLG.exe"
Get-ChildItem "\\fileserver\LRE_26.1" -Recurse -Filter "OneLG_x64.msi"
```
Update `$Global:SetupOneLGExe` or `$Global:OneLGMsi` in `upgrade_config.ps1`.

---

### Script execution is blocked

**Symptom:** "cannot be loaded because running scripts is disabled"

**Resolution:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```
This sets the policy for the current PowerShell session only, without
changing the machine-wide policy.

---

## 12. Rollback Procedure

If the upgrade fails and you need to revert:

> **Prerequisite:** You must have the database backups from Step 3.

### 12.1 Restore databases

Ask your DBA to restore the three databases from the backups taken in Step 3:
- `LRE_LAB`
- `LRE_ADMIN`
- `LRE_SITE`

### 12.2 Uninstall 26.1

On each machine where the upgrade ran:

```powershell
# Server uninstall (silent):
msiexec.exe /uninstall "<installdir>\Setup\Install\Server\LRE_Server.msi" /qnb

# Host uninstall (silent):
msiexec.exe /uninstall "<installdir>\Setup\Install\Host\LoadRunner_x64.msi" /qnb

# Or via Control Panel:
# Control Panel > Programs and Features
# → "OpenText Enterprise Performance Engineering 26.1" > Uninstall
```

### 12.3 Reinstall 25.1

Re-run the 25.1 installer on each machine and follow the Configuration wizard
using your existing settings.

### 12.4 Restore repository

If the repository files were modified, restore from backup.

---

## 13. Configuration Reference

Full list of all settings in `config/upgrade_config.ps1`:

### Installer Media

| Setting | Description | Example |
|---------|-------------|---------|
| `InstallerShare` | UNC path to 26.1 installer | `\\fileserver\LRE_26.1` |
| `SetupServerExe` | Path to server setup executable | Auto-built from share |
| `SetupHostExe` | Path to host setup executable | Auto-built from share |
| `SetupOneLGExe` | Path to OneLG setup executable | Auto-built from share |

### Local Paths (per machine)

| Setting | Description | Default |
|---------|-------------|---------|
| `LogDir` | Upgrade log output folder | `C:\LRE_UpgradeLogs` |
| `ServerInstallDir` | LRE Server install directory | `C:\Program Files\OpenText\LRE` |
| `HostInstallDir` | Host install directory | `C:\Program Files\OpenText\Performance Center Host` |
| `OneLGInstallDir` | OneLG install directory | `C:\Program Files\OpenText\OneLG` |

### System User

| Setting | Description | Default |
|---------|-------------|---------|
| `UseDefaultUser` | `true` = use IUSR_METRO | `true` |
| `DomainName` | `.` = local, or domain name | `.` |
| `SystemUserName` | Windows account name | `IUSR_METRO` |
| `SystemUserPwd` | Password (blank = prompted) | `""` |

### Security

| Setting | Description |
|---------|-------------|
| `CryptoKey` | LW_CRYPTO_INIT_STRING — must match all nodes |
| `SecurePassphrase` | Communication passphrase (min 12 chars) — must match all nodes |

### Database

| Setting | Description |
|---------|-------------|
| `DbType` | `MS-SQL` |
| `DbServerHost` | SQL Server hostname or IP |
| `DbServerPort` | SQL port (default `1433`) |
| `DbAdminUser` | SQL admin account |
| `DbAdminPassword` | Admin password (blank = prompted) |
| `DbUsername` | LRE application DB user |
| `DbPassword` | App user password (blank = prompted) |
| `LabDbName` | Lab database schema name |
| `AdminDbName` | Admin database schema name |
| `SiteDbName` | Site Management schema name |

### IIS / TLS

| Setting | Description |
|---------|-------------|
| `IisWebSiteName` | IIS site name (default `Default Web Site`) |
| `IisSecureConfig` | Enable HTTPS binding (`true`/`false`) |
| `IisSecureHostName` | LRE server FQDN for the HTTPS binding |
| `IisSecurePort` | HTTPS port (default `443`) |
| `ImportCertificate` | `true` = import from file; `false` = use existing store cert |
| `CertificateStore` | Windows cert store name (default `My`) |
| `CertificateName` | Certificate CN or thumbprint |
| `CertificateFilePath` | Path to `.pfx` file (if `ImportCertificate = true`) |
| `CertificatePassword` | `.pfx` password (blank = prompted) |

### Repository

| Setting | Description |
|---------|-------------|
| `FileSystemRoot` | Fully qualified path to LRE repository |

### Service Names

| Setting | Description |
|---------|-------------|
| `ServiceIIS` | `W3SVC` |
| `ServiceBackend` | List of candidate service names checked in order |
| `ServiceAlerts` | List of candidate service names |
| `ServiceAgent` | List of candidate names for LoadRunner Agent Service |
| `ServiceRemoteMgmt` | List of candidate names for Remote Management Agent |

> If any service is not found by the default names, add the correct Windows
> service name to the relevant list in `upgrade_config.ps1`.
> To find the exact name: `Get-Service | Where-Object {$_.DisplayName -like "*LoadRunner*"}`

---

*Guide version 1.0 — LRE 25.1 to 26.1 upgrade — February 2026*
