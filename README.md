# LRE Upgrade Scripts
### OpenText Enterprise Performance Engineering — 25.1 → 26.1

Three self-contained PowerShell scripts to silently upgrade all LRE components.
Copy the relevant files to the target machine and run. No shared config, no network dependencies beyond the installer media and cert share.

---

## Package Contents

| File | Purpose |
|------|---------|
| `Upgrade_LREServer.ps1` | Upgrades the LRE Server |
| `Upgrade_Host.ps1` | Upgrades a Controller or Data Processor host |
| `Upgrade_LoadGenerator.ps1` | Upgrades a OneLG Load Generator |
| `LRE_Server_DEV_UserInput.xml` | Silent install config for LRE Server — DEV |
| `LRE_Server_PROD_UserInput.xml` | Silent install config for LRE Server — PROD |
| `LRE_Host_DEV_UserInput.xml` | Silent install config for Host — DEV |
| `LRE_Host_PROD_UserInput.xml` | Silent install config for Host — PROD |
| `PC_Install.pdf` | Vendor installation guide (reference) |
| `LRE_Upgrade_SOE.xlsx` | Standard Operating Environment checklist |

---

## Prerequisites

- Run as **local Administrator** on each machine (scripts self-elevate if needed)
- Machine must be running LRE **25.1** (scripts verify this before proceeding)
- Minimum free disk space: **20 GB** on D:\ for Server, **10 GB** for Host/LG
- No pending Windows reboot (scripts check and abort if one is detected)
- PowerShell 5.1 or later (built into Windows Server 2016+)

---

## Upgrade Order

> **This order is mandatory. Do not upgrade Hosts or LGs before the Server.**

```
1. LRE Server
2. Controller Host(s)    ─┐
3. Data Processor Host(s) ├─ can run in any order / in parallel
4. Load Generator(s)     ─┘
```

The Server upgrade captures the **Public Key** and saves it to `D:\LRE_UpgradeLogs\PublicKey.txt`.
Copy that file to each Host machine before running `Upgrade_Host.ps1`.

---

## Before You Start

### 1 — Back up databases

Ask your DBA to take full backups of the three LRE SQL databases **before running any script**:
- `lre_default_lab_db_2023_new`
- `lre_siteadmin_db_2023_new`
- `lre_site_management_db_2023_new`

Do not proceed until backups are confirmed.

### 2 — Check the UserInput XMLs

Open the relevant XML for your environment and verify all values are correct, especially passwords.

For **PROD** — the `LRE_Server_PROD_UserInput.xml` and `LRE_Host_PROD_UserInput.xml` files contain `REPLACE_WITH_PROD_*` placeholders. Fill these in before use.

---

## Step 1 — Upgrade the LRE Server

**Files needed on the Server machine:**
- `Upgrade_LREServer.ps1`
- `LRE_Server_DEV_UserInput.xml` (or PROD)

**Copy files to the server, then run:**

```powershell
# DEV:
.\Upgrade_LREServer.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Setup\En\setup_server.exe" -Environment DEV

# PROD:
.\Upgrade_LREServer.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Setup\En\setup_server.exe" -Environment PROD
```

**What the script does automatically:**

| Phase | Action |
|-------|--------|
| 1 | Verifies installer exists, checks version is 25.1, checks disk space and pending reboot |
| 2 | Stops IIS, LRE Backend Service, LRE Alerts Service |
| 3 | Copies UserInput.xml to temp, replaces `{{HOSTNAME}}` with actual hostname |
| 4 | Runs `setup_server.exe /s` with progress monitoring (30–60 min) |
| 5 | Reads last 30 lines of Configuration Wizard log, warns on errors |
| 6 | Restarts services, waits 30 s, captures Public Key via REST API → saves to `D:\LRE_UpgradeLogs\PublicKey.txt` |
| 7.1 | Swaps `web.config` with `web.config-for-ssl` (TLS) |
| 7.2 | Updates PCS.config: `internalUrl` → HTTPS, `ltopIsSecured` → true |
| 7.3 | Replaces CA cert and TLS cert from cert share, runs `gen_cert.exe -verify` |
| 7.4 | Configures `AWS ActiveRegions: [ "eu-west-2" ]` in appsettings.defaults.json |
| 7.5 | Restarts Backend, Alerts, IIS |
| 8 | Verifies web.config SSL, PCS.config HTTPS, HTTPS connectivity, installed version |
| 9 | Prints summary with log paths and remaining manual steps |

**After the script completes:**

1. Note the Public Key file location: `D:\LRE_UpgradeLogs\PublicKey.txt`
2. Copy `PublicKey.txt` to each Host machine you plan to upgrade
3. Verify the LRE web app is accessible: `https://<server-fqdn>/LRE/`
4. Complete the remaining manual steps listed in the summary output

---

## Step 2 — Upgrade Controller / Data Processor Hosts

**Files needed on each Host machine:**
- `Upgrade_Host.ps1`
- `LRE_Host_DEV_UserInput.xml` (or PROD)
- `PublicKey.txt` (copied from the LRE Server `D:\LRE_UpgradeLogs\PublicKey.txt`)

**Copy all three files to the host machine, then run:**

```powershell
# Controller:
.\Upgrade_Host.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Setup\En\setup_host.exe" -ComponentType Controller

# Data Processor:
.\Upgrade_Host.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Setup\En\setup_host.exe" -ComponentType DataProcessor
```

> If `PublicKey.txt` is not present in the same folder as the script, you will
> be prompted to paste the key manually.

**What the script does automatically:**

| Phase | Action |
|-------|--------|
| 1 | Verifies installer, version check (25.1), disk space, pending reboot |
| 2 | Reads `PublicKey.txt` from script folder — prompts if not found |
| 3 | Stops LoadRunner Agent Service, Remote Management Agent |
| 4 | Copies UserInput.xml to temp, replaces `{{HOSTNAME}}` and `{{PUBLIC_KEY}}` |
| 5 | Runs `setup_host.exe /s LRASPCHOST=1` (20–40 min) |
| 6 | Restarts Agent and Remote Management services |
| 7.1 | Swaps `LTOPSvc.exe.config` with SSL version |
| 7.2 | Replaces CA cert and TLS cert from cert share, runs `gen_cert.exe -verify` |
| 7.3 | Runs `lr_agent_settings.exe -check_client_cert 1 -restart_agent` |
| 8 | Checks lts.config, verifies installed version |
| 9 | Prints summary with next steps |

**After the script completes:**

1. In **LRE Administration > Maintenance > Hosts**, find this host
2. If it shows **"Reconfigure needed"** — click **Reconfigure Host**
3. Verify the version shows **26.1**

---

## Step 3 — Upgrade Load Generators

**Files needed on each Load Generator machine:**
- `Upgrade_LoadGenerator.ps1` only (no XML required)

**Copy the script to the LG machine, then run:**

```powershell
.\Upgrade_LoadGenerator.ps1 -InstallerPath "D:\Softwares\OpenText_Enterprise_PE_26.1_MLU\Standalone Applications\SetupOneLG.exe"
```

**What the script does automatically:**

| Phase | Action |
|-------|--------|
| 1 | Verifies installer, version check (25.1, LoadRunner, OneLG), disk space, pending reboot |
| 2 | Stops LoadRunner Agent Service, Remote Management Agent |
| 3 | Runs `SetupOneLG.exe -s` with service mode flags (20–40 min) |
| 4 | Restarts Agent and Remote Management services |
| 5.1 | Replaces CA cert and TLS cert from cert share, runs `gen_cert.exe -verify` |
| 6 | Verifies agent is running as a Windows service, checks installed version |
| 7 | Prints summary with next steps |

**After the script completes:**

1. In **LRE Administration > Maintenance > Hosts**, find this LG
2. If it shows **"Reconfigure needed"** — click **Reconfigure Host**
3. Set **Enable SSL = True** for this Load Generator in the Admin Portal
4. Verify the version shows **26.1**

---

## Logs

All logs are written to `D:\LRE_UpgradeLogs\` on each machine:

| File | Contents |
|------|----------|
| `LREServer_Upgrade_<host>_<date>.log` | Server upgrade main log |
| `LREHost_<type>_Upgrade_<host>_<date>.log` | Host upgrade main log |
| `LoadGenerator_Upgrade_<host>_<date>.log` | LG upgrade main log |
| `*_transcript.log` | Full PowerShell transcript alongside each log |
| `PublicKey.txt` | LRE Public Key (written by Server upgrade, needed by Host upgrade) |

---

## Cert Share Layout Expected

Scripts read certificates from:
```
\\rbsres01\grpareas\LRE\LREadmin\LRE_CERTS\
```

Expected certificate filenames (cert share must contain these):

| Component | CA Cert | TLS Cert |
|-----------|---------|---------|
| Server DEV | `lre_dev_cacert.cer` | `lre_dev_<hostname>_iis_dev.cer` |
| Server PROD | `lre_prod_cacert.cer` | `lre_prod_<hostname>_iis_server.cer` |
| Host (Controller) | `lre_dev_cacert.cer` | `lre_dev_<hostname>_controller.cer` |
| Host (Data Processor) | `lre_dev_cacert.cer` | `lre_dev_<hostname>_dp.cer` |
| Load Generator | `lre_dev_cacert.cer` | `lre_dev_<hostname>_lg.cer` |

If the exact filename is not found, scripts fall back to a wildcard `*<hostname>*.cer` match and log a warning.

---

## Troubleshooting

### Public Key not captured automatically

After the server upgrade, if `D:\LRE_UpgradeLogs\PublicKey.txt` is empty or missing:

1. On the LRE Server, open a browser and go to:
   `http://localhost/Admin/rest/v1/configuration/getPublicKey`
2. Copy the `PublicKey` value from the JSON response
3. Save it to `D:\LRE_UpgradeLogs\PublicKey.txt`
4. Copy that file to each Host machine before running `Upgrade_Host.ps1`

---

### Installer exits with code 1603 (Fatal error)

1. Check the installer log in `D:\LRE_UpgradeLogs\`
2. Search for `Error` or `FAILED` to find the root cause
3. Common causes: IIS not installed, service could not be stopped, wrong password in UserInput.xml, insufficient disk space

---

### Installer exits with code 3010 (Reboot required)

The script will reboot the machine automatically after 30 seconds. After the reboot, re-run the same script — the upgrade itself is already complete; the script will resume from post-install steps.

---

### Host shows "Reconfigure needed"

This is normal after an upgrade. In LRE Administration > Maintenance > Hosts, select the host and click **Reconfigure Host**.

---

### Script execution blocked

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```

This applies to the current session only and does not change machine policy.

---

### Service not found / wrong service name

The scripts try multiple candidate names for each service to handle vendor renames between versions. If a service is still not found, find its exact name and add it to the candidates array near the top of the script:

```powershell
Get-Service | Where-Object { $_.DisplayName -like "*LoadRunner*" -or $_.DisplayName -like "*OpenText*" }
```

---

## Manual Steps (after all upgrades)

These cannot be automated and must be completed in the **LRE Admin Portal**:

1. Update database authentication type (run `07_PostInstall_DBConfig.ps1` if available)
2. Verify Site Admin login: `https://<lre-server>/adminx/login`
3. Update LDAP configuration in Admin Portal
4. Change authentication method from Site Admin to LDAP
5. Update Server URLs in Admin Portal
6. Upgrade projects in Admin Portal
7. Upload license file
8. Verify AWS Cloud account, proxy settings, and templates
9. Reconfigure all hosts in Admin Portal (if not done automatically)
10. Run a smoke test — create a test run with a simple Vuser script
