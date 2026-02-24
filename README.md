# LRE Upgrade Automation Scripts
### OpenText Enterprise Performance Engineering — 25.1 → 26.1

Automated PowerShell scripts to silently upgrade all LRE components on Windows
machines with a single script per component type.

---

## What This Does

| Script | Component | Installer Used |
|--------|-----------|---------------|
| `01_Upgrade_LREServer.ps1` | LRE Server | `setup_server.exe /s` |
| `02_Upgrade_Controller.ps1` | Controller Host | `setup_host.exe /s` |
| `03_Upgrade_DataProcessor.ps1` | Data Processor Host | `setup_host.exe /s` |
| `04_Upgrade_LoadGenerator.ps1` | Load Generator (OneLG) | `SetupOneLG.exe -s` |
| `00_PreUpgradeChecks.ps1` | All machines | Read-only diagnostic |
| `05_PostUpgradeVerify.ps1` | All machines | Read-only health check |

> **Controller** and **Data Processor** use the same host installer —
> their role/purpose assignment in LRE Administration is preserved across upgrade.

---

## Folder Structure

```
LRE-Automates-Scripts/
├── config/
│   ├── upgrade_config.ps1          ← EDIT THIS FIRST — all settings in one place
│   ├── LRE_Server_UserInput.xml    ← Fallback template (Server silent install)
│   └── LRE_Host_UserInput.xml      ← Fallback template (Host silent install)
│
├── scripts/
│   ├── Common-Functions.ps1        ← Shared helpers (logging, XML merge, services)
│   ├── Shared-HostUpgrade.ps1      ← Shared host logic (Controller + DP reuse it)
│   ├── 00_PreUpgradeChecks.ps1
│   ├── 01_Upgrade_LREServer.ps1
│   ├── 02_Upgrade_Controller.ps1
│   ├── 03_Upgrade_DataProcessor.ps1
│   ├── 04_Upgrade_LoadGenerator.ps1
│   └── 05_PostUpgradeVerify.ps1
│
├── docs/
│   └── UserGuide.md                ← Step-by-step guide for the full upgrade
│
└── logs/                           ← Created at runtime on each machine
    └── (upgrade logs + transcripts)
```

---

## Quick Start

### 1 — Edit the config file

Open `config/upgrade_config.ps1` and fill in every line marked `<-- UPDATE THIS`.

Minimum required settings:

```powershell
$Global:InstallerShare   = "\\fileserver\LRE_26.1"
$Global:DbServerHost     = "sql-server-01"
$Global:LabDbName        = "LRE_LAB"
$Global:AdminDbName      = "LRE_ADMIN"
$Global:SiteDbName       = "LRE_SITE"
$Global:IisSecureHostName = "lre.yourdomain.com"
$Global:FileSystemRoot   = "C:\LRE_Repository"
$Global:SecurePassphrase = "YourPassphrase123"
```

### 2 — Run pre-checks on every machine

```powershell
# On each machine (as Administrator):
.\scripts\00_PreUpgradeChecks.ps1
```

### 3 — Back up databases

Ask your DBA to back up: `LRE_LAB`, `LRE_ADMIN`, `LRE_SITE` on your SQL Server.

### 4 — Upgrade in order

```
LRE Server  →  Controller(s)  →  Data Processor(s)  →  Load Generator(s)
```

### 5 — Verify

```powershell
# On each machine after upgrade:
.\scripts\05_PostUpgradeVerify.ps1
```

---

## UserInput.xml Strategy

Each upgrade script uses a **two-path approach** for the silent install XML:

1. **Preferred** — Copy the vendor's `UserInput.xml` directly from the installer
   share and update only the property values defined in `upgrade_config.ps1`.
   This is the safest approach as it preserves any installer-specific settings.

2. **Fallback** — If the share XML is not accessible, use the local template in
   `config/` and substitute `%%PLACEHOLDER%%` tokens.

No passwords are stored in any file on disk. Temp XML files written during
install are **deleted automatically** after the installer exits.

---

## Prerequisites

- PowerShell 5.1 or later (built into Windows Server 2016+)
- Run as **local Administrator** on each machine
- Network access to the UNC installer share from each machine
- Machines already running LRE **25.1** (installer detects version and upgrades)

---

## Logs

All logs are written to `C:\LRE_UpgradeLogs\` on each machine:

| File | Contents |
|------|----------|
| `LREServer_Upgrade_<host>_<date>.log` | Main server upgrade log |
| `Controller_Upgrade_<host>_<date>.log` | Controller host log |
| `DataProcessor_Upgrade_<host>_<date>.log` | Data Processor log |
| `LoadGenerator_Upgrade_<host>_<date>.log` | Load Generator log |
| `*_transcript.log` | Full PowerShell transcript alongside each log |
| `*_Installer_<date>.log` | Raw output from the installer executable |

---

## See Also

- [Step-by-step User Guide](docs/UserGuide.md)
- OpenText LRE 26.1 Installation Guide (`PC_Install.pdf` — vendor documentation)
- OpenText Help Center: https://admhelp.microfocus.com/lre/
