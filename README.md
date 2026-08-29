# Google Workspace → Microsoft 365 Migration Toolkit

Controlled, enterprise-safe migration of **Gmail, Calendar, and Contacts** from Google Workspace to Microsoft 365 using PowerShell, Google Cloud SDK, and Exchange Online migration endpoints.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue) ![License](https://img.shields.io/badge/License-MIT-green) ![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## ⬇️ Download & Run (2 minutes)

### Option A — Download the whole repo (recommended)

Click the green **Code** button above, then **Download ZIP**. Extract and open a PowerShell window in the folder.

**Or download the ZIP directly:**
```
https://github.com/bluespam-cyber/google-to-m365-migration/archive/refs/heads/main.zip
```

### Option B — Download just the scripts you need

| Script | Direct download |
|---|---|
| **Google Cloud prep** (bootstrap) | [GoogleCloudMigrationBootstrap.ps1](https://raw.githubusercontent.com/bluespam-cyber/google-to-m365-migration/main/scripts/GoogleCloudMigrationBootstrap.ps1) |
| **Migration controller** | [GoogleToM365Migration.ps1](https://raw.githubusercontent.com/bluespam-cyber/google-to-m365-migration/main/scripts/GoogleToM365Migration.ps1) |
| **Launcher — bootstrap** | [Run-Bootstrap.ps1](https://raw.githubusercontent.com/bluespam-cyber/google-to-m365-migration/main/Run-Bootstrap.ps1) |
| **Launcher — migration** | [Run-Migration.ps1](https://raw.githubusercontent.com/bluespam-cyber/google-to-m365-migration/main/Run-Migration.ps1) |
| **Sample CSV** | [migration-users.csv](https://raw.githubusercontent.com/bluespam-cyber/google-to-m365-migration/main/examples/migration-users.csv) |

> **Tip:** Download the launcher + its matching script into the **same folder** — the launcher auto-finds the script, so you can run it from anywhere.

---

## 🚀 Run the scripts

### ✨ Fully interactive (recommended) — just run it

Run with **no arguments** and the script walks you through everything. It **auto-detects** what it can and **asks only for what it can't**:

- **Bootstrap** auto-detects: gcloud, your active Google account, GCP projects, existing service accounts
- **Migration** auto-detects: your CSV and your service-account key; asks only for your Google admin email and routing domain

```powershell
# Google Cloud prep — interactive menu (pick a mode, answer only what's needed)
.\Run-Bootstrap.ps1

# Migration — interactive menu (auto-finds CSV + key, asks only for admin email & routing domain)
.\Run-Migration.ps1
```

> **What it will ask you:** only the things it genuinely cannot discover — e.g. your GCP project ID (if you have several), your Google admin email (if gcloud isn't signed in), your routing domain (if Exchange can't be reached), and a **Y/n** confirmation before any mutating step. Everything else is found automatically.

### Explicit mode (for automation / RMM)

Every mode can also be called directly with parameters — perfect for scheduled or unattended runs:

### Google Cloud prep (bootstrap)

```powershell
# 1. Inspect your Google Cloud environment (read-only, safe first step)
.\Run-Bootstrap.ps1 -Mode Inspect

# 2. Install the Google Cloud SDK (opens official installer)
.\Run-Bootstrap.ps1 -Mode InstallSdk

# 3. Enable required APIs (needs approval + existing project)
.\Run-Bootstrap.ps1 -Mode EnableApis -ProjectId <your-project> -ApproveApiEnablement

# 4. Create the service account
.\Run-Bootstrap.ps1 -Mode CreateServiceAccount -ProjectId <your-project>

# 5. Create a JSON key (needs approval)
.\Run-Bootstrap.ps1 -Mode CreateServiceAccountKey -ProjectId <your-project> -ApproveKeyCreation
```

### Migration (controller)

```powershell
# 1. Preflight (read-only) — validate CSV, routing, recipients, Google connectivity
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv -ServiceAccountKeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -TargetDeliveryDomain o365.contoso.com

# 2. Create the Gmail migration endpoint
.\Run-Migration.ps1 -Mode CreateEndpoint -ServiceAccountKeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -ApproveGooglePrerequisites

# 3. Create the migration batch (does NOT auto-start)
.\Run-Migration.ps1 -Mode CreateBatch -CsvPath .\examples\migration-users.csv -TargetDeliveryDomain o365.contoso.com -ApproveGooglePrerequisites

# 4. Start the batch
.\Run-Migration.ps1 -Mode StartBatch -ApproveGooglePrerequisites

# 5. Monitor progress
.\Run-Migration.ps1 -Mode Monitor -CollectUserStatistics

# 6. Complete the batch (cutover) — only after Synced
.\Run-Migration.ps1 -Mode CompleteBatch -ApproveCutover -ApproveGooglePrerequisites
```

> **How auto-location works:** each launcher uses `$PSScriptRoot` to locate the `scripts/` folder relative to itself, then passes all your arguments through unchanged. Keep the launcher and its matching script in the same folder.
>
> **Unattended runs:** add `-NonInteractive` to fail fast with a report instead of prompting (for RMM/SYSTEM).

---

## 📖 Overview

This toolkit is split into **two controlled scripts** that default to **read-only** and never take destructive action without explicit approval:

| Script | Purpose | Default Mode |
|---|---|---|
| `Run-Bootstrap.ps1` → `GoogleCloudMigrationBootstrap.ps1` | Safe Google Cloud preparation | `Inspect` (read-only) |
| `Run-Migration.ps1` → `GoogleToM365Migration.ps1` | Exchange Online migration controller | `Preflight` (read-only) |

### Safety-first design
- **Never** auto-creates a GCP project or service-account key without explicit `-Approve*` flags
- **Never** runs browser authentication automatically
- Auto-detected values (project, CSV, key, admin email, routing domain) are always shown and confirmed before use
- Every mutation requires `ShouldProcess` confirmation (or an interactive **Y/n** prompt)
- Blocking findings halt further action

---

## 🧰 What's Included

| File | Description |
|---|---|
| `Run-Bootstrap.ps1` | **Auto-locating launcher** for the Google Cloud bootstrap script |
| `Run-Migration.ps1` | **Auto-locating launcher** for the migration controller |
| `scripts/GoogleCloudMigrationBootstrap.ps1` | Google Cloud prep: inspect, SDK, APIs, service account, keys |
| `scripts/GoogleToM365Migration.ps1` | Migration controller: preflight, endpoint, batch, monitor, complete |
| `docs/SETUP-GUIDE.md` | Step-by-step setup guide with prerequisites |
| `docs/TROUBLESHOOTING.md` | Common errors and fixes |
| `examples/migration-users.csv` | Sample CSV format for user mapping |
| `config/.env.example` | Environment variable template |

---

## 🔧 Prerequisites

### Microsoft 365
- **Exchange Online** admin access (Global Admin or Exchange Administrator role)
- **PowerShell 5.1+** on Windows
- Internet access to `login.microsoftonline.com` and `outlook.office365.com`

### Google Workspace
- **Google Workspace Super Admin** account
- **Google Cloud** account (free tier works)
- Ability to authorize **Domain-Wide Delegation** in Google Admin Console
- A **service-account JSON key** stored in an approved secure location

---

## 📋 Bootstrap Modes

| Mode | Description | Read-Only? |
|---|---|---|
| `Inspect` | Inspect gcloud, active account, project, service account | ✅ Yes |
| `InstallSdk` | Open official Google Cloud SDK installer | ⚠️ Opens browser |
| `EnableApis` | Enable Gmail, Calendar, People, Admin APIs | ❌ Requires `-ApproveApiEnablement` |
| `CreateServiceAccount` | Create the migration service account | ❌ Requires approval |
| `CreateServiceAccountKey` | Create a JSON private key | ❌ Requires `-ApproveKeyCreation` |
| `ListKeys` | List user-managed keys | ✅ Yes |
| `DisableKey` | Disable a service-account key | ❌ Requires `-ApproveKeyDisable` |

---

## 📋 Migration Modes

| Mode | Description | Read-Only? |
|---|---|---|
| `Preflight` | Validate CSV, routing domain, recipients, Google connectivity | ✅ Yes |
| `CreateEndpoint` | Create the Gmail migration endpoint | ❌ Requires `-ApproveGooglePrerequisites` |
| `CreateBatch` | Create a migration batch (no auto-start) | ❌ Requires `-ApproveGooglePrerequisites` |
| `StartBatch` | Start an existing batch | ❌ Requires `-ApproveGooglePrerequisites` |
| `Monitor` | Monitor batch progress with optional per-user stats | ✅ Yes |
| `CompleteBatch` | Complete the batch (cutover) | ❌ Requires `-ApproveCutover` |
| `Report` | Generate a report without connecting | ✅ Yes |

---

## 📄 License

MIT License — see [LICENSE](LICENSE)

## 👤 Author

**Arwaz Khan**

---

*"Automate the boring stuff. Document everything. Ship fast."*
