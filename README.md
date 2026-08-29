# Google Workspace → Microsoft 365 Migration Toolkit

Controlled, enterprise-safe migration of **Gmail, Calendar, and Contacts** from Google Workspace to Microsoft 365 using PowerShell, Google Cloud SDK, and Exchange Online migration endpoints.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue) ![License](https://img.shields.io/badge/License-MIT-green) ![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## 🚀 Quick Start (auto-locating launchers)

The launchers **automatically find the scripts** — you can run them from any directory without `cd`-ing into the repo or typing full paths.

```powershell
# 1. Inspect your Google Cloud environment (read-only, safe first step)
.\Run-Bootstrap.ps1 -Mode Inspect

# 2. Run a read-only migration preflight
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv -SkipExchangeConnection
```

> **How auto-location works:** each launcher uses `$PSScriptRoot` to locate the `scripts/` folder relative to itself, then passes all your arguments through unchanged. Run them from the repo root, or from anywhere via the full path.

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
- **Never** selects a routing domain automatically
- Every mutation requires `ShouldProcess` confirmation
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

## 🗺️ Recommended Workflow

### Phase 1 — Google Cloud preparation (bootstrap)

```powershell
# Inspect (read-only) — see what's present
.\Run-Bootstrap.ps1 -Mode Inspect

# Install the Google Cloud SDK (opens official installer)
.\Run-Bootstrap.ps1 -Mode InstallSdk

# Enable required APIs (requires approval + existing project)
.\Run-Bootstrap.ps1 -Mode EnableApis -ProjectId <your-project> -ApproveApiEnablement

# Create the service account (requires approval)
.\Run-Bootstrap.ps1 -Mode CreateServiceAccount -ProjectId <your-project>

# Create a JSON key (requires approval)
.\Run-Bootstrap.ps1 -Mode CreateServiceAccountKey -ProjectId <your-project> -ApproveKeyCreation
```

> ⚠️ After creating the service account, complete the **manual domain-wide delegation** step in Google Admin Console using the numeric Client ID the script displays.

### Phase 2 — Migration (controller)

```powershell
# Preflight (read-only) — validate CSV, routing, recipients, Google connectivity
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv -ServiceAccountKeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -TargetDeliveryDomain o365.contoso.com

# Create the Gmail migration endpoint
.\Run-Migration.ps1 -Mode CreateEndpoint -ServiceAccountKeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -ApproveGooglePrerequisites

# Create the migration batch (does NOT auto-start)
.\Run-Migration.ps1 -Mode CreateBatch -CsvPath .\examples\migration-users.csv -TargetDeliveryDomain o365.contoso.com -ApproveGooglePrerequisites

# Start the batch
.\Run-Migration.ps1 -Mode StartBatch -ApproveGooglePrerequisites

# Monitor progress
.\Run-Migration.ps1 -Mode Monitor -CollectUserStatistics

# Complete the batch (cutover) — only after Synced
.\Run-Migration.ps1 -Mode CompleteBatch -ApproveCutover -ApproveGooglePrerequisites
```

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

**Arwaz Khan** — [arwazitbp2003@gmail.com](mailto:arwazitbp2003@gmail.com)

---

*"Automate the boring stuff. Document everything. Ship fast."*
