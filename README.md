# Google Workspace → Microsoft 365 Migration

Controlled, enterprise-safe migration of **Gmail, Calendar, and Contacts** from Google Workspace to Microsoft 365 — a single PowerShell script, **no Google Cloud SDK, no gcloud, no command-line tools required**.

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

| File | Direct download |
|---|---|
| **Migration controller** | [GoogleToM365Migration.ps1](https://raw.githubusercontent.com/bluespam-cyber/google-to-m365-migration/main/scripts/GoogleToM365Migration.ps1) |
| **Launcher** | [Run-Migration.ps1](https://raw.githubusercontent.com/bluespam-cyber/google-to-m365-migration/main/Run-Migration.ps1) |
| **Sample CSV** | [migration-users.csv](https://raw.githubusercontent.com/bluespam-cyber/google-to-m365-migration/main/examples/migration-users.csv) |

> **Tip:** Download the launcher + the script into the **same folder** — the launcher auto-finds the script, so you can run it from anywhere.

---

## 🚀 Quick Start

### ✨ Interactive wizard (recommended) — just run it

Run with **no arguments** and the script walks you through everything. It auto-detects what it can (CSV, service-account key, Exchange connection, accepted domains) and asks only for what it can't:

```powershell
.\Run-Migration.ps1
```

The wizard offers a numbered menu of the modes below, pre-fills values it discovered, and asks a **Y/n** confirmation before any mutating step.

### 📖 Guide mode — the 5-minute Google setup

Before anything else, run the built-in guide. It prints the exact Google Cloud console steps (two browser tabs, ~5 minutes) and the Microsoft 365 prerequisites:

```powershell
.\Run-Migration.ps1 -Mode Guide
```

### 🧪 Preflight — read-only health check

```powershell
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv -KeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -TargetDeliveryDomain o365.contoso.com
```

Review the findings report. Resolve every **Critical** and **High** finding before proceeding.

---

## 📋 Modes

| Mode | Description | Read-Only? |
|---|---|---|
| `Guide` | Prints the Google Cloud + M365 setup steps | ✅ Yes |
| `Preflight` | Validates CSV, routing domain, recipients, Google connectivity | ✅ Yes |
| `CreateEndpoint` | Creates the Gmail migration endpoint (shreds the key afterwards) | ❌ Mutation |
| `CreateBatch` | Creates the migration batch (**never auto-starts it**) | ❌ Mutation |
| `StartBatch` | Starts an existing batch | ❌ Mutation |
| `Monitor` | Monitors batch progress, optional per-user statistics | ✅ Yes |
| `Complete` | Completes the batch (cutover) | ❌ Requires `-ApproveCutover` |
| `Run` | End-to-end: endpoint → batch → start in one session | ❌ Mutation |

> **Default mode is `Preflight`** — running the script with parameters but no `-Mode` is always read-only.

---

## ⚙️ Parameters

| Parameter | Type | Default | Purpose |
|---|---|---|---|
| `-Mode` | string | `Preflight` | One of: `Guide`, `Preflight`, `CreateEndpoint`, `CreateBatch`, `StartBatch`, `Monitor`, `Complete`, `Run` |
| `-KeyPath` | string | — | Path to the Google service-account JSON key (auto-discovered in Downloads/Desktop/cwd if omitted) |
| `-GoogleAdminEmail` | string | — | Google Workspace **Super Admin** email (required for delegation) |
| `-CsvPath` | string | — | Migration CSV (auto-discovered as `migration-users.csv` if omitted) |
| `-TargetDeliveryDomain` | string | — | Verified M365 routing subdomain, e.g. `o365.contoso.com` (**never** `tenant.onmicrosoft.com`) |
| `-EndpointName` | string | `GoogleWorkspaceEndpoint` | Name of the Gmail migration endpoint |
| `-BatchName` | string | `GoogleWorkspaceMigration` | Name of the migration batch |
| `-MonitorIntervalSeconds` | int | `60` | Poll interval for `Monitor` (15–3600) |
| `-MaxMonitorMinutes` | int | `60` | Max monitor duration (1–1440) |
| `-OutputRoot` | string | — | Where reports/state are written (default: current folder) |
| `-DelegationTestCount` | int | `3` | How many users to test delegation against (0–500) |
| `-ApproveCutover` | switch | off | Explicit approval required for `Complete` |
| `-KeepKeyFile` | switch | off | Keep the key file after use (default: securely shredded) |
| `-SkipGoogleTest` | switch | off | Skip the live Google delegation test |
| `-NonInteractive` | switch | off | Fail fast with a report instead of prompting (RMM/SYSTEM) |

---

## ☁️ Google Cloud Setup (one-time, ~5 minutes)

The script does **not** create projects, service accounts, or keys — you do this once in the Google Cloud console, then the script handles everything else. Run `-Mode Guide` to see these steps printed.

1. **Create a project** — [console.cloud.google.com](https://console.cloud.google.com) → project picker → **New Project**. Name it anything (e.g. `M365 Migration`). Note the **Project ID**.
2. **Enable these four APIs** — [console.cloud.google.com/apis/library](https://console.cloud.google.com/apis/library) → search each → **Enable**:
   - Gmail API
   - Google Calendar API
   - Contacts API *(easy to miss — contact migration fails without it)*
   - People API
3. **Create a service account** — [console.cloud.google.com/iam-admin/serviceaccounts](https://console.cloud.google.com/iam-admin/serviceaccounts) → **Create service account** → name it, skip optional roles, **Done**.
4. **Create a JSON key** — open the service account → **Keys** tab → **Add key** → **Create new key** → **JSON**. The downloaded file is your `-KeyPath`. It is a password-equivalent secret — **this script deletes it after use**.
5. **Copy the NUMERIC Client ID** — on the service account **Details** tab, copy the **Unique ID** (a long number; also the `client_id` field inside the JSON key). Using the service account *email* here is the single most common mistake.
6. **Authorize domain-wide delegation** — [admin.google.com](https://admin.google.com) → **Security** → **Access and data control** → **API controls** → **Manage Domain Wide Delegation** → **Add new**:
   - **Client ID:** the numeric ID from step 5
   - **OAuth scopes:** paste this entire line, exactly, no spaces:
     ```
     https://mail.google.com/,https://www.googleapis.com/auth/calendar,https://www.google.com/m8/feeds/,https://www.googleapis.com/auth/gmail.settings.sharing,https://www.googleapis.com/auth/contacts
     ```
   - All five scopes are required. An extra or missing scope breaks the match and the migration fails **after** the batch starts. Propagation takes minutes, occasionally up to 24 hours.

---

## 🏢 Microsoft 365 Prerequisites

1. **Routing subdomains** (both verified and Active):
   - `o365.<yourdomain>` — added in Google Admin as a **User alias domain**, MX pointing to Microsoft 365, accepted in M365. This is your `-TargetDeliveryDomain`.
   - `gsuite.<yourdomain>` — added in Google Admin as a **User alias domain**, MX pointing to Google. This is where MailUser `ExternalEmailAddress` values point.
   - **Do NOT use `tenant.onmicrosoft.com`** as the target delivery domain.
2. **Every user provisioned as a MailUser** (not a mailbox) before the batch is created:
   ```powershell
   New-MailUser -Name "Will" -ExternalEmailAddress will@gsuite.contoso.com -PrimarySmtpAddress will@contoso.com
   ```
   `ExternalEmailAddress` must point at the **Google** routing domain; each user also needs a proxy address at the **M365** routing domain. The script can auto-provision missing MailUsers from your CSV (see below).
3. **Disable MRM and archive policies** until migration completes — otherwise items are flagged "missing", which is very hard to separate from real loss during verification.
4. **Automatic Forwarding enabled** on the Remote Domain (Exchange Online default).

---

## 📄 CSV Format

Two columns:

| Column | Meaning |
|---|---|
| `EmailAddress` | The Microsoft 365 target email address |
| `Username` | The source Google Workspace address (optional — defaults to `EmailAddress` if omitted) |

```csv
EmailAddress,Username
john.doe@contoso.com,john.doe@gsuite.contoso.com
jane.smith@contoso.com,jane.smith@gsuite.contoso.com
```

See `examples/migration-users.csv` for a sample. The script auto-detects a file named `migration-users.csv` in the current folder, `examples/`, `config/`, your Downloads, or Documents.

---

## 🛡️ Safety-First Design

- **Defaults to read-only** — `Preflight` is the default mode; every mutation is an explicit mode choice
- **Never auto-starts a batch** — `CreateBatch` creates it for review; `StartBatch` is a separate explicit step
- **Never selects a routing domain automatically** — if the target isn't an accepted domain, the interactive picker shows candidates and applies your choice immediately
- **Key is shredded after use** — overwritten and deleted in a `finally` block unless `-KeepKeyFile`
- **Cutover requires `-ApproveCutover`** — completing the batch is never implicit
- **Blocking findings halt further action** — Critical/High findings stop the run with a report
- **Resume state** — progress is saved to `GoogleM365Migration\last-run.json` so interrupted runs can continue
- **Delegation wait & retry** — waits up to 30 minutes for domain-wide delegation to propagate, retrying only on `unauthorized_client`/`access_denied` with backoff
- **Auto-provisions missing MailUsers** — accounts that need a password are skipped and listed in `noAccount` rather than guessed

---

## 🧰 What's Included

```
google-to-m365-migration/
├── Run-Migration.ps1                  # Auto-locating launcher → migration controller
├── scripts/
│   └── GoogleToM365Migration.ps1      # The entire migration controller (all 8 modes)
├── docs/
│   ├── SETUP-GUIDE.md                 # Step-by-step setup
│   └── TROUBLESHOOTING.md             # Common errors & fixes
├── examples/
│   └── migration-users.csv            # Sample CSV format
├── config/
│   └── .env.example                   # Parameter reference sheet
└── LICENSE
```

---

## 🤖 Automation / Unattended Runs

Every mode can be called directly with parameters — perfect for scheduled or RMM runs:

```powershell
# 1. Preflight (read-only) — validate CSV, routing, recipients, Google connectivity
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv -KeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -TargetDeliveryDomain o365.contoso.com

# 2. Create the Gmail migration endpoint (key is shredded afterwards)
.\Run-Migration.ps1 -Mode CreateEndpoint -KeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com

# 3. Create the migration batch (does NOT auto-start)
.\Run-Migration.ps1 -Mode CreateBatch -CsvPath .\examples\migration-users.csv -TargetDeliveryDomain o365.contoso.com

# 4. Start the batch
.\Run-Migration.ps1 -Mode StartBatch

# 5. Monitor progress
.\Run-Migration.ps1 -Mode Monitor -CollectUserStatistics

# 6. Complete the batch (cutover) — only after Synced
.\Run-Migration.ps1 -Mode Complete -ApproveCutover

# Or do steps 2-4 in one session:
.\Run-Migration.ps1 -Mode Run -CsvPath .\examples\migration-users.csv -KeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -TargetDeliveryDomain o365.contoso.com
```

> **Unattended runs:** add `-NonInteractive` to fail fast with a report instead of prompting (for RMM/SYSTEM). The ExchangeOnlineManagement module is auto-installed if missing (TLS 1.2, NuGet, PSGallery restore, three strategies).

---

## 📄 License

MIT License — see [LICENSE](LICENSE)

## 👤 Author

**Arwaz Khan**

---

*"Automate the boring stuff. Document everything. Ship fast."*