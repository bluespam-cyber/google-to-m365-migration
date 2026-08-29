# Setup Guide — Google Workspace → Microsoft 365 Migration

Complete step-by-step guide to prepare your environment and run the migration. The migration controller is a **single PowerShell script** — no Google Cloud SDK, no gcloud, no command-line tools required.

---

## 1. Prerequisites Checklist

### Microsoft 365 Requirements
- [ ] **Global Admin** or **Exchange Administrator** role in Microsoft 365
- [ ] PowerShell 5.1+ (Windows PowerShell)
- [ ] Internet access to `login.microsoftonline.com` and `outlook.office365.com`
- [ ] At least one licensed mailbox in the tenant
- [ ] Two **routing subdomains** verified and Active (see §2.2)

### Google Workspace Requirements
- [ ] **Google Workspace Super Admin** account
- [ ] Google Cloud account (free tier is sufficient)
- [ ] Ability to authorize **Domain-Wide Delegation** in Google Admin Console
- [ ] Gmail, Calendar, and Contacts enabled for users being migrated

---

## 2. Microsoft 365 Preparation

### 2.1 Verify Admin Access

```powershell
Get-Module -ListAvailable ExchangeOnlineManagement
```

If the module isn't installed, the migration script **auto-installs it** (TLS 1.2, NuGet, PSGallery restore, three strategies). To install manually:

```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
```

### 2.2 Configure Routing Subdomains

You need **two** routing subdomains, both verified and Active:

| Subdomain | MX points to | Purpose |
|---|---|---|
| `o365.<yourdomain>` | Microsoft 365 | Target delivery domain (`-TargetDeliveryDomain`) |
| `gsuite.<yourdomain>` | Google | Source routing domain (MailUser `ExternalEmailAddress`) |

1. In **Google Admin Console** → **Domains** → **Add a domain alias**, add both subdomains as **User alias domains**.
2. Point the `o365.` subdomain's MX records at Microsoft 365.
3. In **Microsoft 365** → **Domains**, add and verify both subdomains (they must show **Active**).
4. **Never use `tenant.onmicrosoft.com`** as the target delivery domain — the script blocks it.

### 2.3 Provision Migration Targets as MailUsers

Each Google user being migrated must exist in Exchange Online as a **MailUser** (not a mailbox) before the migration batch is created:

```powershell
New-MailUser -Name "John Doe" -ExternalEmailAddress "john.doe@gsuite.contoso.com" -PrimarySmtpAddress "john.doe@contoso.com"
```

- `ExternalEmailAddress` must point at the **Google** routing domain (`gsuite.contoso.com`)
- Each user also needs a **proxy address** at the M365 routing domain (`o365.contoso.com`)

> The migration script can **auto-provision missing MailUsers** from your CSV. Accounts that would need a password are skipped and listed in a `noAccount` report rather than guessed.

### 2.4 Disable MRM and Archive Policies

Until the migration completes, disable **MRM (retention) and archive policies** on the target mailboxes. Otherwise migrated items get flagged "missing" — perceived data loss that is very hard to separate from real loss during verification.

### 2.5 Verify Automatic Forwarding

Confirm **Automatic Forwarding** is enabled on the Remote Domain (this is the Exchange Online default; verify it wasn't disabled by a security policy).

---

## 3. Google Cloud Preparation (one-time, ~5 minutes)

The script **never** creates projects, service accounts, or keys — you do this once in the Google Cloud console. Run `.\Run-Migration.ps1 -Mode Guide` to see these exact steps printed.

### 3.1 Create a Project

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Use the **project picker** → **New Project**
3. Name it anything (e.g. `M365 Migration`) and note the **Project ID**

### 3.2 Enable the Four Required APIs

Go to [console.cloud.google.com/apis/library](https://console.cloud.google.com/apis/library), search each API, and click **Enable**:

| API | Service ID |
|---|---|
| Gmail API | `gmail.googleapis.com` |
| Google Calendar API | `calendar-json.googleapis.com` |
| Contacts API | `contacts.googleapis.com` |
| People API | `people.googleapis.com` |

> **Contacts API is easy to miss** — contact migration fails without it.

### 3.3 Create a Service Account

1. Go to [console.cloud.google.com/iam-admin/serviceaccounts](https://console.cloud.google.com/iam-admin/serviceaccounts)
2. Click **Create service account**
3. Name it (e.g. `m365-migration`), click **Create**, skip the optional role/access steps, click **Done**

### 3.4 Create a JSON Key

1. Open the service account → **Keys** tab → **Add key** → **Create new key** → **JSON**
2. The file downloads — that file is your `-KeyPath`
3. Store it in a secure location with restricted ACLs

> ⚠️ **This file is a password-equivalent secret.** The migration script **deletes it after use** (overwrite + delete) unless you pass `-KeepKeyFile`.

### 3.5 Copy the NUMERIC Client ID

1. On the service account **Details** tab, copy the **Unique ID** (a long number)
2. It is also the `client_id` field inside the JSON key
3. ⚠️ Using the service account **email** here is the single most common mistake

### 3.6 Authorize Domain-Wide Delegation

1. Go to [admin.google.com](https://admin.google.com) → **Security** → **Access and data control** → **API controls**
2. Click **Manage Domain Wide Delegation** → **Add new**
3. **Client ID:** the numeric ID from §3.5
4. **OAuth scopes:** paste this entire line, exactly, no spaces:

```
https://mail.google.com/,https://www.googleapis.com/auth/calendar,https://www.google.com/m8/feeds/,https://www.googleapis.com/auth/gmail.settings.sharing,https://www.googleapis.com/auth/contacts
```

All five scopes are required. An extra or missing scope breaks the match and the migration fails **after** the batch starts, not before. Propagation takes minutes, occasionally up to 24 hours — the script waits and retries for up to 30 minutes.

---

## 4. Run the Migration

### 4.1 Interactive Wizard (recommended)

```powershell
.\Run-Migration.ps1
```

The wizard auto-detects your CSV and service-account key, connects to Exchange Online, and asks only for what it can't find (admin email, routing domain). It shows a numbered menu of modes and asks **Y/n** before any mutating step.

### 4.2 Preflight (read-only)

```powershell
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv -KeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -TargetDeliveryDomain o365.contoso.com
```

Review the findings report. Resolve every **Critical** and **High** finding before proceeding.

### 4.3 Create the Migration Endpoint

```powershell
.\Run-Migration.ps1 -Mode CreateEndpoint -KeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com
```

The key is shredded after the endpoint is created.

### 4.4 Create the Migration Batch

```powershell
.\Run-Migration.ps1 -Mode CreateBatch -CsvPath .\examples\migration-users.csv -TargetDeliveryDomain o365.contoso.com
```

The batch is created **without auto-start** so you can review it first.

### 4.5 Start the Batch

```powershell
.\Run-Migration.ps1 -Mode StartBatch
```

### 4.6 Monitor Progress

```powershell
.\Run-Migration.ps1 -Mode Monitor -CollectUserStatistics
```

### 4.7 Complete the Batch (Cutover)

Only after the batch reaches **Synced** status:

```powershell
.\Run-Migration.ps1 -Mode Complete -ApproveCutover
```

### 4.8 Run Everything in One Session

```powershell
.\Run-Migration.ps1 -Mode Run -CsvPath .\examples\migration-users.csv -KeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -TargetDeliveryDomain o365.contoso.com
```

---

## 5. CSV Format

The migration CSV requires two columns:

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

## 6. Parameters Reference

| Parameter | Type | Default | Purpose |
|---|---|---|---|
| `-Mode` | string | `Preflight` | `Guide`, `Preflight`, `CreateEndpoint`, `CreateBatch`, `StartBatch`, `Monitor`, `Complete`, `Run` |
| `-KeyPath` | string | — | Path to the Google service-account JSON key |
| `-GoogleAdminEmail` | string | — | Google Workspace Super Admin email |
| `-CsvPath` | string | — | Migration CSV path |
| `-TargetDeliveryDomain` | string | — | Verified M365 routing subdomain (e.g. `o365.contoso.com`) |
| `-EndpointName` | string | `GoogleWorkspaceEndpoint` | Migration endpoint name |
| `-BatchName` | string | `GoogleWorkspaceMigration` | Migration batch name |
| `-MonitorIntervalSeconds` | int | `60` | Poll interval (15–3600) |
| `-MaxMonitorMinutes` | int | `60` | Max monitor duration (1–1440) |
| `-OutputRoot` | string | — | Report/state output folder |
| `-DelegationTestCount` | int | `3` | Users to test delegation against (0–500) |
| `-ApproveCutover` | switch | off | Required for `Complete` |
| `-KeepKeyFile` | switch | off | Keep the key file after use |
| `-SkipGoogleTest` | switch | off | Skip the live Google delegation test |
| `-NonInteractive` | switch | off | Fail fast instead of prompting |

---

*"Automate the boring stuff. Document everything. Ship fast."*