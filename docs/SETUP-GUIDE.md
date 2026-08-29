# Setup Guide — Google Workspace → Microsoft 365 Migration

Complete step-by-step guide to prepare your environment before running the migration scripts.

---

## 1. Prerequisites Checklist

### Microsoft 365 Requirements
- [ ] **Global Admin** or **Exchange Administrator** role in Microsoft 365
- [ ] PowerShell 5.1+ (Windows PowerShell)
- [ ] Internet access to `login.microsoftonline.com` and `outlook.office365.com`
- [ ] At least one licensed mailbox in the tenant

### Google Workspace Requirements
- [ ] **Google Workspace Super Admin** account
- [ ] Google Cloud account (free tier is sufficient)
- [ ] Ability to authorize **Domain-Wide Delegation** in Google Admin Console
- [ ] Gmail, Calendar, and Contacts enabled for users being migrated

---

## 2. Microsoft 365 Preparation

### 2.1 Verify Admin Access

```powershell
# Check if you have Exchange admin access
Get-Module -ListAvailable ExchangeOnlineManagement
```

If the module isn't installed:
```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
```

### 2.2 Verify Mailboxes Exist

```powershell
Connect-ExchangeOnline
Get-Mailbox -ResultSize Unlimited | Select Name, PrimarySmtpAddress
```

### 2.3 Provision Migration Targets as MailUsers

Each Google user being migrated must exist in Exchange Online as a **MailUser** (not a mailbox) before the migration batch is created.

```powershell
New-MailUser -Name "John Doe" -ExternalEmailAddress "john.doe@gsuite.com" -PrimarySmtpAddress "john.doe@contoso.com"
```

---

## 3. Google Cloud Preparation (Bootstrap)

> **Quick interactive path:** just run `.\Run-Bootstrap.ps1` with no arguments — it shows a menu, auto-detects your gcloud setup and projects, and asks only for what it can't find.

### 3.1 Inspect Your Environment (read-only)

```powershell
.\Run-Bootstrap.ps1 -Mode Inspect
```

### 3.2 Install the Google Cloud SDK

```powershell
.\Run-Bootstrap.ps1 -Mode InstallSdk
```

This opens the official installer. Follow the prompts and sign in with your Google Workspace Super Admin account.

### 3.3 Enable Required APIs

```powershell
.\Run-Bootstrap.ps1 -Mode EnableApis -ProjectId <your-project> -ApproveApiEnablement
```

This enables: Gmail API, Calendar API, People API, and Admin SDK API.

### 3.4 Create the Service Account

```powershell
.\Run-Bootstrap.ps1 -Mode CreateServiceAccount -ProjectId <your-project>
```

### 3.5 Create a Service Account Key

```powershell
.\Run-Bootstrap.ps1 -Mode CreateServiceAccountKey -ProjectId <your-project> -ApproveKeyCreation
```

The key is stored in a secure location with restricted ACLs. **Keep this path** — you'll need it for the migration controller.

### 3.6 Authorize Domain-Wide Delegation (Manual)

1. Go to **Google Admin Console** → **Security** → **API controls** → **Domain-wide delegation**
2. Click **Add new**
3. Enter the **numeric Client ID** displayed by the bootstrap script
4. Add these OAuth scopes:
   - `https://mail.google.com/`
   - `https://www.googleapis.com/auth/calendar`
   - `https://www.googleapis.com/auth/contacts`
   - `https://www.googleapis.com/auth/admin.directory.user.readonly`

---

## 4. Migration Controller

> **Quick interactive path:** just run `.\Run-Migration.ps1` with no arguments — it auto-finds your CSV and service-account key, then asks only for your Google admin email and routing domain.

### 4.1 Run a Preflight (read-only)

```powershell
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv -ServiceAccountKeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -TargetDeliveryDomain o365.contoso.com
```

Review the findings. Resolve any **Critical** or **High** findings before proceeding.

### 4.2 Create the Migration Endpoint

```powershell
.\Run-Migration.ps1 -Mode CreateEndpoint -ServiceAccountKeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -ApproveGooglePrerequisites
```

### 4.3 Create the Migration Batch

```powershell
.\Run-Migration.ps1 -Mode CreateBatch -CsvPath .\examples\migration-users.csv -TargetDeliveryDomain o365.contoso.com -ApproveGooglePrerequisites
```

The batch is created **without auto-start** so you can review it first.

### 4.4 Start the Batch

```powershell
.\Run-Migration.ps1 -Mode StartBatch -ApproveGooglePrerequisites
```

### 4.5 Monitor Progress

```powershell
.\Run-Migration.ps1 -Mode Monitor -CollectUserStatistics
```

### 4.6 Complete the Batch (Cutover)

Only after the batch reaches **Synced** status:

```powershell
.\Run-Migration.ps1 -Mode CompleteBatch -ApproveCutover -ApproveGooglePrerequisites
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
john.doe@contoso.com,john.doe@gmail.com
jane.smith@contoso.com,jane.smith@gmail.com
```

See `examples/migration-users.csv` for a sample. The script auto-detects a file named `migration-users.csv` in the current folder, `examples/`, `config/`, your Downloads, or Documents.

---

## 6. Environment Variables

Copy `config/.env.example` to `.env` and fill in your values. The scripts accept these as parameters too.

---

*"Automate the boring stuff. Document everything. Ship fast."*
