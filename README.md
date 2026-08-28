# Google Workspace → Microsoft 365 Migration Toolkit

Automated, no-touch migration of **Gmail, Calendar, and Contacts** from Google Workspace to Microsoft 365 using PowerShell, Google Cloud SDK, and Exchange Online migration endpoints.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue) ![License](https://img.shields.io/badge/License-MIT-green) ![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## 🚀 Overview

This toolkit automates the entire Google Workspace → Microsoft 365 email migration process:

1. **Auto-deploys** Google Cloud SDK (if not installed)
2. **Creates** a GCP project, service account, and RSA key
3. **Enables** Gmail, Calendar, People, and Admin SDK APIs
4. **Connects** to Exchange Online
5. **Provisions** a Gmail migration endpoint
6. **Maps** users (bulk paste, domain-swap, auto-map, or CSV)
7. **Launches** the migration batch
8. **Monitors** live sync progress with per-user stats

---

## 📦 What's Included

| File | Description |
|---|---|
| `scripts/NoTouchGoogleToM365Migration.ps1` | **Full no-touch script** — auto-deploys Google Cloud SDK, creates GCP project + service account, provisions endpoint, launches batch, monitors progress |
| `scripts/AutoGoogleToM365Migration.ps1` | **Interactive script** — self-healing module installer, 4 user-mapping modes, auto key discovery, live monitoring |
| `docs/SETUP-GUIDE.md` | Step-by-step setup guide with prerequisites |
| `docs/TROUBLESHOOTING.md` | Common errors and fixes |
| `examples/migration-users.csv` | Sample CSV format for user mapping |
| `config/.env.example` | Environment variable template |

---

## 📋 Prerequisites

### Microsoft 365
- **Exchange Online** admin access (Global Admin or Exchange Administrator role)
- **PowerShell 5.1+** on Windows
- **Internet access** to `login.microsoftonline.com` and `outlook.office365.com`

### Google Workspace
- **Google Workspace Super Admin** account
- **Google Cloud** account (free tier works)
- Ability to authorize **Domain-Wide Delegation** in Google Admin Console

---

## ⚡ Quick Start

### Option A: No-Touch Script (Recommended)

```powershell
# Download and run the no-touch script
& "$HOME\Downloads\NoTouchGoogleToM365Migration.ps1"
```

This script will:
1. Auto-install Google Cloud SDK
2. Prompt you to sign in to Google Cloud (browser)
3. Create a GCP project + service account
4. Generate the RSA key JSON
5. Show you the **Client ID** to authorize in Google Admin
6. Connect to Exchange Online
7. Create the migration endpoint + batch
8. Monitor progress live

### Option B: Interactive Script

```powershell
# Download and run the interactive script
& "$HOME\Downloads\AutoGoogleToM365Migration.ps1"
```

This script gives you **4 user-mapping modes**:
- `[1]` Bulk paste from Excel/Notepad
- `[2]` Domain-swap auto-mapping
- `[3]` Auto-migrate all M365 mailboxes
- `[4]` Import CSV/text file

---

## 🔧 Manual Setup (if you prefer step-by-step)

### Step 1: Install Google Cloud SDK

```powershell
# Download and install
$zipUrl = "https://dl.google.com/dl/cloudsdk/channels/rapid/downloads/google-cloud-cli-windows-x86_64.zip"
$zipPath = "$env:TEMP\google-cloud-sdk.zip"
Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
Expand-Archive -Path $zipPath -DestinationPath "$HOME" -Force
& "$HOME\google-cloud-sdk\install.bat" --quiet --path-update=true
```

### Step 2: Authenticate & Create Project

```powershell
gcloud auth login --update-adc
gcloud projects create m365-migration-XXXX --name="M365 Migration"
gcloud config set project m365-migration-XXXX
```

### Step 3: Enable APIs

```powershell
gcloud services enable gmail.googleapis.com calendar-json.googleapis.com people.googleapis.com admin.googleapis.com
```

### Step 4: Create Service Account & Key

```powershell
gcloud iam service-accounts create m365-migration-sa --display-name="M365 Migration SA"
gcloud iam service-accounts keys create "$HOME\Downloads\gsuite-migration-key.json" --iam-account=m365-migration-sa@PROJECT.iam.gserviceaccount.com
```

### Step 5: Authorize Domain-Wide Delegation

In **Google Admin Console** (admin.google.com):
1. Go to **Security → API controls → Domain-wide delegation**
2. Click **Add new**
3. Enter the **Client ID** from your service account JSON
4. Add these OAuth scopes:
   ```
   https://mail.google.com/
   https://www.googleapis.com/auth/calendar
   https://www.googleapis.com/auth/contacts
   https://www.googleapis.com/auth/admin.directory.user.readonly
   ```

### Step 6: Connect to Exchange Online

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Connect-ExchangeOnline
```

### Step 7: Create Migration Endpoint

```powershell
$jsonBytes = [System.IO.File]::ReadAllBytes("$HOME\Downloads\gsuite-migration-key.json")
New-MigrationEndpoint -Gmail -ServiceAccountKeyFileData $jsonBytes -EmailAddress admin@yourdomain.com -Name "GoogleWorkspaceEndpoint"
```

### Step 8: Create Migration CSV

Create a CSV with columns `EmailAddress,Username`:

```csv
EmailAddress,Username
user1@m365domain.com,user1@google.com
user2@m365domain.com,user2@google.com
```

### Step 9: Launch Migration Batch

```powershell
$csvBytes = [System.IO.File]::ReadAllBytes("$HOME\Downloads\migration-users.csv")
New-MigrationBatch -Name "Google_To_M365_Batch" `
  -SourceEndpoint "GoogleWorkspaceEndpoint" `
  -CSVData $csvBytes `
  -TargetDeliveryDomain "yourdomain.onmicrosoft.com" `
  -AutoStart
```

### Step 10: Monitor Progress

```powershell
Get-MigrationBatch -Identity "Google_To_M365_Batch"
Get-MigrationUser -BatchId "Google_To_M365_Batch" | Get-MigrationUserStatistics
```

---

## 📊 CSV Format

The migration CSV must have exactly two columns:

| Column | Description | Example |
|---|---|---|
| `EmailAddress` | **Target** M365 mailbox | `john.doe@contoso.com` |
| `Username` | **Source** Google account | `john.doe@gsuite.com` |

If source and target are the same email, just repeat it in both columns.

---

## 🔐 Security Considerations

- **Service account key** is a sensitive credential — store it securely, never commit to git
- **Domain-wide delegation** grants broad access — restrict to the minimum scopes needed
- **Exchange Online** connection uses your admin credentials — use a dedicated migration admin account
- The scripts set `Set-ExecutionPolicy Bypass -Scope Process` — this only affects the current PowerShell session

---

## 🛠️ Troubleshooting

See [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) for common issues.

**Quick fixes:**
- **"Failed to create Migration Endpoint"** → Verify the Client ID is authorized in Google Admin with correct scopes
- **"Connection attempt failed"** → Check your M365 admin credentials and network
- **Migration stuck at "Syncing"** → Check `Get-MigrationUserStatistics` for per-user errors
- **Gmail API not enabled** → Run `gcloud services enable gmail.googleapis.com`

---

## 📄 License

MIT License — see [LICENSE](LICENSE)

## 👤 Author

**Arwaz Khan** — [arwazitbp2003@gmail.com](mailto:arwazitbp2003@gmail.com)

---

*"Automate the boring stuff. Document everything. Ship fast."*
