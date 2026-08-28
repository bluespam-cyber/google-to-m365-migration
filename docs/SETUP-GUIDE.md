# Setup Guide — Google Workspace → Microsoft 365 Migration

Complete step-by-step guide to prepare your environment before running the migration scripts.

---

## 1. Prerequisites Checklist

### Microsoft 365 Requirements
- [ ] **Global Admin** or **Exchange Administrator** role in Microsoft 365
- [ ] PowerShell 5.1+ (Windows PowerShell, not PowerShell 7)
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
Get-Mailbox -ResultSize 5 | Select DisplayName, PrimarySmtpAddress
```

### 2.3 Note Your Target Delivery Domain

```powershell
Get-AcceptedDomain | Format-Table DomainName, Default
```

The **Target Delivery Domain** is typically your `.onmicrosoft.com` domain (e.g., `contoso.onmicrosoft.com`).

---

## 3. Google Cloud Preparation

### 3.1 Install Google Cloud SDK

**Option A: Automatic (via No-Touch script)**
The `NoTouchGoogleToM365Migration.ps1` script auto-installs the SDK.

**Option B: Manual**
```powershell
$zipUrl = "https://dl.google.com/dl/cloudsdk/channels/rapid/downloads/google-cloud-cli-windows-x86_64.zip"
$zipPath = "$env:TEMP\google-cloud-sdk.zip"
Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
Expand-Archive -Path $zipPath -DestinationPath "$HOME" -Force
& "$HOME\google-cloud-sdk\install.bat" --quiet --path-update=true
```

### 3.2 Authenticate to Google Cloud

```powershell
gcloud auth login --update-adc
```
This opens a browser — sign in with your Google Workspace Super Admin account.

### 3.3 Create a GCP Project

```powershell
gcloud projects create m365-migration-1234 --name="M365 Migration"
gcloud config set project m365-migration-1234
```

### 3.4 Enable Required APIs

```powershell
gcloud services enable gmail.googleapis.com calendar-json.googleapis.com people.googleapis.com admin.googleapis.com
```

### 3.5 Create Service Account

```powershell
gcloud iam service-accounts create m365-migration-sa --display-name="M365 Migration SA"
```

### 3.6 Generate RSA Key

```powershell
gcloud iam service-accounts keys create "$HOME\Downloads\gsuite-migration-key.json" --iam-account=m365-migration-sa@m365-migration-1234.iam.gserviceaccount.com
```

---

## 4. Google Admin Console — Domain-Wide Delegation

This is the **most critical step**. Without it, the migration will fail.

1. Go to **admin.google.com**
2. Navigate to **Security → API controls → Domain-wide delegation**
3. Click **Add new**
4. Enter the **Client ID** from your service account JSON file
5. Add these OAuth scopes (comma-separated):
   ```
   https://mail.google.com/,https://www.googleapis.com/auth/calendar,https://www.googleapis.com/auth/contacts,https://www.googleapis.com/auth/admin.directory.user.readonly
   ```
6. Click **Authorize**

> ⚠️ **Important:** The Client ID is in the `client_id` field of your JSON key file. It's a long numeric string, NOT the `client_email`.

---

## 5. Prepare Migration CSV

Create a CSV file with exactly two columns:

```csv
EmailAddress,Username
john.doe@contoso.com,john.doe@gsuite.com
jane.smith@contoso.com,jane.smith@gsuite.com
```

| Column | Description |
|---|---|
| `EmailAddress` | **Target** M365 mailbox (where mail goes) |
| `Username` | **Source** Google account (where mail comes from) |

Save it as `migration-users.csv` in your Downloads folder.

---

## 6. Run the Migration

### No-Touch Script
```powershell
& "$HOME\Downloads\NoTouchGoogleToM365Migration.ps1"
```

### Interactive Script
```powershell
& "$HOME\Downloads\AutoGoogleToM365Migration.ps1"
```

---

## 7. Post-Migration Steps

1. **Verify mail flow** — Send test emails to migrated users
2. **Check calendar** — Confirm calendar items transferred
3. **Check contacts** — Confirm contacts transferred
4. **Update DNS** — Point MX records to Microsoft 365 (after cutover)
5. **Complete the batch** — Once verified:
   ```powershell
   Complete-MigrationBatch -Identity "Google_To_M365_AutoBatch"
   ```
6. **Remove the endpoint** — After migration is complete:
   ```powershell
   Remove-MigrationEndpoint -Identity "GoogleWorkspaceEndpoint"
   ```

---

## 8. Common Parameters

| Parameter | Description | Default |
|---|---|---|
| `-GoogleAdminEmail` | Google Workspace Super Admin email | Prompted |
| `-TargetDeliveryDomain` | M365 delivery domain | Auto-detected |
| `-BatchName` | Migration batch name | `Google_To_M365_AutoBatch` |
| `-EndpointName` | Migration endpoint name | `GoogleWorkspaceEndpoint` |
| `-CsvPath` | Path to migration CSV | Auto-detected |
| `-MonitorIntervalSeconds` | Progress refresh interval | `45` |
