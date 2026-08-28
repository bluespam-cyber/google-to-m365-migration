# Troubleshooting Guide

Common issues and fixes for the Google Workspace → Microsoft 365 migration.

---

## 1. "Failed to create Migration Endpoint"

**Error:**
```
Failed to create Migration Endpoint: The remote server returned an error: (401) Unauthorized
```

**Causes & Fixes:**
- **Domain-Wide Delegation not authorized** → Go to admin.google.com → Security → API controls → Domain-wide delegation → verify the Client ID is added with correct scopes
- **Wrong Client ID** → The Client ID is the `client_id` field in your JSON key (long numeric string), NOT `client_email`
- **Service account key expired** → Generate a new key: `gcloud iam service-accounts keys create new-key.json --iam-account=SA@PROJECT.iam.gserviceaccount.com`
- **Wrong admin email** → Use the Google Workspace **Super Admin** email, not a regular user

---

## 2. "Connection attempt failed" (Exchange Online)

**Error:**
```
Connection attempt 1 failed: The user name or password is incorrect
```

**Causes & Fixes:**
- **Wrong credentials** → Verify your M365 admin credentials
- **MFA required** → Use a modern auth-capable account; the script supports MFA prompts
- **Module not installed** → Run `Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force`
- **Network blocked** → Ensure you can reach `login.microsoftonline.com`

---

## 3. Migration stuck at "Syncing"

**Symptom:** Batch status stays at `Syncing` for a long time.

**Diagnosis:**
```powershell
Get-MigrationUser -BatchId "Google_To_M365_AutoBatch" | Get-MigrationUserStatistics | Format-Table Identity, Status, ErrorSummary
```

**Common causes:**
- **Large mailboxes** → Large mailboxes take time; check `ItemsTransferred` is increasing
- **Per-user errors** → Check `ErrorSummary` column for specific errors
- **Rate limiting** → Google API rate limits; the migration will retry automatically

---

## 4. "Gmail API has not been used in project"

**Error:**
```
Error 403: Access Not Configured. Gmail API has not been used in project X before or it is disabled.
```

**Fix:**
```powershell
gcloud services enable gmail.googleapis.com
```

---

## 5. "Failed to parse Google JSON key file"

**Causes & Fixes:**
- **Wrong file selected** → Ensure you're using the service account key JSON (has `type: "service_account"`)
- **Corrupted file** → Regenerate: `gcloud iam service-accounts keys create new-key.json --iam-account=SA@PROJECT.iam.gserviceaccount.com`
- **File permissions** → Ensure the file is readable

---

## 6. "No users were provided for migration"

**Cause:** The user mapping step produced zero users.

**Fixes:**
- **Bulk paste mode** → Ensure you pressed Enter on an empty line to finish
- **CSV import** → Verify the CSV has the correct `EmailAddress,Username` headers
- **Domain-swap** → Verify the source Google domain is correct

---

## 7. Migration batch shows "Failed" status

**Diagnosis:**
```powershell
Get-MigrationUser -BatchId "Google_To_M365_AutoBatch" -Status Failed | Format-Table Identity, Status
Get-MigrationUserStatistics -Identity <failed-user> | Select ErrorSummary
```

**Common causes:**
- **User doesn't exist in Google** → Verify the source email is correct
- **User doesn't exist in M365** → Verify the target mailbox exists
- **Domain-wide delegation missing** → Re-verify the Client ID authorization
- **Mailbox locked** → Ensure the target mailbox isn't on litigation hold or locked

---

## 8. "TargetDeliveryDomain" errors

**Symptom:** Batch fails with delivery domain errors.

**Fix:** Specify the domain explicitly:
```powershell
& "$HOME\Downloads\NoTouchGoogleToM365Migration.ps1" -TargetDeliveryDomain "contoso.onmicrosoft.com"
```

---

## 9. Google Cloud SDK not found

**Fix:**
```powershell
# Add to PATH manually
$env:Path = "$HOME\google-cloud-sdk\bin;$env:Path"
# Or reinstall
& "$HOME\google-cloud-sdk\install.bat" --quiet --path-update=true
```

---

## 10. Migration is very slow

**Tips:**
- **Increase parallelism** — Use `Set-MigrationBatch -Identity <batch> -MaxConcurrentMigrations 20`
- **Check network** — Ensure stable internet connection
- **Monitor** — Use `Get-MigrationUserStatistics` to see per-user progress
- **Large items** — Large attachments take longer; this is expected

---

## Still stuck?

1. Check the **Migration User Statistics** for specific error messages
2. Verify **Domain-Wide Delegation** is correctly configured
3. Confirm the **service account key** is valid and not expired
4. Check **Google Cloud console** → APIs & Services → Dashboard → verify APIs are enabled
5. Open an issue on this repository with the full error output
