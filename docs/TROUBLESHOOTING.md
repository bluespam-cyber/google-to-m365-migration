# Troubleshooting Guide

Common issues and fixes for the Google Workspace → Microsoft 365 migration controller.

---

## 1. "Failed to create Migration Endpoint" (401 Unauthorized)

**Error:**
```
Failed to create Migration Endpoint: The remote server returned an error: (401) Unauthorized
```

**Causes & Fixes:**
- **Domain-Wide Delegation not authorized** → admin.google.com → Security → Access and data control → API controls → Manage Domain Wide Delegation → verify the numeric Client ID is added with the exact five scopes
- **Wrong Client ID** → The Client ID is the **numeric Unique ID** (also the `client_id` field in the JSON key), NOT `client_email`
- **Service account key expired** → Generate a new JSON key in the Google Cloud console (service account → Keys → Add key)
- **Wrong admin email** → Use the Google Workspace **Super Admin** email, not a regular user

---

## 2. "invalid_grant" during the Google delegation test

**Error:**
```
Google delegation test failed: invalid_grant
```

**Causes & Fixes:**
- **Clock skew** → The JWT is time-sensitive. Verify the machine clock is accurate (NTP sync).
- **Key file corrupted or wrong** → Confirm `-KeyPath` points at the JSON key you downloaded (not a renamed copy).
- **Service account deleted** → Recreate the service account and key in the Google Cloud console.

---

## 3. "unauthorized_client" / "access_denied" — delegation not ready

**Error:**
```
Google delegation test failed: unauthorized_client
```

**Causes & Fixes:**
- **Delegation not yet propagated** → Domain-wide delegation can take minutes, occasionally up to 24 hours. The script automatically **waits and retries for up to 30 minutes** (backoff 20s → 60s) before giving up.
- **Wrong Client ID used in Admin Console** → Re-check the numeric Unique ID (step 5 of the guide).
- **Scopes mismatch** → The scopes in Admin Console must match this exact line, no spaces:
  ```
  https://mail.google.com/,https://www.googleapis.com/auth/calendar,https://www.google.com/m8/feeds/,https://www.googleapis.com/auth/gmail.settings.sharing,https://www.googleapis.com/auth/contacts
  ```

---

## 4. "Connection attempt failed" (Exchange Online)

**Error:**
```
Connection attempt 1 failed: The user name or password is incorrect
```

**Causes & Fixes:**
- **Wrong credentials** → Verify your M365 admin credentials
- **MFA not completed** → Complete the multi-factor authentication prompt
- **Module not installed** → The script auto-installs ExchangeOnlineManagement (TLS 1.2, NuGet, PSGallery restore, three strategies). If auto-install fails, install manually:
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
  ```

---

## 5. "A window handle must be configured" (headless run)

**Error:**
```
Error Acquiring Token: A window handle must be configured for the UI
```

**Cause:** Interactive M365 sign-in (MSAL) requires a desktop session. This happens when running under SYSTEM/RMM or a non-interactive service.

**Fix:** Run the interactive wizard (`.\Run-Migration.ps1`) once from a logged-in desktop session to sign in, or use a service principal / certificate auth for fully unattended runs.

---

## 6. "TargetDeliveryDomain is not an accepted domain"

**Cause:** The routing domain isn't verified/Active in Exchange Online, or you used the tenant's `onmicrosoft.com` domain (which is blocked).

**Fix:**
- Add and verify the `o365.<yourdomain>` subdomain in Microsoft 365 (must show **Active**)
- In interactive mode, the script shows a **picker of accepted domains** and applies your choice immediately
- Never use `tenant.onmicrosoft.com`

---

## 7. "Recipient is not MailUser"

**Cause:** Migration targets must be provisioned as **MailUsers**, not mailboxes.

**Fix:** The script can **auto-provision missing MailUsers** from your CSV. To provision manually:
```powershell
New-MailUser -Name "John Doe" -ExternalEmailAddress "john.doe@gsuite.contoso.com" -PrimarySmtpAddress "john.doe@contoso.com"
```
Accounts that would need a password are skipped and listed in the `noAccount` report rather than guessed.

---

## 8. "Batch status is not Synced"

**Cause:** You tried to complete a batch that hasn't reached **Synced** status.

**Fix:** Monitor until the batch is `Synced`, review statistics, then complete:
```powershell
.\Run-Migration.ps1 -Mode Monitor -CollectUserStatistics
.\Run-Migration.ps1 -Mode Complete -ApproveCutover
```

---

## 9. "Blocking findings exist"

**Cause:** The script found Critical or High findings and halted to prevent unsafe action.

**Fix:** Review the findings report (written to the output folder), resolve the blocking findings, and rerun. Run `-Mode Preflight` to re-validate after fixing.

---

## 10. "Key file not found" / auto-discovery failed

**Cause:** No `-KeyPath` was supplied and no service-account JSON key was found in Downloads, Desktop, or the current folder.

**Fix:** Pass the path explicitly:
```powershell
.\Run-Migration.ps1 -Mode Preflight -KeyPath C:\keys\gws.json ...
```

---

## 11. CSV not found / "CSV is empty or truncated"

**Cause:** No `-CsvPath` was supplied and no `migration-users.csv` was found in the current folder, `examples/`, `config/`, Downloads, or Documents. Or the file is empty/truncated (the script rejects files under 3 bytes and files with a UTF-8 BOM).

**Fix:** Pass the path explicitly and save the CSV as plain UTF-8 (no BOM):
```powershell
.\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv ...
```

---

## 12. Launcher says "Script not found"

**Cause:** The launcher couldn't find the target script in the `scripts/` folder.

**Fix:** Ensure the launcher and `scripts/` folder are in the same directory (i.e., don't move the launcher out of the repo).

---

*"Automate the boring stuff. Document everything. Ship fast."*