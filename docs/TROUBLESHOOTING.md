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
- **Service account key expired** → Generate a new key with the bootstrap script: `.\Run-Bootstrap.ps1 -Mode CreateServiceAccountKey -ProjectId <project> -ApproveKeyCreation`
- **Wrong admin email** → Use the Google Workspace **Super Admin** email, not a regular user

---

## 2. "Connection attempt failed" (Exchange Online)

**Error:**
```
Connection attempt 1 failed: The user name or password is incorrect
```

**Causes & Fixes:**
- **Wrong credentials** → Verify your M365 admin credentials
- **MFA not completed** → Complete the multi-factor authentication prompt
- **Module not installed** → Run with `-InstallExchangeModule` or install manually

---

## 3. "gcloud was not found"

**Error:**
```
[FAIL] Google Cloud CLI - gcloud was not found.
```

**Fix:** Run the bootstrap script with `-Mode InstallSdk` to open the official installer, or install the Google Cloud SDK manually.

---

## 4. "ProjectId is required" / "ProjectId has an invalid format"

**Cause:** The bootstrap script requires an **existing** GCP project ID. It never creates one.

**Fix:** Provide a valid project ID:
```powershell
.\Run-Bootstrap.ps1 -Mode Inspect -ProjectId my-gcp-project
```

---

## 5. "Key creation requires -ApproveKeyCreation"

**Cause:** The bootstrap script never creates keys without explicit approval.

**Fix:** Add the approval flag only when you're ready:
```powershell
.\Run-Bootstrap.ps1 -Mode CreateServiceAccountKey -ProjectId <project> -ApproveKeyCreation
```

---

## 6. "Google prerequisites were not explicitly acknowledged"

**Cause:** The migration controller requires you to confirm that routing, domain-wide delegation, API access, and MailUser provisioning are complete.

**Fix:** Add `-ApproveGooglePrerequisites` only after verifying all prerequisites.

---

## 7. "TargetDeliveryDomain is not an accepted domain"

**Cause:** The routing domain isn't verified in Exchange Online, or you used the tenant's `onmicrosoft.com` domain (which is blocked).

**Fix:** Use a verified Google Workspace routing subdomain (e.g., `o365.contoso.com`), not `tenant.onmicrosoft.com`.

---

## 8. "Recipient is not MailUser"

**Cause:** Migration targets must be provisioned as **MailUsers**, not mailboxes.

**Fix:**
```powershell
New-MailUser -Name "John Doe" -ExternalEmailAddress "john.doe@gsuite.com" -PrimarySmtpAddress "john.doe@contoso.com"
```

---

## 9. "Batch status is not Synced"

**Cause:** You tried to complete a batch that hasn't reached **Synced** status.

**Fix:** Monitor until the batch is `Synced`, review statistics, then complete.

---

## 10. "Blocking findings exist"

**Cause:** The script found Critical or High findings and halted to prevent unsafe action.

**Fix:** Review the report (`MigrationPreflightFindings.csv` / `GoogleCloudBootstrapFindings.csv`), resolve the blocking findings, and rerun.

---

## 11. Launcher says "Script not found"

**Cause:** The launcher couldn't find the target script in the `scripts/` folder.

**Fix:** Ensure the launcher and `scripts/` folder are in the same directory (i.e., don't move the launcher out of the repo).

---

*"Automate the boring stuff. Document everything. Ship fast."*
