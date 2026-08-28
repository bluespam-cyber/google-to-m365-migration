[CmdletBinding()]
param (
    # Optional: Direct path to Google Service Account JSON key (Auto-discovered if omitted)
    [Parameter(Mandatory = $false)]
    [string]$JsonPath,

    # Google Workspace Super Admin Email (e.g. admin@yourdomain.com)
    [Parameter(Mandatory = $false)]
    [string]$GoogleAdminEmail,

    # Target Delivery Domain (Auto-detected from M365 if omitted)
    [Parameter(Mandatory = $false)]
    [string]$TargetDeliveryDomain,

    # Migration Batch Name
    [Parameter(Mandatory = $false)]
    [string]$BatchName = "Google_To_M365_MigrationBatch",

    # Migration Endpoint Name
    [Parameter(Mandatory = $false)]
    [string]$EndpointName = "GoogleWorkspaceEndpoint",

    # Monitoring interval in seconds
    [Parameter(Mandatory = $false)]
    [int]$MonitorIntervalSeconds = 45
)

# Enforce TLS 1.2 and bypass process execution policy
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

function Write-StepHeader ($Title) {
    Write-Host "`n==============================================================" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "==============================================================" -ForegroundColor Cyan
}

# --- STEP 1: Environment & Module Auto-Resolver ---
Write-StepHeader "1. Environment Setup & Module Self-Healing"

# Ensure NuGet Provider
if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
    Write-Host "[*] Installing NuGet package provider..." -ForegroundColor Yellow
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "NuGet standard install notice: $_"
    }
}

# Set PSGallery to Trusted to prevent hanging prompts
try {
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
} catch {}

# Check and install ExchangeOnlineManagement with fallback
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "[*] Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "[+] Module installed successfully." -ForegroundColor Green
    } catch {
        Write-Warning "Standard install failed ($($_)). Attempting repair install..."
        Install-Module -Name PowerShellGet -Scope CurrentUser -Force -AllowClobber -ErrorAction SilentlyContinue
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "[+] Module repair-installation successful." -ForegroundColor Green
    }
} else {
    Write-Host "[+] ExchangeOnlineManagement module is available." -ForegroundColor Green
}

Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

# --- STEP 2: Exchange Online Connection with Retry ---
Write-StepHeader "2. Establishing Exchange Online Connection"

$connected = $false
if (Get-ConnectionInformation -ErrorAction SilentlyContinue) {
    Write-Host "[+] Active Exchange Online session found." -ForegroundColor Green
    $connected = $true
} else {
    $maxAttempts = 3
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        Write-Host "[*] Connecting to Exchange Online (Attempt $attempt of $maxAttempts)..." -ForegroundColor Yellow
        try {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $connected = $true
            Write-Host "[+] Connected to Exchange Online." -ForegroundColor Green
            break
        } catch {
            Write-Warning "Connection attempt $attempt failed: $_"
            if ($attempt -lt $maxAttempts) {
                Write-Host "Retrying in 5 seconds..." -ForegroundColor Gray
                Start-Sleep -Seconds 5
            }
        }
    }
}

if (-not $connected) {
    Write-Error "Could not connect to Exchange Online after multiple attempts. Please verify credentials."
    return
}

# --- STEP 3: Bulk User Mapping Engine ---
Write-StepHeader "3. Bulk Source & Target User Mapping"

$csvFilePath = "$HOME\Downloads\GoogleMigrationUsers.csv"
$userMappings = @()

Write-Host "Choose how to provide migration users:" -ForegroundColor Yellow
Write-Host " [1] Bulk Paste from Excel / Notepad (Paste rows of source and target emails)"
Write-Host " [2] Domain-Swap Auto-Mapping (Auto-map all M365 tenant users to a Google domain)"
Write-Host " [3] Auto-Migrate All M365 Mailboxes (Same email address in Google & M365)"
Write-Host " [4] Import CSV / Text file (Auto-detects local test CSV or custom file)"
$choice = Read-Host "`nSelect an option (1, 2, 3, or 4) [Default: 1]"

switch ($choice) {
    '2' {
        Write-Host "`n--- Domain-Swap Auto-Mapping ---" -ForegroundColor Cyan
        $sourceDomain = (Read-Host "Enter Source Google Domain (e.g. gsuitedomain.com)").Trim('@').Trim()

        Write-Host "[*] Querying user mailboxes from Microsoft 365..." -ForegroundColor Yellow
        $m365Users = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Where-Object { $_.UserPrincipalName -notmatch 'DiscoverySearchMailbox' }

        foreach ($u in $m365Users) {
            $alias = $u.Alias
            $userMappings += [PSCustomObject]@{
                "Source (Google)" = "$alias@$sourceDomain"
                "Target (M365)"   = $u.PrimarySmtpAddress
            }
        }
        Write-Host "[+] Auto-mapped $($userMappings.Count) user(s) using domain '@$sourceDomain'." -ForegroundColor Green
    }
    '3' {
        Write-Host "`n--- Auto-Migrate All M365 Mailboxes ---" -ForegroundColor Cyan
        Write-Host "[*] Querying all active tenant mailboxes..." -ForegroundColor Yellow
        $m365Users = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Where-Object { $_.UserPrincipalName -notmatch 'DiscoverySearchMailbox' }

        foreach ($u in $m365Users) {
            $userMappings += [PSCustomObject]@{
                "Source (Google)" = $u.PrimarySmtpAddress
                "Target (M365)"   = $u.PrimarySmtpAddress
            }
        }
        Write-Host "[+] Loaded $($userMappings.Count) user mailbox(es)." -ForegroundColor Green
    }
    '4' {
        Write-Host "`n--- Import CSV / Text File ---" -ForegroundColor Cyan
        $defaultCsv = "$HOME\Downloads\Google_to_M365_Migration_Users_Dummy.csv"
        $promptMsg = if (Test-Path -Path $defaultCsv) { "Enter CSV file path [Press Enter for detected test file: $defaultCsv]" } else { "Enter CSV file path or drag and drop file here" }
        $filePath = (Read-Host $promptMsg).Trim('"').Trim('''')
        if ([string]::IsNullOrWhiteSpace($filePath) -and (Test-Path -Path $defaultCsv)) {
            $filePath = $defaultCsv
        }

        if (Test-Path -Path $filePath) {
            $rawContent = Import-Csv -Path $filePath -ErrorAction SilentlyContinue
            if ($rawContent -and $rawContent[0].PSObject.Properties['EmailAddress']) {
                $userMappings = foreach ($row in $rawContent) {
                    [PSCustomObject]@{
                        "Source (Google)" = if ($row.Username) { $row.Username } else { $row.EmailAddress }
                        "Target (M365)"   = $row.EmailAddress
                    }
                }
            } else {
                $lines = Get-Content -Path $filePath
                foreach ($line in $lines) {
                    $parts = $line -split '[,;\t]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
                    if ($parts.Count -ge 2) {
                        $userMappings += [PSCustomObject]@{ "Source (Google)" = $parts[1]; "Target (M365)" = $parts[0] }
                    } elseif ($parts.Count -eq 1) {
                        $userMappings += [PSCustomObject]@{ "Source (Google)" = $parts[0]; "Target (M365)" = $parts[0] }
                    }
                }
            }
            Write-Host "[+] Loaded $($userMappings.Count) user mapping(s) from file." -ForegroundColor Green
        } else {
            Write-Warning "File not found. Switching to Bulk Paste mode."
            $choice = '1'
        }
    }
    Default {
        Write-Host "`n--- Bulk Paste Mode ---" -ForegroundColor Cyan
        Write-Host "Paste your user list below. Formats accepted:" -ForegroundColor White
        Write-Host "  - source@google.com, target@m365.com" -ForegroundColor Gray
        Write-Host "  - Two columns copied directly from Excel (Tab-separated)" -ForegroundColor Gray
        Write-Host "  - Single email per line (if source & target are identical)" -ForegroundColor Gray
        Write-Host "`nPaste lines below and press ENTER ON AN EMPTY LINE when done:" -ForegroundColor Yellow

        $pastedLines = @()
        while ($true) {
            $line = Read-Host
            if ([string]::IsNullOrWhiteSpace($line)) { break }
            $pastedLines += $line.Trim()
        }

        foreach ($entry in $pastedLines) {
            $parts = $entry -split '[,;\t]|->' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
            if ($parts.Count -ge 2) {
                $userMappings += [PSCustomObject]@{
                    "Source (Google)" = $parts[0]
                    "Target (M365)"   = $parts[1]
                }
            } elseif ($parts.Count -eq 1) {
                $userMappings += [PSCustomObject]@{
                    "Source (Google)" = $parts[0]
                    "Target (M365)"   = $parts[0]
                }
            }
        }
        Write-Host "`n[+] Processed $($userMappings.Count) user mapping(s) from pasted input." -ForegroundColor Green
    }
}

if ($userMappings.Count -eq 0) {
    Write-Error "No users were provided for migration. Process aborted."
    return
}

# Display sample of users
Write-Host "`nSample of User Mappings to be Migrated (First 5 of $($userMappings.Count)):" -ForegroundColor Cyan
$userMappings | Select-Object -First 5 | Format-Table -AutoSize

# Export to Migration CSV format
$csvLines = @("EmailAddress,Username")
foreach ($u in $userMappings) {
    $csvLines += "$($u.'Target (M365)'),$($u.'Source (Google)')"
}
$csvLines | Set-Content -Path $csvFilePath -Encoding UTF8
$csvBytes = [System.IO.File]::ReadAllBytes($csvFilePath)
Write-Host "[+] Migration CSV ready at: $csvFilePath" -ForegroundColor Green

# --- STEP 4: Google JSON Key Discovery ---
Write-StepHeader "4. Google Service Account Key Discovery"

if (-not $JsonPath -or -not (Test-Path -Path $JsonPath)) {
    Write-Host "[*] Scanning Downloads, Desktop, and Documents for Google key..." -ForegroundColor Yellow

    $searchLocations = @("$HOME\Downloads", "$HOME\Desktop", "$HOME\Documents", (Get-Location).Path)
    $validGoogleKeys = @()

    foreach ($dir in $searchLocations) {
        if (Test-Path -Path $dir) {
            $files = Get-ChildItem -Path $dir -Filter "*.json" -File -ErrorAction SilentlyContinue
            foreach ($file in $files) {
                try {
                    $parsed = Get-Content -Path $file.FullName -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue
                    if ($parsed.type -eq "service_account" -and $parsed.client_id -and $parsed.private_key) {
                        $validGoogleKeys += [PSCustomObject]@{
                            Path         = $file.FullName
                            Name         = $file.Name
                            ClientId     = $parsed.client_id
                            ClientEmail  = $parsed.client_email
                            LastModified = $file.LastWriteTime
                        }
                    }
                } catch {}
            }
        }
    }

    if ($validGoogleKeys.Count -gt 0) {
        $selectedKey = $validGoogleKeys | Sort-Object LastModified -Descending | Select-Object -First 1
        $JsonPath = $selectedKey.Path
        Write-Host "[+] Found Google Service Account key file:" -ForegroundColor Green
        Write-Host "    $($selectedKey.Path)" -ForegroundColor White
    } else {
        Write-Warning "No Google Service Account JSON file found in Downloads or Desktop."
        $JsonPath = (Read-Host "Enter the path or drag & drop your JSON file here").Trim('"').Trim('''')
    }
}

# Parse key data
try {
    $jsonContent = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json
    $clientId    = $jsonContent.client_id
    $clientEmail = $jsonContent.client_email

    Write-Host "`n[+] Loaded Credentials:" -ForegroundColor Green
    Write-Host "    - Service Account: $clientEmail" -ForegroundColor White
    Write-Host "    - Client ID:       $clientId" -ForegroundColor White
} catch {
    Write-Error "Failed to parse Google JSON key file ($JsonPath): $_"
    return
}

# Resolve Google Admin Email
if (-not $GoogleAdminEmail) {
    $GoogleAdminEmail = Read-Host "`nEnter Google Workspace Super Admin Email (e.g. admin@yourdomain.com)"
}

Write-StepHeader "Domain-Wide Delegation Verification"
Write-Host "Ensure this Client ID is authorized in Google Admin Console (admin.google.com):" -ForegroundColor Yellow
Write-Host "  - Client ID : $clientId" -ForegroundColor Green
Write-Host "  - Scopes    : https://mail.google.com/,https://www.googleapis.com/auth/calendar,https://www.googleapis.com/auth/contacts,https://www.googleapis.com/auth/admin.directory.user.readonly`n" -ForegroundColor Green

# --- STEP 5: Migration Endpoint Setup & Self-Healing ---
Write-StepHeader "5. Migration Endpoint Setup"

$existingEndpoint = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue
$jsonBytes = [System.IO.File]::ReadAllBytes((Resolve-Path $JsonPath).Path)

if (-not $existingEndpoint) {
    Write-Host "[*] Creating Migration Endpoint '$EndpointName' in Exchange Online..." -ForegroundColor Yellow
    try {
        New-MigrationEndpoint -Gmail -ServiceAccountKeyFileData $jsonBytes -EmailAddress $GoogleAdminEmail -Name $EndpointName -ErrorAction Stop
        Write-Host "[+] Migration Endpoint '$EndpointName' created." -ForegroundColor Green
    } catch {
        Write-Error "Failed to create Migration Endpoint: $_"
        return
    }
} else {
    Write-Host "[+] Using existing Migration Endpoint '$EndpointName'." -ForegroundColor Green
}

# Determine Target Delivery Domain
if (-not $TargetDeliveryDomain) {
    $domains = Get-AcceptedDomain | Where-Object { $_.DomainName -match '\.onmicrosoft\.com$' }
    $TargetDeliveryDomain = if ($domains) { $domains[0].DomainName } else { (Get-AcceptedDomain | Where-Object { $_.Default }).DomainName }
    Write-Host "[i] Target Delivery Domain: $TargetDeliveryDomain" -ForegroundColor Cyan
}

# --- STEP 6: Migration Batch Execution & Self-Healing ---
Write-StepHeader "6. Launching Migration Batch"

$existingBatch = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue

if (-not $existingBatch) {
    Write-Host "[*] Initializing Migration Batch '$BatchName' for $($userMappings.Count) user(s)..." -ForegroundColor Yellow
    try {
        New-MigrationBatch -Name $BatchName `
            -SourceEndpoint $EndpointName `
            -CSVData $csvBytes `
            -TargetDeliveryDomain $TargetDeliveryDomain `
            -AutoStart `
            -AutoComplete:$false `
            -ErrorAction Stop

        Write-Host "[+] Migration Batch '$BatchName' created and started." -ForegroundColor Green
    } catch {
        Write-Error "Failed to create Migration Batch: $_"
        return
    }
} else {
    Write-Host "[+] Migration Batch '$BatchName' already exists (Status: $($existingBatch.Status))." -ForegroundColor Yellow
    if ($existingBatch.Status -eq 'Created' -or $existingBatch.Status -eq 'Stopped') {
        Start-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
        Write-Host "[+] Migration Batch started." -ForegroundColor Green
    }
}

# --- STEP 7: Live Progress Monitoring Loop ---
Write-StepHeader "7. Live Bulk Migration Progress Tracker"
Write-Host "Monitoring synchronization for $($userMappings.Count) users. Press Ctrl+C at any time to exit monitoring (cloud migration continues in the background)." -ForegroundColor Gray

$isComplete = $false

while (-not $isComplete) {
    $batch = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
    if (-not $batch) { break }

    Write-Host "`n--- [$(Get-Date -Format 'HH:mm:ss')] Batch Status: $($batch.Status) | Synced: $($batch.SyncedCount)/$($batch.TotalCount) | InProgress: $($batch.InProgressCount) | Failed: $($batch.FailedCount) ---" -ForegroundColor Cyan

    try {
        $users = Get-MigrationUser -BatchId $BatchName -ErrorAction SilentlyContinue
        if ($users) {
            $userStats = foreach ($u in $users) {
                $stat = Get-MigrationUserStatistics -Identity $u.Identity -ErrorAction SilentlyContinue
                [PSCustomObject]@{
                    "Target Mailbox (M365)" = $u.Identity
                    "Sync Status"           = $u.Status
                    "Items Transferred"     = if ($stat) { $stat.ItemsTransferred } else { 0 }
                    "Data Transferred"      = if ($stat) { $stat.BytesTransferred.ToString() } else { "0 B" }
                    "Percent Complete"      = if ($stat) { "$($stat.PercentageComplete) %" } else { "0 %" }
                    "Error / Notice"        = if ($stat -and $stat.ErrorSummary) { $stat.ErrorSummary } else { "None" }
                }
            }
            $userStats | Format-Table -AutoSize
        }
    } catch {}

    if ($batch.Status -eq 'Completed' -or $batch.Status -eq 'Synced') {
        $isComplete = $true
        Write-Host "`n[√] Bulk migration batch synchronization is complete!" -ForegroundColor Green
        break
    }

    Start-Sleep -Seconds $MonitorIntervalSeconds
}
