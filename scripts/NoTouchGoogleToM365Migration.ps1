[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$GoogleAdminEmail,

    [Parameter(Mandatory = $false)]
    [string]$TargetDeliveryDomain,

    [Parameter(Mandatory = $false)]
    [string]$BatchName = "Google_To_M365_AutoBatch",

    [Parameter(Mandatory = $false)]
    [string]$EndpointName = "GoogleWorkspaceEndpoint",

    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [int]$MonitorIntervalSeconds = 45
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

function Write-StepHeader ($Title) {
    Write-Host "`n==============================================================" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "==============================================================" -ForegroundColor Cyan
}

# --- STEP 1: Fast Silent Google Cloud SDK Setup ---
Write-StepHeader "1. Google Cloud SDK Deployment"

function Sync-GcloudPath {
    $paths = @(
        "$HOME\google-cloud-sdk\bin",
        "$env:LOCALAPPDATA\Google\Cloud SDK\google-cloud-sdk\bin",
        "$env:ProgramFiles\Google\Cloud SDK\google-cloud-sdk\bin"
    )
    foreach ($p in $paths) {
        if ((Test-Path -Path $p) -and ($env:Path -notlike "*$p*")) {
            $env:Path = "$p;$env:Path"
        }
    }
}

Sync-GcloudPath

if (-not (Get-Command gcloud -ErrorAction SilentlyContinue)) {
    Write-Host "[*] Setting up Google Cloud CLI bundle in background..." -ForegroundColor Yellow
    $zipUrl = "https://dl.google.com/dl/cloudsdk/channels/rapid/downloads/google-cloud-cli-windows-x86_64.zip"
    $zipPath = "$env:TEMP\google-cloud-sdk.zip"

    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    Expand-Archive -Path $zipPath -DestinationPath "$HOME" -Force
    & "$HOME\google-cloud-sdk\install.bat" --quiet --path-update=true --command-completion=false 2>$null
    $env:Path = "$HOME\google-cloud-sdk\bin;$env:Path"
    Write-Host "[+] Google Cloud CLI ready." -ForegroundColor Green
} else {
    Write-Host "[+] Google Cloud CLI detected." -ForegroundColor Green
}

# --- STEP 2: Google Cloud Services & Key Auto-Generation ---
Write-StepHeader "2. Google Cloud API & Credentials Automation"

$activeAccount = (& gcloud auth list --filter=status:ACTIVE --format="value(account)" 2>$null)

if (-not $activeAccount) {
    Write-Host "[*] Initiating Google Cloud browser authentication..." -ForegroundColor Yellow
    & gcloud auth login --update-adc
    $activeAccount = (& gcloud auth list --filter=status:ACTIVE --format="value(account)" 2>$null)
    Write-Host "[+] Authenticated as: $activeAccount" -ForegroundColor Green
} else {
    Write-Host "[+] Active Google Account: $activeAccount" -ForegroundColor Green
}

# Auto-assign or create GCP Project
$gcpProject = (& gcloud config get-value project 2>$null)
if (-not $gcpProject -or $gcpProject -eq '(unset)') {
    $randomSuffix = Get-Random -Minimum 1000 -Maximum 9999
    $gcpProject = "m365-migration-$randomSuffix"
    Write-Host "[*] Creating Google Cloud Project '$gcpProject'..." -ForegroundColor Yellow
    & gcloud projects create $gcpProject --name="M365 Migration" --quiet 2>$null
    & gcloud config set project $gcpProject --quiet
}
Write-Host "[+] Active Project: $gcpProject" -ForegroundColor Green

# Enable APIs in batch
Write-Host "[*] Activating Gmail, Calendar, People, and Admin SDK APIs..." -ForegroundColor Yellow
& gcloud services enable gmail.googleapis.com calendar-json.googleapis.com people.googleapis.com admin.googleapis.com --quiet
Write-Host "[+] APIs active." -ForegroundColor Green

# Create Service Account & RSA Key
$saName = "m365-migration-sa"
$saEmail = "$saName@$gcpProject.iam.gserviceaccount.com"
$JsonPath = "$HOME\Downloads\gsuite-migration-key.json"

Write-Host "[*] Creating Service Account ($saEmail)..." -ForegroundColor Yellow
& gcloud iam service-accounts create $saName --display-name="M365 Migration Service Account" --quiet 2>$null

Write-Host "[*] Generating RSA Private Key JSON..." -ForegroundColor Yellow
& gcloud iam service-accounts keys create $JsonPath --iam-account=$saEmail --quiet

$jsonContent = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json
$clientId = $jsonContent.client_id

Write-StepHeader "Domain-Wide Delegation Authorization"
Write-Host "Authorize this Client ID in Google Admin (admin.google.com):" -ForegroundColor Yellow
Write-Host "  - Client ID : $clientId" -ForegroundColor Green
Write-Host "  - Scopes    : https://mail.google.com/,https://www.googleapis.com/auth/calendar,https://www.googleapis.com/auth/contacts,https://www.googleapis.com/auth/admin.directory.user.readonly`n" -ForegroundColor Green

# --- STEP 3: Exchange Online Connection ---
Write-StepHeader "3. Exchange Online Setup & Connection"

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "[*] Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "[*] Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "[+] Connected to Exchange Online." -ForegroundColor Green
} else {
    Write-Host "[+] Active Exchange Online session found." -ForegroundColor Green
}

# Resolve Google Admin Email
if (-not $GoogleAdminEmail) {
    $domain = (Get-AcceptedDomain | Where-Object { -not $_.DomainName.EndsWith(".onmicrosoft.com") } | Select-Object -First 1).DomainName
    if (-not $domain) { $domain = (Get-AcceptedDomain | Select-Object -First 1).DomainName }
    $GoogleAdminEmail = Read-Host "Enter Google Workspace Super Admin Email (e.g. admin@$domain)"
}

# --- STEP 4: Migration Endpoint Setup ---
Write-StepHeader "4. Migration Endpoint Provisioning"

$existingEndpoint = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue
$jsonBytes = [System.IO.File]::ReadAllBytes((Resolve-Path $JsonPath).Path)

if (-not $existingEndpoint) {
    Write-Host "[*] Provisioning Migration Endpoint '$EndpointName'..." -ForegroundColor Yellow
    try {
        New-MigrationEndpoint -Gmail -ServiceAccountKeyFileData $jsonBytes -EmailAddress $GoogleAdminEmail -Name $EndpointName -ErrorAction Stop
        Write-Host "[+] Migration Endpoint '$EndpointName' created." -ForegroundColor Green
    } catch {
        Write-Error "Failed to create Migration Endpoint: $_"
        return
    }
} else {
    Write-Host "[+] Migration Endpoint '$EndpointName' is ready." -ForegroundColor Green
}

# --- STEP 5: User Mapping Preparation ---
Write-StepHeader "5. User Migration List Setup"

$finalCsvPath = "$HOME\Downloads\AutoMigrationUsers.csv"

if ($CsvPath -and (Test-Path -Path $CsvPath)) {
    $finalCsvPath = (Resolve-Path $CsvPath).Path
    Write-Host "[+] Using provided CSV: $finalCsvPath" -ForegroundColor Green
} elseif (Test-Path "$HOME\Downloads\Google_to_M365_Migration_Users_Dummy.csv") {
    $finalCsvPath = "$HOME\Downloads\Google_to_M365_Migration_Users_Dummy.csv"
    Write-Host "[+] Using detected test CSV: $finalCsvPath" -ForegroundColor Green
} else {
    Write-Host "[*] Auto-mapping all active Microsoft 365 tenant mailboxes..." -ForegroundColor Yellow
    $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Where-Object { $_.UserPrincipalName -notmatch 'DiscoverySearchMailbox' }

    $csvLines = @("EmailAddress,Username")
    foreach ($m in $mailboxes) {
        $csvLines += "$($m.PrimarySmtpAddress),$($m.PrimarySmtpAddress)"
    }
    $csvLines | Set-Content -Path $finalCsvPath -Encoding UTF8
    Write-Host "[+] Auto-mapped $($mailboxes.Count) mailbox(es)." -ForegroundColor Green
}

$csvBytes = [System.IO.File]::ReadAllBytes($finalCsvPath)

if (-not $TargetDeliveryDomain) {
    $domains = Get-AcceptedDomain | Where-Object { $_.DomainName -match '\.onmicrosoft\.com$' }
    $TargetDeliveryDomain = if ($domains) { $domains[0].DomainName } else { (Get-AcceptedDomain | Where-Object { $_.Default }).DomainName }
    Write-Host "[i] Target Delivery Domain: $TargetDeliveryDomain" -ForegroundColor Cyan
}

# --- STEP 6: Launch Migration Batch ---
Write-StepHeader "6. Launching Migration Batch"

$existingBatch = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue

if (-not $existingBatch) {
    Write-Host "[*] Creating and starting Migration Batch '$BatchName'..." -ForegroundColor Yellow
    try {
        New-MigrationBatch -Name $BatchName `
            -SourceEndpoint $EndpointName `
            -CSVData $csvBytes `
            -TargetDeliveryDomain $TargetDeliveryDomain `
            -AutoStart `
            -AutoComplete:$false `
            -ErrorAction Stop

        Write-Host "[+] Migration Batch '$BatchName' launched successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to start Migration Batch: $_"
        return
    }
} else {
    Write-Host "[+] Migration Batch '$BatchName' exists (Status: $($existingBatch.Status))." -ForegroundColor Yellow
    if ($existingBatch.Status -eq 'Created' -or $existingBatch.Status -eq 'Stopped') {
        Start-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
        Write-Host "[+] Migration Batch started." -ForegroundColor Green
    }
}

# --- STEP 7: Live Synchronization Monitoring ---
Write-StepHeader "7. Live Migration Progress Tracker"
Write-Host "Monitoring sync progress. Press Ctrl+C at any time to exit (cloud sync remains active)." -ForegroundColor Gray

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
                    "User Mailbox"      = $u.Identity
                    "Sync Status"       = $u.Status
                    "Items Transferred" = if ($stat) { $stat.ItemsTransferred } else { 0 }
                    "Data Transferred"  = if ($stat) { $stat.BytesTransferred.ToString() } else { "0 B" }
                    "Percent Complete"  = if ($stat) { "$($stat.PercentageComplete) %" } else { "0 %" }
                    "Error / Notice"    = if ($stat -and $stat.ErrorSummary) { $stat.ErrorSummary } else { "None" }
                }
            }
            $userStats | Format-Table -AutoSize
        }
    } catch {}

    if ($batch.Status -eq 'Completed' -or $batch.Status -eq 'Synced') {
        $isComplete = $true
        Write-Host "`n[√] Migration batch sync is complete!" -ForegroundColor Green
        break
    }

    Start-Sleep -Seconds $MonitorIntervalSeconds
}
