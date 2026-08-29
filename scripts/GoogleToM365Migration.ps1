#Requires -Version 5.1
<#
.SYNOPSIS
    Controlled Google Workspace to Exchange Online migration controller.

.DESCRIPTION
    Defaults to a read-only preflight.  Creating an endpoint, batch, starting a
    batch, and completing a batch are separate operations.  The script never
    creates Google projects or service-account keys and never selects a routing
    domain automatically.

    Before any mutation, complete Microsoft's Google Workspace migration
    prerequisites: routing subdomains, Google domain-wide delegation, a service
    account key kept in an approved secure location, and provisioned MailUsers.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [ValidateSet('Preflight','CreateEndpoint','CreateBatch','StartBatch','Monitor','CompleteBatch','Report')]
    [string]$Mode = 'Preflight',
    [string]$CsvPath,
    [string]$ServiceAccountKeyPath,
    [string]$GoogleAdminEmail,
    [string]$TargetDeliveryDomain,
    [string]$EndpointName = 'GoogleWorkspaceEndpoint',
    [string]$BatchName = 'GoogleWorkspaceMigration',
    [string]$OutputRoot,
    [string[]]$NotificationEmails = @(),
    [ValidateRange(15,3600)][int]$MonitorIntervalSeconds = 60,
    [ValidateRange(1,1440)][int]$MaxMonitorMinutes = 60,
    [switch]$CollectUserStatistics,
    [switch]$InstallExchangeModule,
    [switch]$SkipExchangeConnection,
    [switch]$ApproveGooglePrerequisites,
    [switch]$ApproveCutover
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Started = Get-Date
$script:Findings = [System.Collections.Generic.List[object]]::new()
$script:Actions = [System.Collections.Generic.List[object]]::new()
$script:Connected = $false

function Write-Status {
    param([ValidateSet('INFO','OK','WARN','FAIL')][string]$Level,[string]$Message)
    $color = @{ INFO='Cyan'; OK='Green'; WARN='Yellow'; FAIL='Red' }[$Level]
    Write-Host "[$Level] $Message" -ForegroundColor $color
}
function Add-Finding {
    param([ValidateSet('Critical','High','Medium','Low','Info')][string]$Severity,[string]$Area,[string]$Message,[string]$Resolution='')
    $script:Findings.Add([pscustomobject]@{ Severity=$Severity; Area=$Area; Message=$Message; Resolution=$Resolution })
    $level = if ($Severity -in 'Critical','High') {'FAIL'} elseif ($Severity -in 'Medium','Low') {'WARN'} else {'INFO'}
    Write-Status $level "$Area - $Message"
}
function Add-Action {
    param([string]$Name,[ValidateSet('Applied','Skipped','Failed','NotNeeded')][string]$Result,[string]$Detail='')
    $script:Actions.Add([pscustomobject]@{ Time=(Get-Date).ToString('s'); Action=$Name; Result=$Result; Detail=$Detail })
    $level = if ($Result -eq 'Applied') {'OK'} elseif ($Result -eq 'Failed') {'FAIL'} elseif ($Result -eq 'Skipped') {'WARN'} else {'INFO'}
    Write-Status $level "${Name}: $Result. $Detail"
}
function New-RunFolder {
    $documents = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    if ([string]::IsNullOrWhiteSpace($documents) -or -not (Test-Path -LiteralPath $documents)) { $documents = $env:USERPROFILE }
    $base = if ($OutputRoot) {$OutputRoot} else { Join-Path $documents 'GoogleToM365Migration' }
    $path = Join-Path $base (Get-Date -Format 'yyyyMMdd_HHmmss')
    New-Item -ItemType Directory -Path $path -Force | Out-Null
    $path
}
function Test-EmailAddress {
    param([string]$Value)
    try { $mail = [Net.Mail.MailAddress]::new($Value); return ($mail.Address -eq $Value) } catch { return $false }
}
function Import-MigrationCsv {
    param([string]$Path)
    if (-not $Path) { Add-Finding Critical 'CSV' 'CsvPath is required.' 'Provide a UTF-8 CSV containing EmailAddress and optional Username.'; return @() }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { Add-Finding Critical 'CSV' "CSV was not found: $Path" 'Correct CsvPath.'; return @() }
    try { $rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop) } catch { Add-Finding Critical 'CSV' 'CSV could not be parsed.' $_.Exception.Message; return @() }
    if (-not $rows) { Add-Finding Critical 'CSV' 'CSV contains no rows.' 'Supply at least one migration user.'; return @() }
    $first = $rows[0].PSObject.Properties.Name
    if ($first -notcontains 'EmailAddress') { Add-Finding Critical 'CSV' 'CSV lacks the required EmailAddress column.' 'Use EmailAddress,Username headers.'; return @() }
    $validated = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $rows) {
        $target = ([string]$row.EmailAddress).Trim()
        $source = if ($row.PSObject.Properties.Name -contains 'Username' -and $row.Username) { ([string]$row.Username).Trim() } else { $target }
        if (-not (Test-EmailAddress $target) -or -not (Test-EmailAddress $source)) {
            Add-Finding High 'CSV' "Invalid address mapping: '$source' -> '$target'" 'Correct the affected CSV row.'
        } else { $validated.Add([pscustomobject]@{ EmailAddress=$target; Username=$source }) }
    }
    foreach ($group in ($validated | Group-Object EmailAddress | Where-Object Count -gt 1)) { Add-Finding High 'CSV' "Duplicate Microsoft 365 target: $($group.Name)" 'Each target can appear only once per batch.' }
    foreach ($group in ($validated | Group-Object Username | Where-Object Count -gt 1)) { Add-Finding High 'CSV' "Duplicate Google source: $($group.Name)" 'Each Google source can appear only once per batch.' }
    $validated.ToArray()
}
function Test-ServiceAccountKey {
    param([string]$Path)
    if (-not $Path) { Add-Finding Critical 'Google key' 'ServiceAccountKeyPath is required.' 'Provide an existing approved service-account JSON key; do not generate one with this migration script.'; return $null }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { Add-Finding Critical 'Google key' "Key file was not found: $Path" 'Correct ServiceAccountKeyPath.'; return $null }
    try { $key = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json } catch { Add-Finding Critical 'Google key' 'Key file is not valid JSON.' $_.Exception.Message; return $null }
    if ($key.type -ne 'service_account' -or -not $key.client_id -or -not $key.client_email -or -not $key.private_key) {
        Add-Finding Critical 'Google key' 'Key file is not a complete Google service-account key.' 'Use the exact JSON key issued for the migration service account.'; return $null
    }
    Write-Status OK "Validated service account key for $($key.client_email). Private key material is never written to reports."
    $key
}
function Ensure-ExchangeModule {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) { Import-Module ExchangeOnlineManagement -ErrorAction Stop; return $true }
    if (-not $InstallExchangeModule) { Add-Finding Critical 'Exchange module' 'ExchangeOnlineManagement is not installed.' 'Install it from an approved repository, or rerun with -InstallExchangeModule.'; return $false }
    if ($PSCmdlet.ShouldProcess('Current user PowerShell modules','Install ExchangeOnlineManagement')) {
        try { Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop; Import-Module ExchangeOnlineManagement -ErrorAction Stop; Add-Action 'Install ExchangeOnlineManagement' 'Applied' 'Module installed for the current user.'; return $true }
        catch { Add-Action 'Install ExchangeOnlineManagement' 'Failed' $_.Exception.Message; return $false }
    }
    return $false
}
function Connect-ExchangeSafely {
    if ($SkipExchangeConnection) { Add-Action 'Connect to Exchange Online' 'Skipped' 'SkipExchangeConnection was specified.'; return $false }
    if (-not (Ensure-ExchangeModule)) { return $false }
    try {
        if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) { Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop }
        $script:Connected = $true; Add-Action 'Connect to Exchange Online' 'Applied' 'Connected using the signed-in administrator.'; return $true
    } catch { Add-Finding Critical 'Exchange Online' 'Connection failed.' $_.Exception.Message; return $false }
}
function Test-ExchangePrerequisites {
    param([object[]]$Mappings,[switch]$ValidateRouting,[switch]$ValidateRecipients,[switch]$TestGoogleConnectivity)
    if (-not $script:Connected) { return }
    if ($ValidateRouting) {
        $accepted = @(Get-AcceptedDomain -ErrorAction Stop)
        if (-not $TargetDeliveryDomain) { Add-Finding Critical 'Routing domain' 'TargetDeliveryDomain is required.' 'Specify the verified migration routing subdomain; do not use the tenant onmicrosoft.com domain.' }
        else {
            $match = $accepted | Where-Object { $_.DomainName.ToString().Equals($TargetDeliveryDomain,[StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
            if (-not $match) { Add-Finding Critical 'Routing domain' "$TargetDeliveryDomain is not an accepted domain." 'Add and verify the migration routing subdomain before creating a batch.' }
            elseif ($TargetDeliveryDomain -match '\.onmicrosoft\.com$') { Add-Finding Critical 'Routing domain' 'The built-in tenant onmicrosoft.com domain is blocked as a migration routing domain.' 'Use a verified Google Workspace routing subdomain, for example o365.contoso.com.' }
            else { Add-Finding Info 'Routing domain' "$TargetDeliveryDomain is an accepted domain." 'Verify matching Google alias-domain and MX routing outside this script.' }
        }
    }
    if ($ValidateRecipients) {
        foreach ($mapping in $Mappings) {
            try {
                $recipient = Get-Recipient -Identity $mapping.EmailAddress -ErrorAction Stop
                if ($recipient.RecipientTypeDetails -ne 'MailUser') { Add-Finding High 'Recipient readiness' "$($mapping.EmailAddress) is $($recipient.RecipientTypeDetails), not MailUser." 'Provision migration targets as MailUsers before the Google batch is started.' }
            } catch { Add-Finding High 'Recipient readiness' "$($mapping.EmailAddress) was not found in Exchange Online." 'Provision this migration target as a MailUser first.' }
        }
    }
    if ($TestGoogleConnectivity -and $ServiceAccountKeyPath -and $GoogleAdminEmail -and (Test-Path -LiteralPath $ServiceAccountKeyPath)) {
        try {
            # This is Microsoft's supported live verification before endpoint creation.
            Test-MigrationServerAvailability -Gmail -ServiceAccountKeyFileData (Get-EndpointBytes $ServiceAccountKeyPath) -EmailAddress $GoogleAdminEmail -ErrorAction Stop | Out-Null
            Add-Action 'Test Google migration connectivity' 'Applied' 'Exchange Online accepted the Google service-account connectivity test.'
        } catch { Add-Finding Critical 'Google migration connectivity' 'Exchange Online could not validate the Google service-account connection.' $_.Exception.Message }
    }
}
function Export-ValidatedMigrationCsv {
    param([object[]]$Mappings)
    if (-not $Mappings) { return $null }
    $path = Join-Path $script:RunFolder 'ValidatedMigrationUsers.csv'
    $Mappings | Select-Object EmailAddress,Username | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    Add-Action 'Create validated migration CSV' 'Applied' $path
    $path
}
function Assert-MutationReady {
    if (-not $ApproveGooglePrerequisites) { Add-Finding Critical 'Approval' 'Google prerequisites were not explicitly acknowledged.' 'Rerun with -ApproveGooglePrerequisites only after routing, domain-wide delegation, API access, and MailUser provisioning are verified.'; return $false }
    if (@($script:Findings | Where-Object Severity -in 'Critical','High').Count -gt 0) { Add-Finding Critical 'Safety gate' 'A blocking preflight finding exists.' 'Resolve blocking findings before changing Exchange Online.'; return $false }
    if (-not $script:Connected) { Add-Finding Critical 'Exchange Online' 'An Exchange Online connection is required for this mode.' 'Remove -SkipExchangeConnection and connect as a migration administrator.'; return $false }
    $true
}
function Get-EndpointBytes { param([string]$Path) [IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path) }
function Invoke-CreateEndpoint {
    param($Key)
    if (-not (Assert-MutationReady)) { return }
    $existing = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue
    if ($existing) { Add-Action 'Create migration endpoint' 'NotNeeded' "Endpoint $EndpointName already exists; it was not changed."; return }
    if ($PSCmdlet.ShouldProcess($EndpointName,"Create Gmail migration endpoint for $($Key.client_email)")) {
        try { New-MigrationEndpoint -Gmail -ServiceAccountKeyFileData (Get-EndpointBytes $ServiceAccountKeyPath) -EmailAddress $GoogleAdminEmail -Name $EndpointName -ErrorAction Stop | Out-Null; Add-Action 'Create migration endpoint' 'Applied' $EndpointName }
        catch { Add-Action 'Create migration endpoint' 'Failed' $_.Exception.Message }
    }
}
function Invoke-CreateBatch {
    param($Mappings)
    if (-not (Assert-MutationReady)) { return }
    $endpoint = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue
    if (-not $endpoint) { Add-Finding Critical 'Migration endpoint' "Endpoint $EndpointName was not found." 'Run CreateEndpoint after preflight passes.'; return }
    $existing = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
    if ($existing) { Add-Action 'Create migration batch' 'NotNeeded' "Batch $BatchName already exists with status $($existing.Status); it was not changed."; return }
    if (-not $script:ValidatedCsvPath) { Add-Finding Critical 'CSV' 'No validated migration CSV is available.' 'Run the current operation with a valid CsvPath.'; return }
    $data = [IO.File]::ReadAllBytes($script:ValidatedCsvPath)
    $params = @{ Name=$BatchName; SourceEndpoint=$EndpointName; CSVData=$data; TargetDeliveryDomain=$TargetDeliveryDomain; AutoComplete=$false; ErrorAction='Stop' }
    if ($NotificationEmails.Count) { $params['NotificationEmails'] = $NotificationEmails }
    if ($PSCmdlet.ShouldProcess($BatchName,"Create a Google Workspace migration batch for $($Mappings.Count) validated users without starting it")) {
        try { New-MigrationBatch @params | Out-Null; Add-Action 'Create migration batch' 'Applied' 'Created without AutoStart. Review then use -Mode StartBatch.' }
        catch { Add-Action 'Create migration batch' 'Failed' $_.Exception.Message }
    }
}
function Invoke-StartBatch {
    if (-not (Assert-MutationReady)) { return }
    $batch = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
    if (-not $batch) { Add-Finding Critical 'Migration batch' "$BatchName was not found." 'Create and review the batch before starting it.'; return }
    if ($batch.Status -notin 'Created','Stopped') { Add-Action 'Start migration batch' 'NotNeeded' "Batch status is $($batch.Status)."; return }
    if ($PSCmdlet.ShouldProcess($BatchName,'Start migration batch')) { try { Start-MigrationBatch -Identity $BatchName -ErrorAction Stop | Out-Null; Add-Action 'Start migration batch' 'Applied' 'Batch started.' } catch { Add-Action 'Start migration batch' 'Failed' $_.Exception.Message } }
}
function Invoke-MonitorBatch {
    if (-not $script:Connected) { Add-Finding Critical 'Exchange Online' 'Monitoring requires an active Exchange Online connection.' 'Remove -SkipExchangeConnection and sign in as a migration administrator.'; return }
    $end = (Get-Date).AddMinutes($MaxMonitorMinutes)
    do {
        $batch = Get-MigrationBatch -Identity $BatchName -ErrorAction Stop
        Write-Status INFO "Batch ${BatchName}: $($batch.Status) | Synced=$($batch.SyncedCount) | In progress=$($batch.InProgressCount) | Failed=$($batch.FailedCount)"
        if ($CollectUserStatistics) {
            try {
                $stats = @(Get-MigrationUser -BatchId $BatchName -ErrorAction Stop | ForEach-Object { Get-MigrationUserStatistics -Identity $_.Identity -ErrorAction Stop })
                $stats | Select-Object Identity,Status,ItemsTransferred,BytesTransferred,PercentageComplete,ErrorSummary | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'MigrationUserStatistics.csv') -NoTypeInformation -Encoding UTF8
            } catch { Add-Finding Medium 'Monitoring' 'Detailed user statistics could not be collected; batch monitoring will continue.' $_.Exception.Message }
        }
        if ($batch.Status -in 'Completed','Failed','Stopped','Synced','SyncedWithErrors') { break }
        Start-Sleep -Seconds $MonitorIntervalSeconds
    } while ((Get-Date) -lt $end)
    Add-Action 'Monitor migration batch' 'Applied' "Monitoring finished; inspect MigrationUserStatistics.csv before taking the next step."
}
function Invoke-CompleteBatch {
    if (-not $ApproveCutover) { Add-Finding Critical 'Cutover approval' 'Completion requires -ApproveCutover.' 'Verify user statistics, mail routing, support readiness, and licence plan before cutover.'; return }
    if (-not (Assert-MutationReady)) { return }
    $batch = Get-MigrationBatch -Identity $BatchName -ErrorAction Stop
    if ($batch.Status -notmatch '^Synced') { Add-Finding Critical 'Migration batch' "Batch status is $($batch.Status), not Synced." 'Do not complete until statistics are reviewed and the batch is Synced.'; return }
    if ($PSCmdlet.ShouldProcess($BatchName,'Complete migration batch and begin cutover routing changes')) { try { Complete-MigrationBatch -Identity $BatchName -ErrorAction Stop | Out-Null; Add-Action 'Complete migration batch' 'Applied' 'Completion requested. Monitor until status becomes Completed, then assign Exchange licences.' } catch { Add-Action 'Complete migration batch' 'Failed' $_.Exception.Message } }
}
function Write-Report {
    $summary = [pscustomobject]@{ ScriptVersion='1.0.0'; Started=$script:Started; Completed=Get-Date; Mode=$Mode; Endpoint=$EndpointName; Batch=$BatchName; Findings=$script:Findings; Actions=$script:Actions }
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $script:RunFolder 'MigrationControllerReport.json') -Encoding UTF8
    $script:Findings | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'MigrationPreflightFindings.csv') -NoTypeInformation -Encoding UTF8
    $script:Actions | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'MigrationActions.csv') -NoTypeInformation -Encoding UTF8
    Write-Status OK "Reports written to $script:RunFolder"
}
function Invoke-Main {
    $script:RunFolder = New-RunFolder
    Write-Status INFO "Google to Microsoft 365 Migration Controller - $Mode mode"
    $mappings = @(); $key = $null
    $needsMappings = $Mode -in 'Preflight','CreateBatch'
    $needsKey = $Mode -in 'Preflight','CreateEndpoint'
    $needsGoogleAdmin = $Mode -in 'Preflight','CreateEndpoint'
    $needsRouting = $Mode -in 'Preflight','CreateBatch'
    if ($needsMappings) { $mappings = @(Import-MigrationCsv $CsvPath); $script:ValidatedCsvPath = Export-ValidatedMigrationCsv $mappings }
    if ($needsKey) { $key = Test-ServiceAccountKey $ServiceAccountKeyPath }
    if ($needsGoogleAdmin -and $GoogleAdminEmail -and -not (Test-EmailAddress $GoogleAdminEmail)) { Add-Finding Critical 'Google admin' 'GoogleAdminEmail is invalid.' 'Provide a valid Google Workspace super-admin address.' }
    elseif ($needsGoogleAdmin -and -not $GoogleAdminEmail) { Add-Finding Critical 'Google admin' 'GoogleAdminEmail is required.' 'Provide the delegated Google Workspace administrator address.' }
    if ($Mode -ne 'Report') { Connect-ExchangeSafely | Out-Null }
    if ($Mode -eq 'Preflight') { Test-ExchangePrerequisites $mappings -ValidateRouting -ValidateRecipients -TestGoogleConnectivity }
    elseif ($Mode -eq 'CreateBatch') { Test-ExchangePrerequisites $mappings -ValidateRouting -ValidateRecipients }
    elseif ($Mode -eq 'CreateEndpoint') { Test-ExchangePrerequisites @() -TestGoogleConnectivity }
    switch ($Mode) {
        'CreateEndpoint' { if ($key) { Invoke-CreateEndpoint $key } }
        'CreateBatch'    { Invoke-CreateBatch $mappings }
        'StartBatch'     { Invoke-StartBatch }
        'Monitor'        { Invoke-MonitorBatch }
        'CompleteBatch'  { Invoke-CompleteBatch }
    }
    Write-Report
    if (@($script:Findings | Where-Object { $_.Severity -in 'Critical','High' }).Count -gt 0) { Write-Status FAIL 'Blocking findings exist. No further migration action should be taken.' }
}
try { Invoke-Main } catch { Write-Error "Unexpected failure at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"; if ($script:RunFolder) { Write-Report } }
