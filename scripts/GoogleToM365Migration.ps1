#Requires -Version 5.1
<#
.SYNOPSIS
    Interactive Google Workspace to Microsoft 365 migration helper (Exchange Online).

.DESCRIPTION
    Fully interactive and self-detecting. Auto-detects the migration CSV, the
    service-account key, the Google administrator email, and the target routing
    domain. Prompts only for values it cannot detect on its own. Read-only by
    default; never creates, starts, or completes anything without approval.

    Run with no arguments for a guided menu. Use -NonInteractive for RMM/SYSTEM
    automation (fails fast instead of prompting).

.EXAMPLE
    .\Run-Migration.ps1

.EXAMPLE
    .\Run-Migration.ps1 -Mode Preflight

.EXAMPLE
    .\Run-Migration.ps1 -Mode CreateEndpoint -NonInteractive
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [ValidateSet('Preflight','CreateEndpoint','CreateBatch','StartBatch','Monitor','CompleteBatch','Report')]
    [string]$Mode,
    [string]$CsvPath,
    [string]$ServiceAccountKeyPath,
    [string]$GoogleAdminEmail,
    [string]$TargetDeliveryDomain,
    [string]$EndpointName = 'GoogleWorkspaceEndpoint',
    [string]$BatchName = 'GoogleWorkspaceMigration',
    [string]$OutputRoot,
    [string[]]$NotificationEmails,
    [int]$MonitorIntervalSeconds = 60,
    [int]$MaxMonitorMinutes = 120,
    [switch]$CollectUserStatistics,
    [switch]$InstallExchangeModule,
    [switch]$SkipExchangeConnection,
    [switch]$ApproveGooglePrerequisites,
    [switch]$ApproveCutover,
    [switch]$NonInteractive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Findings = [System.Collections.Generic.List[object]]::new()
$script:Actions = [System.Collections.Generic.List[object]]::new()
$script:ExchangeConnected = $false
$script:RunFolder = $null
$script:ResolvedCsv = $null
$script:ResolvedKey = $null
$script:ResolvedAdmin = $null
$script:ResolvedDomain = $null

function Write-Status {
    param([ValidateSet('INFO','OK','WARN','FAIL')][string]$Level,[string]$Message)
    Write-Host "[$Level] $Message" -ForegroundColor (@{INFO='Cyan';OK='Green';WARN='Yellow';FAIL='Red'}[$Level])
}
function Add-Finding {
    param([ValidateSet('Critical','High','Medium','Low','Info')][string]$Severity,[string]$Area,[string]$Message,[string]$Resolution='')
    $script:Findings.Add([pscustomobject]@{Severity=$Severity;Area=$Area;Message=$Message;Resolution=$Resolution})
    $level=if($Severity -in 'Critical','High'){'FAIL'}elseif($Severity -in 'Medium','Low'){'WARN'}else{'INFO'}
    Write-Status $level "$Area - $Message"
}
function Add-Action {
    param([string]$Name,[ValidateSet('Applied','Skipped','Failed','NotNeeded')][string]$Result,[string]$Detail='')
    $script:Actions.Add([pscustomobject]@{Time=(Get-Date).ToString('s');Action=$Name;Result=$Result;Detail=$Detail})
    $level=if($Result -eq 'Applied'){'OK'}elseif($Result -eq 'Failed'){'FAIL'}elseif($Result -eq 'Skipped'){'WARN'}else{'INFO'}
    Write-Status $level "${Name}: $Result. $Detail"
}
function New-RunFolder {
    $documents=[Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    if([string]::IsNullOrWhiteSpace($documents) -or -not(Test-Path -LiteralPath $documents)){$documents=$env:USERPROFILE}
    $base=if($OutputRoot){$OutputRoot}else{Join-Path $documents 'GoogleToM365Migration'}
    $run=Join-Path $base (Get-Date -Format 'yyyyMMdd_HHmmss')
    New-Item -ItemType Directory -Path $run -Force | Out-Null
    $run
}

# --- Interactive helpers ---
function Read-Value {
    param([string]$Prompt,[string]$Default='')
    if($NonInteractive){ return $Default }
    if($Default){
        $resp=Read-Host "$Prompt [$Default]"
        if([string]::IsNullOrWhiteSpace($resp)){ return $Default }
        return $resp.Trim()
    }
    $resp=Read-Host $Prompt
    if($null -eq $resp){ return '' }
    return $resp.Trim()
}
function Confirm-YesNo {
    param([string]$Message,[bool]$DefaultYes=$true)
    if($NonInteractive){ return $DefaultYes }
    $suffix=if($DefaultYes){'[Y/n]'}else{'[y/N]'}
    $resp=Read-Host "$Message $suffix"
    if([string]::IsNullOrWhiteSpace($resp)){ return $DefaultYes }
    return $resp -match '^(y|yes)$'
}
function Show-ModeMenu {
    $modes=@('Preflight','CreateEndpoint','CreateBatch','StartBatch','Monitor','CompleteBatch','Report')
    $desc=@{
        'Preflight'='Validate CSV, key, admin email, routing domain (read-only)'
        'CreateEndpoint'='Create the Gmail migration endpoint'
        'CreateBatch'='Create the migration batch from the CSV (no auto-start)'
        'StartBatch'='Start the migration batch'
        'Monitor'='Monitor batch progress until completion'
        'CompleteBatch'='Complete the batch and begin cutover routing changes'
        'Report'='Generate a migration report from the last run'
    }
    Write-Host ''
    Write-Host 'Select a mode:' -ForegroundColor Cyan
    for($i=0;$i -lt $modes.Count;$i++){ Write-Host "  $($i+1)) $($modes[$i]) - $($desc[$modes[$i]])" }
    $choice=Read-Value 'Enter number' '1'
    $idx=0
    if([int]::TryParse($choice,[ref]$idx) -and $idx -ge 1 -and $idx -le $modes.Count){ return $modes[$idx-1] }
    $match=$modes | Where-Object { $_ -like "*$choice*" } | Select-Object -First 1
    if($match){ return $match }
    return $modes[0]
}

# --- Auto-detection ---
function Find-MigrationCsv {
    if($CsvPath){ return [IO.Path]::GetFullPath($CsvPath) }
    $locations=@((Get-Location).Path,(Join-Path (Get-Location).Path 'examples'),(Join-Path (Get-Location).Path 'config'),[Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)+'\Downloads',[Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))
    $patterns=@('migration-users.csv','*migration*.csv','*users*.csv')
    foreach($loc in $locations){
        if(-not $loc -or -not(Test-Path -LiteralPath $loc)){ continue }
        foreach($pat in $patterns){
            $found=Get-ChildItem -LiteralPath $loc -Filter $pat -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if($found){ return $found.FullName }
        }
    }
    return $null
}
function Find-ServiceAccountKey {
    if($ServiceAccountKeyPath){ return [IO.Path]::GetFullPath($ServiceAccountKeyPath) }
    $locations=@((Get-Location).Path,[Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)+'\Downloads',[Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments),(Join-Path $env:LOCALAPPDATA '.m365-migration'))
    foreach($loc in $locations){
        if(-not $loc -or -not(Test-Path -LiteralPath $loc)){ continue }
        $jsons=Get-ChildItem -LiteralPath $loc -Filter '*.json' -File -ErrorAction SilentlyContinue
        foreach($j in $jsons){
            try {
                $obj=Get-Content -LiteralPath $j.FullName -Raw -ErrorAction Stop | ConvertFrom-Json
                if($obj.type -eq 'service_account' -and $obj.private_key){ return $j.FullName }
            } catch { }
        }
    }
    return $null
}
function Get-DetectedAdminEmail {
    if($GoogleAdminEmail){ return $GoogleAdminEmail }
    try {
        $gcloud=Get-Command gcloud.exe -ErrorAction SilentlyContinue
        if($gcloud){
            $acct=& $gcloud.Source config get-value account 2>$null
            if(-not $acct){ $acct=& $gcloud.Source auth list --filter=status:ACTIVE --format=value(account) 2>$null | Select-Object -First 1 }
            if($acct){ return ([string]$acct).Trim() }
        }
    } catch { }
    return $null
}
function Get-DetectedRoutingDomain {
    if($TargetDeliveryDomain){ return $TargetDeliveryDomain }
    if($script:ExchangeConnected){
        try {
            $domains=@(Get-AcceptedDomain -ErrorAction Stop | Where-Object { $_.DomainName -notmatch '\.onmicrosoft\.com$' })
            if($domains.Count -eq 1){ return ([string]$domains[0].DomainName).Trim() }
            if($domains.Count -gt 1){
                Write-Host 'Detected accepted domains:' -ForegroundColor Cyan
                for($i=0;$i -lt $domains.Count;$i++){ Write-Host "  $($i+1)) $($domains[$i].DomainName)" }
                $choice=Read-Value 'Select the target routing domain' '1'
                $idx=0
                if([int]::TryParse($choice,[ref]$idx) -and $idx -ge 1 -and $idx -le $domains.Count){ return ([string]$domains[$idx-1].DomainName).Trim() }
            }
        } catch { }
    }
    return $null
}
function Resolve-CsvPath {
    $found=Find-MigrationCsv
    if($found){ Write-Status INFO "Auto-detected migration CSV: $found"; $script:ResolvedCsv=$found; return $found }
    if($NonInteractive){ Add-Finding Critical 'CSV' 'No migration CSV was found and -CsvPath was not provided.' 'Provide -CsvPath or place migration-users.csv next to the script.'; return $null }
    $resp=Read-Value 'Enter the full path to the migration CSV (required)'
    if([string]::IsNullOrWhiteSpace($resp)){ Add-Finding Critical 'CSV' 'No migration CSV was provided.' 'Provide the CSV path.'; return $null }
    $script:ResolvedCsv=[IO.Path]::GetFullPath($resp)
    $script:ResolvedCsv
}
function Resolve-KeyPath {
    $found=Find-ServiceAccountKey
    if($found){ Write-Status INFO "Auto-detected service-account key: $found"; $script:ResolvedKey=$found; return $found }
    if($NonInteractive){ Add-Finding Critical 'Service-account key' 'No service-account key was found and -ServiceAccountKeyPath was not provided.' 'Provide -ServiceAccountKeyPath or run GoogleCloudMigrationBootstrap.ps1 -Mode CreateServiceAccountKey.'; return $null }
    $resp=Read-Value 'Enter the full path to the service-account JSON key (required)'
    if([string]::IsNullOrWhiteSpace($resp)){ Add-Finding Critical 'Service-account key' 'No service-account key was provided.' 'Provide the key path.'; return $null }
    $script:ResolvedKey=[IO.Path]::GetFullPath($resp)
    $script:ResolvedKey
}
function Resolve-AdminEmail {
    $found=Get-DetectedAdminEmail
    if($found){ Write-Status INFO "Auto-detected Google administrator: $found"; $script:ResolvedAdmin=$found; return $found }
    if($NonInteractive){ Add-Finding Critical 'Google administrator' 'No Google administrator email was detected and -GoogleAdminEmail was not provided.' 'Provide -GoogleAdminEmail.'; return $null }
    $resp=Read-Value 'Enter the Google Workspace administrator email (required)'
    if([string]::IsNullOrWhiteSpace($resp)){ Add-Finding Critical 'Google administrator' 'No Google administrator email was provided.' 'Provide the administrator email.'; return $null }
    $script:ResolvedAdmin=$resp.Trim()
    $script:ResolvedAdmin
}
function Resolve-RoutingDomain {
    $found=Get-DetectedRoutingDomain
    if($found){ Write-Status INFO "Auto-detected target routing domain: $found"; $script:ResolvedDomain=$found; return $found }
    if($NonInteractive){ Add-Finding Critical 'Target routing domain' 'No target routing domain was detected and -TargetDeliveryDomain was not provided.' 'Provide -TargetDeliveryDomain.'; return $null }
    $resp=Read-Value 'Enter the target routing domain (e.g. contoso.com) (required)'
    if([string]::IsNullOrWhiteSpace($resp)){ Add-Finding Critical 'Target routing domain' 'No target routing domain was provided.' 'Provide the routing domain.'; return $null }
    $script:ResolvedDomain=$resp.Trim()
    $script:ResolvedDomain
}

# --- Exchange Online ---
function Connect-Exchange {
    if($SkipExchangeConnection){ return }
    if($script:ExchangeConnected){ return }
    if(-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)){
        if($InstallExchangeModule -or (-not $NonInteractive -and (Confirm-YesNo 'ExchangeOnlineManagement module is missing. Install it now?'))){
            if($PSCmdlet.ShouldProcess('Current user PowerShell modules','Install ExchangeOnlineManagement')){
                try { Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop; Add-Action 'Install ExchangeOnlineManagement' 'Applied' 'Installed for the current user.' } catch { Add-Action 'Install ExchangeOnlineManagement' 'Failed' $_.Exception.Message; return }
            }
        } else { Add-Finding Critical 'Exchange Online' 'ExchangeOnlineManagement module is not installed.' 'Run with -InstallExchangeModule or install the module manually.'; return }
    }
    try {
        if(-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)){ Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop }
        $script:ExchangeConnected=$true
        Add-Action 'Connect Exchange Online' 'Applied' 'Connected.'
    } catch {
        $msg=($_.Exception.Message -replace '\s+',' ').Trim()
        Add-Action 'Connect Exchange Online' 'Failed' $msg
        if($msg -match 'window handle|interactive|device code'){ Add-Finding High 'Exchange Online' 'Interactive sign-in could not complete in this session.' 'Run the script from an interactive PowerShell window so the MFA sign-in prompt can appear.' }
    }
}

# --- Modes ---
function Invoke-Preflight {
    $csv=Resolve-CsvPath
    if($csv){
        if(-not(Test-Path -LiteralPath $csv)){ Add-Finding Critical 'CSV' "CSV not found: $csv" 'Provide a valid path.' }
        else {
            try {
                $rows=@(Import-Csv -LiteralPath $csv -ErrorAction Stop)
                if($rows.Count -eq 0){ Add-Finding Critical 'CSV' 'The CSV contains no data rows.' 'Add at least one user row.' }
                else {
                    $headers=@($rows[0].PSObject.Properties.Name)
                    $missing=@(@('EmailAddress','TargetMailbox','MailboxType') | Where-Object { $_ -notin $headers })
                    if($missing.Count -gt 0){ Add-Finding Critical 'CSV' "CSV is missing required columns: $($missing -join ', ')" 'Expected columns: EmailAddress, TargetMailbox, MailboxType.' }
                    else {
                        $bad=@($rows | Where-Object { [string]::IsNullOrWhiteSpace($_.EmailAddress) -or [string]::IsNullOrWhiteSpace($_.TargetMailbox) })
                        if($bad.Count -gt 0){ Add-Finding High 'CSV' "$($bad.Count) row(s) have empty EmailAddress or TargetMailbox." 'Fix or remove those rows.' }
                        else { Add-Action 'Validate migration CSV' 'Applied' "$($rows.Count) user(s) validated." }
                    }
                }
            } catch { Add-Finding Critical 'CSV' "Could not read CSV: $csv" $_.Exception.Message }
        }
    }
    $key=Resolve-KeyPath
    if($key){
        if(-not(Test-Path -LiteralPath $key)){ Add-Finding Critical 'Service-account key' "Key not found: $key" 'Provide a valid path.' }
        else {
            try {
                $k=Get-Content -LiteralPath $key -Raw -ErrorAction Stop | ConvertFrom-Json
                if($k.type -ne 'service_account' -or -not $k.private_key){ Add-Finding Critical 'Service-account key' 'The file is not a valid service-account JSON key.' 'Recreate the key with GoogleCloudMigrationBootstrap.ps1.' }
                else { Add-Action 'Validate service-account key' 'Applied' "Client email: $($k.client_email)" }
            } catch { Add-Finding Critical 'Service-account key' "Could not parse key: $key" $_.Exception.Message }
        }
    }
    $admin=Resolve-AdminEmail
    if($admin){ Add-Action 'Google administrator' 'Applied' $admin }
    Connect-Exchange
    $domain=Resolve-RoutingDomain
    if($domain){ Add-Action 'Target routing domain' 'Applied' $domain }
}
function Invoke-CreateEndpoint {
    $key=Resolve-KeyPath
    if(-not $key){ return }
    if(-not(Test-Path -LiteralPath $key)){ Add-Finding Critical 'Service-account key' "Key not found: $key" 'Provide a valid path.'; return }
    $admin=Resolve-AdminEmail
    if(-not $admin){ return }
    if(-not $ApproveGooglePrerequisites){
        if($NonInteractive){ Add-Finding Critical 'Approval' 'Endpoint creation requires -ApproveGooglePrerequisites.' 'Confirm Google-side prerequisites (service account, key, domain-wide delegation) are complete.'; return }
        if(-not (Confirm-YesNo 'Create the Gmail migration endpoint? (Google-side prerequisites must already be complete)')){ Add-Action 'Create migration endpoint' 'Skipped' 'Declined by user.'; return }
    }
    Connect-Exchange
    if(-not $script:ExchangeConnected){ return }
    try {
        $k=Get-Content -LiteralPath $key -Raw -ErrorAction Stop | ConvertFrom-Json
        $existing=Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue
        if($existing){ Add-Action 'Create migration endpoint' 'NotNeeded' "Endpoint '$EndpointName' already exists." }
        else {
            if($PSCmdlet.ShouldProcess($EndpointName,"Create Gmail migration endpoint for $($k.client_email)")){
                New-MigrationEndpoint -Gmail -ServiceAccountCredentialFile $key -EmailAddress $admin -Name $EndpointName -ErrorAction Stop | Out-Null
                Add-Action 'Create migration endpoint' 'Applied' "Endpoint '$EndpointName' created for $($k.client_email)."
            }
        }
    } catch { Add-Action 'Create migration endpoint' 'Failed' $_.Exception.Message }
}
function Invoke-CreateBatch {
    $csv=Resolve-CsvPath
    if(-not $csv){ return }
    if(-not(Test-Path -LiteralPath $csv)){ Add-Finding Critical 'CSV' "CSV not found: $csv" 'Provide a valid path.'; return }
    $domain=Resolve-RoutingDomain
    if(-not $domain){ return }
    if(-not $ApproveGooglePrerequisites){
        if($NonInteractive){ Add-Finding Critical 'Approval' 'Batch creation requires -ApproveGooglePrerequisites.' 'Confirm the endpoint exists and the CSV is final.'; return }
        if(-not (Confirm-YesNo 'Create the migration batch from the CSV? (it will NOT auto-start)')){ Add-Action 'Create migration batch' 'Skipped' 'Declined by user.'; return }
    }
    Connect-Exchange
    if(-not $script:ExchangeConnected){ return }
    try {
        $mappings=@(Import-Csv -LiteralPath $csv -ErrorAction Stop)
        $existing=Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
        if($existing){ Add-Action 'Create migration batch' 'NotNeeded' "Batch '$BatchName' already exists." }
        else {
            if($PSCmdlet.ShouldProcess($BatchName,"Create a Google Workspace migration batch for $($mappings.Count) validated users without starting it")){
                New-MigrationBatch -Name $BatchName -SourceEndpoint $EndpointName -TargetDeliveryDomain $domain -CSVData ([System.IO.File]::ReadAllBytes($csv)) -NotificationEmails $NotificationEmails -AutoStart:$false -ErrorAction Stop | Out-Null
                Add-Action 'Create migration batch' 'Applied' "Batch '$BatchName' created for $($mappings.Count) user(s) targeting $domain."
            }
        }
    } catch { Add-Action 'Create migration batch' 'Failed' $_.Exception.Message }
}
function Invoke-StartBatch {
    if(-not $ApproveGooglePrerequisites){
        if($NonInteractive){ Add-Finding Critical 'Approval' 'Starting the batch requires -ApproveGooglePrerequisites.' 'Confirm the batch is final before starting.'; return }
        if(-not (Confirm-YesNo 'Start the migration batch now?')){ Add-Action 'Start migration batch' 'Skipped' 'Declined by user.'; return }
    }
    Connect-Exchange
    if(-not $script:ExchangeConnected){ return }
    try {
        $batch=Get-MigrationBatch -Identity $BatchName -ErrorAction Stop
        if($batch.Status -in 'Synced','Completed','CompletedWithErrors'){ Add-Action 'Start migration batch' 'NotNeeded' "Batch is already in state $($batch.Status)." }
        else {
            if($PSCmdlet.ShouldProcess($BatchName,'Start migration batch')){
                Start-MigrationBatch -Identity $BatchName -ErrorAction Stop | Out-Null
                Add-Action 'Start migration batch' 'Applied' 'Batch started.'
            }
        }
    } catch { Add-Action 'Start migration batch' 'Failed' $_.Exception.Message }
}
function Invoke-Monitor {
    Connect-Exchange
    if(-not $script:ExchangeConnected){ return }
    try {
        $deadline=(Get-Date).AddMinutes($MaxMonitorMinutes)
        do {
            $batch=Get-MigrationBatch -Identity $BatchName -ErrorAction Stop
            $stats=Get-MigrationBatch -Identity $BatchName -IncludeReport -ErrorAction Stop
            $synced=if($stats.Report){$stats.Report.SyncedItemCount}else{0}
            $failed=if($stats.Report){$stats.Report.FailedItemCount}else{0}
            Write-Status INFO "Batch '$BatchName' status: $($batch.Status) | Synced: $synced | Failed: $failed"
            if($batch.Status -in 'Completed','CompletedWithErrors','Failed','Stopped'){ break }
            if((Get-Date) -gt $deadline){ Add-Finding Medium 'Monitor' "Monitor window of $MaxMonitorMinutes minutes elapsed; batch still $($batch.Status)." 'Re-run Monitor later.'; break }
            Start-Sleep -Seconds $MonitorIntervalSeconds
        } while($true)
        Add-Action 'Monitor migration batch' 'Applied' "Final status: $($batch.Status)."
    } catch { Add-Action 'Monitor migration batch' 'Failed' $_.Exception.Message }
}
function Invoke-CompleteBatch {
    if(-not $ApproveCutover){
        if($NonInteractive){ Add-Finding Critical 'Approval' 'Completing the batch requires -ApproveCutover.' 'Confirm the migration is verified before cutover.'; return }
        if(-not (Confirm-YesNo 'Complete the migration batch and begin cutover routing changes?')){ Add-Action 'Complete migration batch' 'Skipped' 'Declined by user.'; return }
    }
    Connect-Exchange
    if(-not $script:ExchangeConnected){ return }
    try {
        $batch=Get-MigrationBatch -Identity $BatchName -ErrorAction Stop
        if($batch.Status -in 'Completed','CompletedWithErrors'){ Add-Action 'Complete migration batch' 'NotNeeded' "Batch is already $($batch.Status)." }
        else {
            if($PSCmdlet.ShouldProcess($BatchName,'Complete migration batch and begin cutover routing changes')){
                Complete-MigrationBatch -Identity $BatchName -ErrorAction Stop | Out-Null
                Add-Action 'Complete migration batch' 'Applied' 'Completion requested. Monitor until status becomes Completed, then assign Exchange licences.'
            }
        }
    } catch { Add-Action 'Complete migration batch' 'Failed' $_.Exception.Message }
}
function Invoke-Report {
    $parent=Split-Path $script:RunFolder -Parent
    $latest=Get-ChildItem -LiteralPath $parent -Directory -ErrorAction SilentlyContinue | Where-Object { $_.FullName -ne $script:RunFolder -and (Test-Path -LiteralPath (Join-Path $_.FullName 'GoogleToM365MigrationReport.json')) } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if(-not $latest){ Add-Finding Critical 'Report' 'No previous run report was found.' 'Run any mode first.'; return }
    $json=Join-Path $latest.FullName 'GoogleToM365MigrationReport.json'
    if(-not(Test-Path -LiteralPath $json)){ Add-Finding Critical 'Report' "No report found in $($latest.FullName)" 'Run any mode first.'; return }
    $data=Get-Content -LiteralPath $json -Raw | ConvertFrom-Json
    Write-Host ''
    Write-Host "=== Migration report from $($data.Completed) ===" -ForegroundColor Cyan
    Write-Host "Mode: $($data.Mode) | Endpoint: $($data.EndpointName) | Batch: $($data.BatchName)"
    Write-Host "CSV: $($data.CsvPath)"
    Write-Host "Key: $($data.ServiceAccountKeyPath)"
    Write-Host "Admin: $($data.GoogleAdminEmail) | Routing domain: $($data.TargetDeliveryDomain)"
    Write-Host ''
    Write-Host 'Findings:' -ForegroundColor Cyan
    foreach($f in $data.Findings){ Write-Host "  [$($f.Severity)] $($f.Area) - $($f.Message)" }
    Write-Host ''
    Write-Host 'Actions:' -ForegroundColor Cyan
    foreach($a in $data.Actions){ Write-Host "  $($a.Time) | $($a.Action) | $($a.Result) | $($a.Detail)" }
}

# --- Report ---
function Write-Report {
    $report=[pscustomobject]@{ScriptVersion='2.0.0';Completed=(Get-Date).ToString('s');Mode=$Mode;CsvPath=$script:ResolvedCsv;ServiceAccountKeyPath=$script:ResolvedKey;GoogleAdminEmail=$script:ResolvedAdmin;TargetDeliveryDomain=$script:ResolvedDomain;EndpointName=$EndpointName;BatchName=$BatchName;Findings=$script:Findings;Actions=$script:Actions}
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $script:RunFolder 'GoogleToM365MigrationReport.json') -Encoding UTF8
    $script:Findings | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'GoogleToM365MigrationFindings.csv') -NoTypeInformation -Encoding UTF8
    Write-Status OK "Report written to $script:RunFolder"
}

# --- Main ---
function Invoke-Main {
    $script:RunFolder=New-RunFolder
    if(-not $Mode){
        if($NonInteractive){ $Mode='Preflight' } else { $Mode=Show-ModeMenu }
    }
    Write-Status INFO "Google to M365 Migration - $Mode mode"
    switch($Mode){
        'Preflight' { Invoke-Preflight }
        'CreateEndpoint' { Invoke-CreateEndpoint }
        'CreateBatch' { Invoke-CreateBatch }
        'StartBatch' { Invoke-StartBatch }
        'Monitor' { Invoke-Monitor }
        'CompleteBatch' { Invoke-CompleteBatch }
        'Report' { Invoke-Report }
    }
    Write-Report
    if(@($script:Findings | Where-Object { $_.Severity -in 'Critical','High' }).Count -gt 0){ Write-Status FAIL 'Blocking findings exist. No further migration action should be taken.' }
}
try { Invoke-Main } catch { Write-Error "Unexpected failure at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"; if($script:RunFolder){Write-Report} }