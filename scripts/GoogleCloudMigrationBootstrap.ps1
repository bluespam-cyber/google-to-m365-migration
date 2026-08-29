#Requires -Version 5.1
<#
.SYNOPSIS
    Safe Google Cloud preparation helper for a Google Workspace to M365 migration.

.DESCRIPTION
    Inspect is read-only and is the default. This script never creates a project,
    never runs browser authentication automatically, and never creates a private
    key unless CreateServiceAccountKey is explicitly selected with -ApproveKeyCreation.
    Use the companion GoogleToM365Migration.ps1 to create the Exchange endpoint.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
param(
    [ValidateSet('Inspect','InstallSdk','EnableApis','CreateServiceAccount','CreateServiceAccountKey','ListKeys','DisableKey')]
    [string]$Mode = 'Inspect',
    [string]$ProjectId,
    [ValidatePattern('^[a-z][a-z0-9-]{4,28}[a-z0-9]$')][string]$ServiceAccountName = 'm365-migration',
    [string]$ServiceAccountEmail,
    [string]$KeyOutputPath,
    [string]$KeyId,
    [switch]$ApproveApiEnablement,
    [switch]$ApproveKeyCreation,
    [switch]$ApproveKeyDisable,
    [string]$OutputRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Findings = [System.Collections.Generic.List[object]]::new()
$script:Actions = [System.Collections.Generic.List[object]]::new()

function Write-Status { param([ValidateSet('INFO','OK','WARN','FAIL')][string]$Level,[string]$Message) Write-Host "[$Level] $Message" -ForegroundColor (@{INFO='Cyan';OK='Green';WARN='Yellow';FAIL='Red'}[$Level]) }
function Add-Finding { param([ValidateSet('Critical','High','Medium','Low','Info')][string]$Severity,[string]$Area,[string]$Message,[string]$Resolution='')
    $script:Findings.Add([pscustomobject]@{Severity=$Severity;Area=$Area;Message=$Message;Resolution=$Resolution}); $level=if($Severity -in 'Critical','High'){'FAIL'}elseif($Severity -in 'Medium','Low'){'WARN'}else{'INFO'}; Write-Status $level "$Area - $Message"
}
function Add-Action { param([string]$Name,[ValidateSet('Applied','Skipped','Failed','NotNeeded')][string]$Result,[string]$Detail='')
    $script:Actions.Add([pscustomobject]@{Time=(Get-Date).ToString('s');Action=$Name;Result=$Result;Detail=$Detail}); $level=if($Result -eq 'Applied'){'OK'}elseif($Result -eq 'Failed'){'FAIL'}elseif($Result -eq 'Skipped'){'WARN'}else{'INFO'}; Write-Status $level "${Name}: $Result. $Detail"
}
function New-RunFolder {
    $documents=[Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments); if([string]::IsNullOrWhiteSpace($documents) -or -not(Test-Path -LiteralPath $documents)){$documents=$env:USERPROFILE}; $base=if($OutputRoot){$OutputRoot}else{Join-Path $documents 'GoogleCloudMigrationBootstrap'}; $run=Join-Path $base (Get-Date -Format 'yyyyMMdd_HHmmss'); New-Item -ItemType Directory -Path $run -Force | Out-Null; $run
}
function Get-Gcloud {
    $command=Get-Command gcloud.exe -ErrorAction SilentlyContinue
    if(-not $command){ Add-Finding Critical 'Google Cloud CLI' 'gcloud was not found.' 'Install the signed Google Cloud CLI, or run this script with -Mode InstallSdk.'; return $null }
    $script:GcloudPath=if($command.Path){$command.Path}else{$command.Source}
    $script:GcloudPath
}
function Invoke-Gcloud {
    param([string[]]$Arguments)
    if(-not $script:GcloudPath){ Get-Gcloud | Out-Null }
    if(-not $script:GcloudPath){ throw 'Google Cloud CLI is unavailable.' }
    $output=& $script:GcloudPath @Arguments 2>&1
    if($LASTEXITCODE -ne 0){ throw (($output | Out-String).Trim()) }
    $output
}
function Test-ProjectInput {
    if(-not $ProjectId){ Add-Finding Critical 'GCP project' 'ProjectId is required for this mode.' 'Use an existing, approved GCP project; this tool never creates one.'; return $false }
    if($ProjectId -notmatch '^[a-z][a-z0-9-]{4,28}[a-z0-9]$'){ Add-Finding Critical 'GCP project' 'ProjectId has an invalid format.' 'Provide the existing GCP project ID.'; return $false }
    $true
}
function Get-SecureKeyPath {
    if($KeyOutputPath){ return [IO.Path]::GetFullPath($KeyOutputPath) }
    $isWindows = $env:OS -eq 'Windows_NT'
    $base = if($isWindows -and $env:LOCALAPPDATA){$env:LOCALAPPDATA}else{Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)) '.m365-migration'}
    Join-Path $base ("M365Migration{0}Secrets{0}gws-migration-{1}.json" -f [IO.Path]::DirectorySeparatorChar,(Get-Date -Format 'yyyyMMddHHmmss'))
}
function Protect-KeyFile {
    param([string]$Path)
    try {
        if($env:OS -ne 'Windows_NT'){
            $chmod=Get-Command chmod -ErrorAction Stop
            & $chmod.Source 600 $Path
            if($LASTEXITCODE -ne 0){throw 'chmod 600 failed.'}
            Add-Action 'Protect service-account key file' 'Applied' 'POSIX permissions set to owner read/write only.'
            return
        }
        $acl=Get-Acl -LiteralPath $Path
        $acl.SetAccessRuleProtection($true,$false)
        foreach($rule in @($acl.Access)){ $acl.RemoveAccessRule($rule) | Out-Null }
        $user=[Security.Principal.WindowsIdentity]::GetCurrent().Name
        $access=New-Object Security.AccessControl.FileSystemAccessRule($user,'FullControl','Allow')
        $acl.AddAccessRule($access)
        Set-Acl -LiteralPath $Path -AclObject $acl
        Add-Action 'Protect service-account key file' 'Applied' 'Inheritance removed; access granted only to the current Windows user.'
    } catch { Add-Finding High 'Service-account key' 'The key was created but its file ACL could not be restricted.' $_.Exception.Message }
}
function Write-DomainWideDelegationInstructions {
    param([string]$Email)
    if(-not $ProjectId -or -not $Email){return}
    try {
        $details=(Invoke-Gcloud @('iam','service-accounts','describe',$Email,"--project=$ProjectId",'--format=json') | Out-String | ConvertFrom-Json)
        if($details.uniqueId){
            Write-Status WARN 'Complete this manual Google Workspace domain-wide-delegation step before creating the Exchange endpoint:'
            Write-Host "  Numeric Client ID: $($details.uniqueId)" -ForegroundColor Yellow
            Write-Host '  Scopes: https://mail.google.com/,https://www.googleapis.com/auth/calendar,https://www.googleapis.com/auth/contacts,https://www.googleapis.com/auth/admin.directory.user.readonly' -ForegroundColor Yellow
            Add-Action 'Domain-wide delegation handoff' 'Applied' "Numeric Client ID $($details.uniqueId) was displayed; configure these scopes in Google Workspace Admin Console."
        }
    } catch { Add-Finding Medium 'Domain-wide delegation' 'Could not retrieve the service account numeric client ID.' $_.Exception.Message }
}
function Invoke-Inspect {
    $gcloud=Get-Gcloud
    if(-not $gcloud){ return }
    try { $version=(Invoke-Gcloud @('--version')) -join '; '; Add-Action 'Inspect Google Cloud CLI' 'Applied' $version } catch { Add-Action 'Inspect Google Cloud CLI' 'Failed' $_.Exception.Message }
    try { $account=(Invoke-Gcloud @('auth','list','--filter=status:ACTIVE','--format=value(account)') | Select-Object -First 1); if($account){ Add-Action 'Inspect active Google account' 'Applied' $account }else{ Add-Finding High 'Google authentication' 'No active gcloud account was detected.' 'Sign in manually with the approved Google migration administrator, then rerun Inspect.' } } catch { Add-Action 'Inspect active Google account' 'Failed' $_.Exception.Message }
    if($ProjectId){
        try { $project=(Invoke-Gcloud @('projects','describe',$ProjectId,'--format=value(projectNumber,name,lifecycleState)')) -join ' | '; Add-Action 'Inspect GCP project' 'Applied' $project } catch { Add-Finding Critical 'GCP project' "Project $ProjectId is inaccessible." $_.Exception.Message }
        Write-DomainWideDelegationInstructions (Get-ServiceAccountAddress)
    }
}
function Invoke-InstallSdk {
    if(-not $PSCmdlet.ShouldProcess('Google Cloud CLI','Open the official installer download page; installation remains user-controlled')){ return }
    Start-Process 'https://cloud.google.com/sdk/docs/install-sdk'
    Add-Action 'Google Cloud CLI installation' 'Skipped' 'Opened official installation instructions. Automatic binary download is intentionally not used.'
}
function Invoke-EnableApis {
    if(-not (Test-ProjectInput) -or -not $ApproveApiEnablement){ Add-Finding Critical 'Approval' 'API enablement requires -ApproveApiEnablement and an existing ProjectId.' 'Review billing, IAM, and organisational policy before enabling APIs.'; return }
    $apis=@('gmail.googleapis.com','calendar-json.googleapis.com','people.googleapis.com','admin.googleapis.com')
    if($PSCmdlet.ShouldProcess($ProjectId,"Enable required Google Workspace migration APIs: $($apis -join ', ')")){
        try { Invoke-Gcloud (@('services','enable') + $apis + @("--project=$ProjectId",'--quiet')) | Out-Null; Add-Action 'Enable Google APIs' 'Applied' ($apis -join ', ') } catch { Add-Action 'Enable Google APIs' 'Failed' $_.Exception.Message }
    }
}
function Get-ServiceAccountAddress {
    if($ServiceAccountEmail){ return $ServiceAccountEmail }
    "$ServiceAccountName@$ProjectId.iam.gserviceaccount.com"
}
function Invoke-CreateServiceAccount {
    if(-not (Test-ProjectInput)){ return }
    $email=Get-ServiceAccountAddress
    try { Invoke-Gcloud @('iam','service-accounts','describe',$email,"--project=$ProjectId",'--format=value(email)') | Out-Null; Add-Action 'Create service account' 'NotNeeded' "$email already exists."; return } catch { }
    if($PSCmdlet.ShouldProcess($email,"Create service account in existing project $ProjectId")){
        try { Invoke-Gcloud @('iam','service-accounts','create',$ServiceAccountName,"--project=$ProjectId",'--display-name=Microsoft 365 Google Workspace Migration','--description=Dedicated service account for Exchange Online Google Workspace migration','--quiet') | Out-Null; Add-Action 'Create service account' 'Applied' $email; Write-DomainWideDelegationInstructions $email } catch { Add-Action 'Create service account' 'Failed' $_.Exception.Message }
    }
}
function Invoke-CreateKey {
    if(-not (Test-ProjectInput) -or -not $ApproveKeyCreation){ Add-Finding Critical 'Approval' 'Key creation requires -ApproveKeyCreation and an existing ProjectId.' 'Prefer a centrally approved key-management process; only create a key when Microsoft migration requires it.'; return }
    $email=Get-ServiceAccountAddress; $path=Get-SecureKeyPath
    if(Test-Path -LiteralPath $path){ Add-Finding Critical 'Service-account key' "Key output already exists: $path" 'Choose a new empty path; this tool never overwrites a key.'; return }
    New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
    if($PSCmdlet.ShouldProcess($email,"Create a JSON private key at $path")){
        try { Invoke-Gcloud @('iam','service-accounts','keys','create',$path,"--iam-account=$email","--project=$ProjectId",'--key-file-type=json','--quiet') | Out-Null; Protect-KeyFile $path; Add-Action 'Create service-account key' 'Applied' "Key stored at $path. Supply this exact path to GoogleToM365Migration.ps1, then rotate/remove it under the organisation’s key policy." } catch { if(Test-Path -LiteralPath $path){ Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }; Add-Action 'Create service-account key' 'Failed' $_.Exception.Message }
    }
}
function Invoke-ListKeys {
    if(-not (Test-ProjectInput)){ return }; $email=Get-ServiceAccountAddress
    try { Invoke-Gcloud @('iam','service-accounts','keys','list',"--iam-account=$email","--project=$ProjectId",'--managed-by=user','--format=table(name.basename(),validAfterTime,validBeforeTime,keyType)') | ForEach-Object { Write-Host $_ }; Add-Action 'List service-account keys' 'Applied' $email } catch { Add-Action 'List service-account keys' 'Failed' $_.Exception.Message }
}
function Invoke-DisableKey {
    if(-not (Test-ProjectInput) -or -not $KeyId -or -not $ApproveKeyDisable){ Add-Finding Critical 'Approval' 'Disabling a key requires ProjectId, KeyId, and -ApproveKeyDisable.' 'Confirm no migration process still requires the key before disabling it.'; return }
    $email=Get-ServiceAccountAddress
    if($PSCmdlet.ShouldProcess($KeyId,"Disable service-account key for $email")){ try { Invoke-Gcloud @('iam','service-accounts','keys','disable',$KeyId,"--iam-account=$email","--project=$ProjectId",'--quiet') | Out-Null; Add-Action 'Disable service-account key' 'Applied' $KeyId } catch { Add-Action 'Disable service-account key' 'Failed' $_.Exception.Message } }
}
function Write-Report {
    $reportServiceAccount = if($ProjectId){Get-ServiceAccountAddress}else{$ServiceAccountEmail}
    $report=[pscustomobject]@{ScriptVersion='1.0.0';Completed=Get-Date;Mode=$Mode;ProjectId=$ProjectId;ServiceAccount=$reportServiceAccount;Findings=$script:Findings;Actions=$script:Actions}
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $script:RunFolder 'GoogleCloudBootstrapReport.json') -Encoding UTF8
    $script:Findings | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'GoogleCloudBootstrapFindings.csv') -NoTypeInformation -Encoding UTF8
    Write-Status OK "Report written to $script:RunFolder"
}
function Invoke-Main {
    $script:RunFolder=New-RunFolder; Write-Status INFO "Google Cloud Migration Bootstrap - $Mode mode"
    switch($Mode){
        'Inspect' { Invoke-Inspect }
        'InstallSdk' { Invoke-InstallSdk }
        'EnableApis' { Invoke-EnableApis }
        'CreateServiceAccount' { Invoke-CreateServiceAccount }
        'CreateServiceAccountKey' { Invoke-CreateKey }
        'ListKeys' { Invoke-ListKeys }
        'DisableKey' { Invoke-DisableKey }
    }
    Write-Report
    if(@($script:Findings | Where-Object { $_.Severity -in 'Critical','High' }).Count -gt 0){ Write-Status FAIL 'Blocking findings exist. No further cloud action should be taken.' }
}
try { Invoke-Main } catch { Write-Error "Unexpected failure at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"; if($script:RunFolder){Write-Report} }
