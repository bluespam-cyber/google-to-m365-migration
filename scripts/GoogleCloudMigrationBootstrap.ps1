#Requires -Version 5.1
<#
.SYNOPSIS
    Interactive Google Cloud preparation helper for a Google Workspace to M365 migration.

.DESCRIPTION
    Fully interactive and self-detecting. Auto-detects gcloud, your active Google
    account, GCP projects, and existing service accounts/keys. Prompts only for
    values it cannot detect on its own. Read-only by default; never creates a
    project or key without explicit approval.

    Run with no arguments for a guided menu. Use -NonInteractive for RMM/SYSTEM
    automation (fails fast instead of prompting).

.EXAMPLE
    .\Run-Bootstrap.ps1

.EXAMPLE
    .\Run-Bootstrap.ps1 -Mode CreateServiceAccountKey

.EXAMPLE
    .\Run-Bootstrap.ps1 -Mode Inspect -NonInteractive
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
param(
    [ValidateSet('Inspect','InstallSdk','EnableApis','CreateServiceAccount','CreateServiceAccountKey','ListKeys','DisableKey')]
    [string]$Mode,
    [string]$ProjectId,
    [ValidatePattern('^[a-z][a-z0-9-]{4,28}[a-z0-9]$')][string]$ServiceAccountName = 'm365-migration',
    [string]$ServiceAccountEmail,
    [string]$KeyOutputPath,
    [string]$KeyId,
    [switch]$ApproveApiEnablement,
    [switch]$ApproveKeyCreation,
    [switch]$ApproveKeyDisable,
    [switch]$NonInteractive,
    [string]$OutputRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Findings = [System.Collections.Generic.List[object]]::new()
$script:Actions = [System.Collections.Generic.List[object]]::new()
$script:GcloudPath = $null
$script:GcloudMissing = $false
$script:RunFolder = $null
$script:ResolvedProjectId = $null

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
    $base=if($OutputRoot){$OutputRoot}else{Join-Path $documents 'GoogleCloudMigrationBootstrap'}
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
    $modes=@('Inspect','InstallSdk','EnableApis','CreateServiceAccount','CreateServiceAccountKey','ListKeys','DisableKey')
    $desc=@{
        'Inspect'='Inspect gcloud, account, project, service account (read-only)'
        'InstallSdk'='Install the Google Cloud SDK (opens official installer)'
        'EnableApis'='Enable Gmail, Calendar, People, Admin APIs'
        'CreateServiceAccount'='Create the migration service account'
        'CreateServiceAccountKey'='Create a JSON private key for the service account'
        'ListKeys'='List user-managed service-account keys'
        'DisableKey'='Disable a service-account key'
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

# --- gcloud discovery ---
function Get-Gcloud {
    if($script:GcloudPath){ return $script:GcloudPath }
    $command=Get-Command gcloud.exe -ErrorAction SilentlyContinue
    if(-not $command){
        if(-not $script:GcloudMissing){ $script:GcloudMissing=$true; Add-Finding Critical 'Google Cloud CLI' 'gcloud was not found.' 'Install the signed Google Cloud CLI, or run this script with -Mode InstallSdk.' }
        return $null
    }
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
function Get-ActiveAccount {
    if(-not $script:GcloudPath){ return $null }
    try {
        $acct=& $script:GcloudPath auth list --filter=status:ACTIVE --format=value(account) 2>$null | Select-Object -First 1
        if($acct){ return ([string]$acct).Trim() }
    } catch { }
    return $null
}
function Get-ProjectList {
    if(-not $script:GcloudPath){ return @() }
    try {
        $projects=@(& $script:GcloudPath projects list --format=value(projectId) 2>$null | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
        return $projects
    } catch { return @() }
}
function Resolve-ProjectId {
    if($ProjectId){ $script:ResolvedProjectId=$ProjectId; return $ProjectId }
    $projects=@(Get-ProjectList)
    if($projects.Count -eq 1){
        Write-Status INFO "Auto-detected GCP project: $($projects[0])"
        $script:ResolvedProjectId=$projects[0]
        return $projects[0]
    }
    if($projects.Count -gt 1){
        Write-Host 'Detected GCP projects:' -ForegroundColor Cyan
        for($i=0;$i -lt $projects.Count;$i++){ Write-Host "  $($i+1)) $($projects[$i])" }
        $choice=Read-Value 'Select project number' '1'
        $idx=0
        if([int]::TryParse($choice,[ref]$idx) -and $idx -ge 1 -and $idx -le $projects.Count){ $script:ResolvedProjectId=$projects[$idx-1]; return $projects[$idx-1] }
        $script:ResolvedProjectId=Read-Value 'Enter the GCP project ID'
        return $script:ResolvedProjectId
    }
    $script:ResolvedProjectId=Read-Value 'Enter your existing GCP project ID (required)'
    $script:ResolvedProjectId
}
function Get-ServiceAccountAddress {
    if($ServiceAccountEmail){ return $ServiceAccountEmail }
    $proj=if($script:ResolvedProjectId){$script:ResolvedProjectId}else{$ProjectId}
    "$ServiceAccountName@$proj.iam.gserviceaccount.com"
}
function Resolve-ServiceAccountEmail {
    if($ServiceAccountEmail){ return $ServiceAccountEmail }
    if(-not $ProjectId){ return $null }
    "$ServiceAccountName@$ProjectId.iam.gserviceaccount.com"
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

# --- Modes ---
function Invoke-Inspect {
    $gcloud=Get-Gcloud
    if(-not $gcloud){ return }
    try { $version=(Invoke-Gcloud @('--version')) -join '; '; Add-Action 'Inspect Google Cloud CLI' 'Applied' $version } catch { Add-Action 'Inspect Google Cloud CLI' 'Failed' $_.Exception.Message }
    $account=Get-ActiveAccount
    if($account){ Add-Action 'Inspect active Google account' 'Applied' $account }else{ Add-Finding High 'Google authentication' 'No active gcloud account was detected.' 'Sign in manually with the approved Google migration administrator, then rerun Inspect.' }
    if(-not $ProjectId){
        $detected=@(Get-ProjectList)
        if($detected.Count -eq 1){ $script:DetectedProject=$detected[0]; Write-Status INFO "Auto-detected project: $($detected[0])" }
        elseif($detected.Count -gt 1){ Write-Status INFO "Detected $($detected.Count) projects; use -ProjectId to select one." }
    }
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
    if(-not $ProjectId){ $ProjectId=Resolve-ProjectId }
    if(-not $ProjectId){ Add-Finding Critical 'GCP project' 'A GCP project is required to enable APIs.' 'Provide an existing, approved GCP project.'; return }
    if(-not $ApproveApiEnablement){
        if($NonInteractive){ Add-Finding Critical 'Approval' 'API enablement requires -ApproveApiEnablement.' 'Review billing, IAM, and organisational policy before enabling APIs.'; return }
        if(-not (Confirm-YesNo "Enable required Google Workspace migration APIs in project '$ProjectId'?")){ Add-Action 'Enable Google APIs' 'Skipped' 'Declined by user.'; return }
    }
    $apis=@('gmail.googleapis.com','calendar-json.googleapis.com','people.googleapis.com','admin.googleapis.com')
    if($PSCmdlet.ShouldProcess($ProjectId,"Enable required Google Workspace migration APIs: $($apis -join ', ')")){
        try { Invoke-Gcloud (@('services','enable') + $apis + @("--project=$ProjectId",'--quiet')) | Out-Null; Add-Action 'Enable Google APIs' 'Applied' ($apis -join ', ') } catch { Add-Action 'Enable Google APIs' 'Failed' $_.Exception.Message }
    }
}
function Invoke-CreateServiceAccount {
    if(-not $ProjectId){ $ProjectId=Resolve-ProjectId }
    if(-not $ProjectId){ Add-Finding Critical 'GCP project' 'A GCP project is required to create a service account.' 'Provide an existing, approved GCP project.'; return }
    $email=Get-ServiceAccountAddress
    try { Invoke-Gcloud @('iam','service-accounts','describe',$email,"--project=$ProjectId",'--format=value(email)') | Out-Null; Add-Action 'Create service account' 'NotNeeded' "$email already exists."; Write-DomainWideDelegationInstructions $email; return } catch { }
    if($PSCmdlet.ShouldProcess($email,"Create service account in existing project $ProjectId")){
        try { Invoke-Gcloud @('iam','service-accounts','create',$ServiceAccountName,"--project=$ProjectId",'--display-name=Microsoft 365 Google Workspace Migration','--description=Dedicated service account for Exchange Online Google Workspace migration','--quiet') | Out-Null; Add-Action 'Create service account' 'Applied' $email; Write-DomainWideDelegationInstructions $email } catch { Add-Action 'Create service account' 'Failed' $_.Exception.Message }
    }
}
function Invoke-CreateKey {
    if(-not $ProjectId){ $ProjectId=Resolve-ProjectId }
    if(-not $ProjectId){ Add-Finding Critical 'GCP project' 'A GCP project is required to create a key.' 'Provide an existing, approved GCP project.'; return }
    if(-not $ApproveKeyCreation){
        if($NonInteractive){ Add-Finding Critical 'Approval' 'Key creation requires -ApproveKeyCreation.' 'Prefer a centrally approved key-management process; only create a key when Microsoft migration requires it.'; return }
        if(-not (Confirm-YesNo 'Create a JSON private key for the migration service account?')){ Add-Action 'Create service-account key' 'Skipped' 'Declined by user.'; return }
    }
    $email=Get-ServiceAccountAddress; $path=Get-SecureKeyPath
    if(Test-Path -LiteralPath $path){ Add-Finding Critical 'Service-account key' "Key output already exists: $path" 'Choose a new empty path; this tool never overwrites a key.'; return }
    New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
    if($PSCmdlet.ShouldProcess($email,"Create a JSON private key at $path")){
        try { Invoke-Gcloud @('iam','service-accounts','keys','create',$path,"--iam-account=$email","--project=$ProjectId",'--key-file-type=json','--quiet') | Out-Null; Protect-KeyFile $path; Add-Action 'Create service-account key' 'Applied' "Key stored at $path. Supply this exact path to GoogleToM365Migration.ps1, then rotate/remove it under the organisation's key policy." } catch { if(Test-Path -LiteralPath $path){ Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }; Add-Action 'Create service-account key' 'Failed' $_.Exception.Message }
    }
}
function Invoke-ListKeys {
    if(-not $ProjectId){ $ProjectId=Resolve-ProjectId }
    if(-not $ProjectId){ Add-Finding Critical 'GCP project' 'A GCP project is required to list keys.' 'Provide an existing, approved GCP project.'; return }
    $email=Get-ServiceAccountAddress
    try { Invoke-Gcloud @('iam','service-accounts','keys','list',"--iam-account=$email","--project=$ProjectId",'--managed-by=user','--format=table(name.basename(),validAfterTime,validBeforeTime,keyType)') | ForEach-Object { Write-Host $_ }; Add-Action 'List service-account keys' 'Applied' $email } catch { Add-Action 'List service-account keys' 'Failed' $_.Exception.Message }
}
function Invoke-DisableKey {
    if(-not $ProjectId){ $ProjectId=Resolve-ProjectId }
    if(-not $ProjectId){ Add-Finding Critical 'GCP project' 'A GCP project is required to disable a key.' 'Provide an existing, approved GCP project.'; return }
    if(-not $KeyId){
        if($NonInteractive){ Add-Finding Critical 'Approval' 'Disabling a key requires KeyId and -ApproveKeyDisable.' 'Confirm no migration process still requires the key before disabling it.'; return }
        $KeyId=Read-Value 'Enter the full key ID (or key name) to disable'
    }
    if(-not $ApproveKeyDisable){
        if($NonInteractive){ Add-Finding Critical 'Approval' 'Disabling a key requires -ApproveKeyDisable.' 'Confirm no migration process still requires the key before disabling it.'; return }
        if(-not (Confirm-YesNo "Disable service-account key '$KeyId'?")){ Add-Action 'Disable service-account key' 'Skipped' 'Declined by user.'; return }
    }
    $email=Get-ServiceAccountAddress
    if($PSCmdlet.ShouldProcess($KeyId,"Disable service-account key for $email")){ try { Invoke-Gcloud @('iam','service-accounts','keys','disable',$KeyId,"--iam-account=$email","--project=$ProjectId",'--quiet') | Out-Null; Add-Action 'Disable service-account key' 'Applied' $KeyId } catch { Add-Action 'Disable service-account key' 'Failed' $_.Exception.Message } }
}

# --- Report ---
function Write-Report {
    $reportProjectId=if($script:ResolvedProjectId){$script:ResolvedProjectId}else{$ProjectId}
    $reportServiceAccount = if($reportProjectId){Get-ServiceAccountAddress}else{$ServiceAccountEmail}
    $report=[pscustomobject]@{ScriptVersion='2.0.0';Completed=(Get-Date).ToString('s');Mode=$Mode;ProjectId=$reportProjectId;ServiceAccount=$reportServiceAccount;Findings=$script:Findings;Actions=$script:Actions}
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $script:RunFolder 'GoogleCloudBootstrapReport.json') -Encoding UTF8
    $script:Findings | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'GoogleCloudBootstrapFindings.csv') -NoTypeInformation -Encoding UTF8
    Write-Status OK "Report written to $script:RunFolder"
}

# --- Main ---
function Invoke-Main {
    $script:RunFolder=New-RunFolder
    if(-not $Mode){
        if($NonInteractive){ $Mode='Inspect' } else { $Mode=Show-ModeMenu }
    }
    Write-Status INFO "Google Cloud Migration Bootstrap - $Mode mode"
    Get-Gcloud | Out-Null
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