#Requires -Version 5.1
<#
.SYNOPSIS
    Google Workspace to Microsoft 365 mailbox migration controller.
    No Google Cloud SDK. No gcloud. No Google PowerShell modules.

.DESCRIPTION
    The Google side (project, APIs, service account, key, domain-wide delegation)
    is done ONCE by hand in the Google consoles - it is a five minute job and the
    script prints the exact click-path and the exact scope string to paste.

    This script then does something the SDK-based approach never did: it PROVES
    the Google side is correct before a single mailbox is touched. It signs a JWT
    with the service account key and asks Google for a delegated access token for
    a real user. If domain-wide delegation is missing, the scopes are wrong, or
    the key is disabled, Google says so - and the exact reason is decoded into
    plain English instead of surfacing later as a failed migration batch.

    Design rules:
      * NOTHING destructive without an explicit approval switch.
      * The batch is NEVER auto-started. You review, then start.
      * Every blocking finding stops the run. A preflight that checks nothing
        never reports success.
      * The private key is shredded from disk in a finally block.
      * Everything is logged to a transcript and a findings CSV.

.PARAMETER Mode
    Guide       Print the Google-side setup instructions and exit. No connections.
    Preflight   Verify everything. Changes nothing. THE DEFAULT.
    CreateEndpoint  Create the Exchange migration endpoint.
    CreateBatch     Create the batch (never auto-started).
    StartBatch      Start a created batch.
    Monitor         Watch progress.
    Complete        Finalise the batch (cutover). Requires -ApproveCutover.
    Run             Preflight, then endpoint, then batch - stopping on any blocker.

.EXAMPLE
    .\Invoke-GoogleToM365Migration.ps1 -Mode Guide

.EXAMPLE
    .\Invoke-GoogleToM365Migration.ps1 -Mode Preflight `
        -KeyPath .\key.json -GoogleAdminEmail admin@contoso.com `
        -CsvPath .\users.csv -TargetDeliveryDomain o365.contoso.com

.NOTES
    Author  : Arwaz Khan
    Requires: Windows PowerShell 5.1+, ExchangeOnlineManagement 3.0.0+
    Network : login.microsoftonline.com, outlook.office365.com, oauth2.googleapis.com
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [ValidateSet('Guide','Preflight','CreateEndpoint','CreateBatch','StartBatch','Monitor','Complete','Run')]
    [string]$Mode = 'Preflight',

    [string]$KeyPath,

    [string]$GoogleAdminEmail,

    [string]$CsvPath,

    [string]$TargetDeliveryDomain,

    [ValidateNotNullOrEmpty()]
    [string]$EndpointName = 'GoogleWorkspaceEndpoint',

    [ValidateNotNullOrEmpty()]
    [string]$BatchName = 'GoogleWorkspaceMigration',

    [ValidateRange(15,3600)]
    [int]$MonitorIntervalSeconds = 60,

    [ValidateRange(1,1440)]
    [int]$MaxMonitorMinutes = 60,

    [string]$OutputRoot,

    # Number of users to test delegation against. 0 tests every user in the CSV.
    [ValidateRange(0,500)]
    [int]$DelegationTestCount = 3,

    # Required before the batch is finalised - cutover redirects live mail flow.
    [switch]$ApproveCutover,

    # Keep the service account key on disk after the run. Not recommended.
    [switch]$KeepKeyFile,

    [switch]$SkipGoogleTest,

    [switch]$NonInteractive
)

# ----------------------------------------------------------------------------------
# STATE
# ----------------------------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$script:Version    = '3.0.0'
$script:Findings   = New-Object System.Collections.ArrayList
$script:Actions    = New-Object System.Collections.ArrayList
$script:StartTime  = Get-Date
$script:RunFolder  = $null
$script:Connected  = $false
$script:KeyBytes   = $null
$script:ValidatedCsv = $null
$script:Mappings   = @()
$script:GoogleRouting      = ''
$script:WaitForDelegation  = $false
$script:AutoCreateMailUsers = $false
$script:MissingRecipients  = @()
# Captured at script scope: $PSBoundParameters inside a function resolves through
# the scope chain, which is fragile. Decide interactivity from the real value.
$script:ModeWasSpecified   = $PSBoundParameters.ContainsKey('Mode')
$script:ResolvedTargetDomain = ''
# Preflight is a read-only contract. Only these modes may change the tenant.
$script:AllowMutation      = $false

# The scope string Microsoft requires for Google Workspace migration, verbatim.
# Source: learn.microsoft.com/exchange/mailbox-migration/manually-configuring-gsuite-for-migration
# These must be authorised EXACTLY - an extra or missing scope breaks the match
# and the migration fails only AFTER the batch starts.
$script:RequiredScopes = @(
    'https://mail.google.com/'
    'https://www.googleapis.com/auth/calendar'
    'https://www.google.com/m8/feeds/'
    'https://www.googleapis.com/auth/gmail.settings.sharing'
    'https://www.googleapis.com/auth/contacts'
)
$script:ScopeString = ($script:RequiredScopes -join ',')

# APIs that must be enabled on the Google Cloud project.
$script:RequiredApis = @(
    @{ Name = 'Gmail API';           Id = 'gmail.googleapis.com' }
    @{ Name = 'Google Calendar API'; Id = 'calendar-json.googleapis.com' }
    @{ Name = 'Contacts API';        Id = 'contacts.googleapis.com' }
    @{ Name = 'People API';          Id = 'people.googleapis.com' }
)

# ----------------------------------------------------------------------------------
# OUTPUT HELPERS
# ----------------------------------------------------------------------------------
function Write-Head {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 76) -ForegroundColor DarkCyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ('=' * 76) -ForegroundColor DarkCyan
}
function Write-Step { param([string]$T) Write-Host "  -> $T" -ForegroundColor Gray }
function Write-Good { param([string]$T) Write-Host "  [OK]    $T" -ForegroundColor Green }
function Write-Warn2 { param([string]$T) Write-Host "  [WARN]  $T" -ForegroundColor Yellow }
function Write-Bad  { param([string]$T) Write-Host "  [FAIL]  $T" -ForegroundColor Red }
function Write-Info { param([string]$T) Write-Host "  [INFO]  $T" -ForegroundColor White }

function Add-Finding {
    param(
        [ValidateSet('Critical','High','Medium','Low','Info')][string]$Severity,
        [string]$Area,
        [string]$Title,
        [string]$Evidence = '',
        [string]$Remedy = ''
    )
    $null = $script:Findings.Add([PSCustomObject]@{
        Severity = $Severity; Area = $Area; Title = $Title
        Evidence = $Evidence; Remedy = $Remedy
    })
    switch ($Severity) {
        'Critical' { Write-Bad $Title;   if ($Remedy) { Write-Host "          FIX: $Remedy" -ForegroundColor Gray } }
        'High'     { Write-Bad $Title;   if ($Remedy) { Write-Host "          FIX: $Remedy" -ForegroundColor Gray } }
        'Medium'   { Write-Warn2 $Title; if ($Remedy) { Write-Host "          FIX: $Remedy" -ForegroundColor Gray } }
        'Low'      { Write-Warn2 $Title }
        default    { Write-Info $Title }
    }
}

function Add-Action {
    param(
        [string]$Name,
        [ValidateSet('Applied','Skipped','Failed','NotNeeded','Verified')][string]$Result,
        [string]$Detail = ''
    )
    $null = $script:Actions.Add([PSCustomObject]@{
        Time = (Get-Date).ToString('HH:mm:ss'); Name = $Name; Result = $Result; Detail = $Detail
    })
    switch ($Result) {
        'Applied'  { Write-Good "$Name - $Detail" }
        'Verified' { Write-Good "$Name - $Detail" }
        'Failed'   { Write-Bad  "$Name - $Detail" }
        'Skipped'  { Write-Warn2 "$Name - $Detail" }
        default    { Write-Step "$Name - not needed" }
    }
}

function Test-Blocking {
    return (@($script:Findings | Where-Object { $_.Severity -in @('Critical','High') }).Count -gt 0)
}

function Assert-NoBlockers {
    param([string]$Because)
    if (Test-Blocking) {
        $n = @($script:Findings | Where-Object { $_.Severity -in @('Critical','High') }).Count
        Write-Host ''
        Write-Bad "$n blocking finding(s). Refusing to $Because."
        Write-Info 'Fix the items above and re-run. Nothing was changed.'
        throw "Blocked: $n unresolved blocking finding(s)."
    }
}

function New-RunFolder {
    if (-not $OutputRoot) {
        $dl = Join-Path $env:USERPROFILE 'Downloads'
        try {
            $guid = '{374DE290-123F-4565-9164-39C4925E467B}'
            $v = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction SilentlyContinue).$guid
            if ($v) { $dl = [Environment]::ExpandEnvironmentVariables($v) }
        } catch { }
        $OutputRoot = Join-Path $dl 'GoogleM365Migration'
    }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $OutputRoot $stamp
    $null = New-Item -Path $path -ItemType Directory -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path $path)) {
        $path = Join-Path $env:TEMP "GoogleM365Migration_$stamp"
        $null = New-Item -Path $path -ItemType Directory -Force -ErrorAction SilentlyContinue
    }
    $script:RunFolder = $path
    return $path
}

# ----------------------------------------------------------------------------------
# GOOGLE AUTH WITHOUT THE SDK
#
#   Everything here is pure .NET and REST. The service account key is a PKCS#8
#   PEM inside the JSON. Windows PowerShell 5.1 has no ImportPkcs8PrivateKey, so
#   the DER is parsed by hand into RSAParameters and the JWT is signed with RS256.
#
#   This is what lets the script PROVE domain-wide delegation works before any
#   mailbox is touched, which is the single largest cause of failed migrations.
# ----------------------------------------------------------------------------------

function Read-DerLength {
    param([byte[]]$Data, [ref]$Pos)
    $first = $Data[$Pos.Value]; $Pos.Value++
    if ($first -lt 0x80) { return [int]$first }
    $count = $first -band 0x7F
    if ($count -gt 4) { throw 'DER length field is too large to be a valid RSA key.' }
    $len = 0
    for ($i = 0; $i -lt $count; $i++) {
        $len = ($len -shl 8) -bor $Data[$Pos.Value]
        $Pos.Value++
    }
    return $len
}

function Read-DerTag {
    param([byte[]]$Data, [ref]$Pos, [byte]$Expected)
    $tag = $Data[$Pos.Value]; $Pos.Value++
    if ($tag -ne $Expected) {
        throw ("Unexpected DER tag 0x{0:X2} at offset {1}, expected 0x{2:X2}." -f $tag, ($Pos.Value - 1), $Expected)
    }
    return Read-DerLength -Data $Data -Pos $Pos
}

function Read-DerInteger {
    param([byte[]]$Data, [ref]$Pos)
    $len = Read-DerTag -Data $Data -Pos $Pos -Expected 0x02
    $bytes = New-Object -TypeName 'byte[]' -ArgumentList $len
    [Array]::Copy($Data, $Pos.Value, $bytes, 0, $len)
    $Pos.Value += $len
    # DER stores a leading 0x00 to keep the value positive - strip it.
    if ($bytes.Length -gt 1 -and $bytes[0] -eq 0) {
        $trimmed = New-Object -TypeName 'byte[]' -ArgumentList ($bytes.Length - 1)
        [Array]::Copy($bytes, 1, $trimmed, 0, $trimmed.Length)
        return $trimmed
    }
    return $bytes
}

function Expand-ToLength {
    param([byte[]]$Bytes, [int]$Length)
    if ($Bytes.Length -eq $Length) { return $Bytes }
    if ($Bytes.Length -gt $Length) {
        $t = New-Object -TypeName 'byte[]' -ArgumentList $Length
        [Array]::Copy($Bytes, $Bytes.Length - $Length, $t, 0, $Length)
        return $t
    }
    $p = New-Object -TypeName 'byte[]' -ArgumentList $Length
    [Array]::Copy($Bytes, 0, $p, $Length - $Bytes.Length, $Bytes.Length)
    return $p
}

function ConvertFrom-PemPrivateKey {
    <#  PKCS#8 PEM -> RSAParameters, without any external library.  #>
    param([Parameter(Mandatory)][string]$Pem)

    $b64 = $Pem -replace '-----BEGIN [A-Z ]+-----', '' `
                -replace '-----END [A-Z ]+-----', '' `
                -replace '\s', ''
    if (-not $b64) { throw 'The private_key field in the JSON key is empty.' }

    $der = [Convert]::FromBase64String($b64)
    $pos = 0

    # PrivateKeyInfo ::= SEQUENCE { version, AlgorithmIdentifier, OCTET STRING }
    $null = Read-DerTag -Data $der -Pos ([ref]$pos) -Expected 0x30
    $null = Read-DerInteger -Data $der -Pos ([ref]$pos)          # version
    $algLen = Read-DerTag -Data $der -Pos ([ref]$pos) -Expected 0x30
    $pos += $algLen                                              # skip AlgorithmIdentifier
    $null = Read-DerTag -Data $der -Pos ([ref]$pos) -Expected 0x04

    # Inner RSAPrivateKey
    $null = Read-DerTag -Data $der -Pos ([ref]$pos) -Expected 0x30
    $null = Read-DerInteger -Data $der -Pos ([ref]$pos)          # version

    $n  = Read-DerInteger -Data $der -Pos ([ref]$pos)
    $e  = Read-DerInteger -Data $der -Pos ([ref]$pos)
    $d  = Read-DerInteger -Data $der -Pos ([ref]$pos)
    $p  = Read-DerInteger -Data $der -Pos ([ref]$pos)
    $q  = Read-DerInteger -Data $der -Pos ([ref]$pos)
    $dp = Read-DerInteger -Data $der -Pos ([ref]$pos)
    $dq = Read-DerInteger -Data $der -Pos ([ref]$pos)
    $iq = Read-DerInteger -Data $der -Pos ([ref]$pos)

    $modLen  = $n.Length
    $halfLen = [int][math]::Ceiling($modLen / 2)

    $params = New-Object System.Security.Cryptography.RSAParameters
    $params.Modulus  = $n
    $params.Exponent = $e
    $params.D        = Expand-ToLength -Bytes $d  -Length $modLen
    $params.P        = Expand-ToLength -Bytes $p  -Length $halfLen
    $params.Q        = Expand-ToLength -Bytes $q  -Length $halfLen
    $params.DP       = Expand-ToLength -Bytes $dp -Length $halfLen
    $params.DQ       = Expand-ToLength -Bytes $dq -Length $halfLen
    $params.InverseQ = Expand-ToLength -Bytes $iq -Length $halfLen
    return $params
}

function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    # Google rejects padding and newlines in the assertion.
    return ([Convert]::ToBase64String($Bytes) -replace '\+', '-' -replace '/', '_' -replace '=', '')
}

function New-SignedJwt {
    <#  RS256-signed JWT for the Google service account flow.  #>
    param(
        [Parameter(Mandatory)]$KeyObject,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string[]]$Scopes
    )

    # Explicit UTC epoch. -UFormat %s is locale- and timezone-fragile, and Google
    # rejects a JWT whose iat/exp are outside the valid window.
    $epoch = New-Object DateTime(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
    $now = [int]([DateTime]::UtcNow - $epoch).TotalSeconds
    $header = @{ alg = 'RS256'; typ = 'JWT'; kid = $KeyObject.private_key_id } | ConvertTo-Json -Compress
    $claims = @{
        iss   = $KeyObject.client_email
        sub   = $Subject
        scope = ($Scopes -join ' ')          # scope claim is SPACE separated
        aud   = 'https://oauth2.googleapis.com/token'
        iat   = $now
        exp   = $now + 3600
    } | ConvertTo-Json -Compress

    $h = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($header))
    $c = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($claims))
    $signingInput = [Text.Encoding]::ASCII.GetBytes("$h.$c")

    $rsaParams = ConvertFrom-PemPrivateKey -Pem $KeyObject.private_key
    $sig = $null

    # RSACng is the clean path on .NET 4.6+. RSACryptoServiceProvider needs the
    # AES provider (type 24) before it will do SHA-256, hence the fallback.
    try {
        $rsa = New-Object System.Security.Cryptography.RSACng
        $rsa.ImportParameters($rsaParams)
        $sig = $rsa.SignData($signingInput,
                [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $rsa.Dispose()
    }
    catch {
        # Fallback for systems where RSACng is unavailable. RSACryptoServiceProvider
        # can only do SHA-256 when it is backed by the AES provider (type 24), and
        # the key must be imported into THAT provider - not imported and then moved.
        try {
            $csp = New-Object System.Security.Cryptography.CspParameters
            $csp.ProviderType = 24
            $csp.Flags = [System.Security.Cryptography.CspProviderFlags]::CreateEphemeralKey
            $rsa2 = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList 2048, $csp
            $rsa2.PersistKeyInCsp = $false
            $rsa2.ImportParameters($rsaParams)
            $sig = $rsa2.SignData($signingInput, 'SHA256')
            $rsa2.Dispose()
        }
        catch {
            # Last resort: default provider, explicit OID.
            $rsa3 = New-Object System.Security.Cryptography.RSACryptoServiceProvider
            $rsa3.ImportParameters($rsaParams)
            $sha = [System.Security.Cryptography.SHA256]::Create()
            $hash = $sha.ComputeHash($signingInput)
            $fmt = New-Object System.Security.Cryptography.RSAPKCS1SignatureFormatter($rsa3)
            $fmt.SetHashAlgorithm('SHA256')
            $sig = $fmt.CreateSignature($hash)
            $sha.Dispose(); $rsa3.Dispose()
        }
    }

    if (-not $sig) { throw 'Could not sign the JWT with the supplied private key.' }
    return "$h.$c." + (ConvertTo-Base64Url $sig)
}

function Resolve-GoogleAuthError {
    <#
        Decodes Google's OAuth error into the actual cause and fix.
        Source: developers.google.com/identity/protocols/oauth2/service-account
    #>
    param([string]$ErrorCode, [string]$Description)

    $d = "$Description"
    switch ($ErrorCode) {
        'unauthorized_client' {
            if ($d -match 'not authorized for any of the scopes|unauthorized to retrieve access tokens') {
                return @{
                    Cause  = 'The service account is not authorised in the Admin console, OR it was added using the service account EMAIL instead of the numeric Client ID.'
                    Remedy = 'Google Admin > Security > Access and data control > API controls > Manage Domain Wide Delegation. Remove the entry and re-add it using the NUMERIC client_id from the JSON key. A Google Group cannot be used.'
                }
            }
            return @{
                Cause  = 'Domain-wide delegation is not authorised for this service account in the user''s domain.'
                Remedy = 'Add the numeric Client ID with the exact scope string in Manage Domain Wide Delegation. Allow up to 24 hours to propagate.'
            }
        }
        'access_denied' {
            return @{
                Cause  = 'One or more requested scopes are NOT authorised in the Admin console. This is a scope mismatch, not a credential problem.'
                Remedy = "Re-authorise using this EXACT string, comma separated, no spaces:`n               $($script:ScopeString)"
            }
        }
        'admin_policy_enforced' {
            return @{
                Cause  = 'A Google Workspace admin policy blocks one or more scopes for this OAuth client.'
                Remedy = 'Google Admin > Security > Access and data control > API controls > App access control. Explicitly trust the app / client ID.'
            }
        }
        'invalid_grant' {
            if ($d -match 'Invalid JWT Signature') {
                return @{
                    Cause  = 'The private key does not match the service account, or the key was deleted or disabled in Google Cloud.'
                    Remedy = 'Generate a fresh JSON key for this service account and use that file.'
                }
            }
            if ($d -match 'short-lived|reasonable timeframe|Invalid JWT:') {
                return @{
                    Cause  = 'This computer''s clock is out of sync with Google. The signed token was rejected as outside its valid window.'
                    Remedy = 'Run: w32tm /resync /force  (elevated), then re-run this script.'
                }
            }
            return @{
                Cause  = 'The impersonated user does not exist in the Google Workspace tenant.'
                Remedy = 'Check the Username column in the CSV. It must be the real Google Workspace address, not an alias or a consumer gmail.com address.'
            }
        }
        'invalid_scope' {
            return @{
                Cause  = 'A scope in the request is malformed or does not exist.'
                Remedy = 'This indicates a script-side problem. Report it with the run transcript.'
            }
        }
        'disabled_client' {
            return @{
                Cause  = 'The key used to sign the request is disabled.'
                Remedy = 'Google Cloud console > IAM & Admin > Service Accounts. Enable the service account and its key.'
            }
        }
        'invalid_client' {
            return @{
                Cause  = 'The OAuth client or the JWT is invalid or misconfigured.'
                Remedy = 'Confirm the JSON key file is complete and unmodified, and that client_email matches the service account.'
            }
        }
        'deleted_client' {
            return @{
                Cause  = 'The OAuth client used for this request has been deleted.'
                Remedy = 'Create a new service account and key, then re-authorise delegation.'
            }
        }
        default {
            return @{
                Cause  = "Google returned '$ErrorCode'."
                Remedy = $d
            }
        }
    }
}

function Test-GoogleDelegation {
    <#
        Requests a real delegated access token for one user. Returns a result
        object; never throws. This is the definitive proof that the Google side
        is configured correctly.
    #>
    param(
        [Parameter(Mandatory)]$KeyObject,
        [Parameter(Mandatory)][string]$UserEmail
    )

    $result = [PSCustomObject]@{
        User = $UserEmail; Success = $false; ErrorCode = ''
        Cause = ''; Remedy = ''; Raw = ''
    }

    try {
        $jwt = New-SignedJwt -KeyObject $KeyObject -Subject $UserEmail -Scopes $script:RequiredScopes
    }
    catch {
        $result.ErrorCode = 'jwt_signing_failed'
        $result.Cause = 'The service account private key could not be used to sign a token.'
        $result.Remedy = "Confirm the JSON key is a valid, unmodified service account key. Detail: $($_.Exception.Message)"
        return $result
    }

    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12 -bor [Net.ServicePointManager]::SecurityProtocol

    $body = 'grant_type=' + [Uri]::EscapeDataString('urn:ietf:params:oauth:grant-type:jwt-bearer') +
            '&assertion=' + [Uri]::EscapeDataString($jwt)

    try {
        $req = [Net.HttpWebRequest]::Create('https://oauth2.googleapis.com/token')
        $req.Method = 'POST'
        $req.ContentType = 'application/x-www-form-urlencoded'
        $req.Timeout = 30000
        try {
            $wp = [Net.WebRequest]::GetSystemWebProxy()
            $wp.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
            $req.Proxy = $wp
        } catch { }

        $bytes = [Text.Encoding]::ASCII.GetBytes($body)
        $req.ContentLength = $bytes.Length
        $s = $req.GetRequestStream(); $s.Write($bytes, 0, $bytes.Length); $s.Close()

        $resp = $req.GetResponse()
        $sr = New-Object IO.StreamReader($resp.GetResponseStream())
        $txt = $sr.ReadToEnd(); $sr.Close(); $resp.Close()

        $json = $txt | ConvertFrom-Json
        if ($json.access_token) {
            $result.Success = $true
            return $result
        }
        $result.ErrorCode = 'no_token'
        $result.Cause = 'Google responded without an access token.'
        $result.Raw = $txt
        return $result
    }
    catch [Net.WebException] {
        $wex = $_.Exception
        $txt = ''
        if ($wex.Response) {
            try {
                $r = New-Object IO.StreamReader($wex.Response.GetResponseStream())
                $txt = $r.ReadToEnd(); $r.Close()
            } catch { }
        }
        $result.Raw = $txt
        $code = ''; $desc = ''
        if ($txt) {
            try {
                $j = $txt | ConvertFrom-Json
                $code = "$($j.error)"; $desc = "$($j.error_description)"
            } catch { $desc = $txt }
        }
        if (-not $code) {
            $result.ErrorCode = 'network'
            $result.Cause = 'Could not reach oauth2.googleapis.com.'
            $result.Remedy = "Check proxy and firewall access to oauth2.googleapis.com. Detail: $($wex.Message)"
            return $result
        }
        $result.ErrorCode = $code
        $decoded = Resolve-GoogleAuthError -ErrorCode $code -Description $desc
        $result.Cause = $decoded.Cause
        $result.Remedy = $decoded.Remedy
        return $result
    }
    catch {
        $result.ErrorCode = 'unexpected'
        $result.Cause = $_.Exception.Message
        return $result
    }
}

# ----------------------------------------------------------------------------------
# GUIDE - the manual Google-side setup, printed exactly
# ----------------------------------------------------------------------------------
function Show-GoogleSetupGuide {
    Write-Head 'GOOGLE SIDE SETUP - DO THIS ONCE, BY HAND'

    Write-Host '  No SDK, no gcloud, no command line. Two browser tabs, about five minutes.' -ForegroundColor Gray
    Write-Host ''

    Write-Host '  STEP 1 - Create a project' -ForegroundColor Cyan
    Write-Host '    console.cloud.google.com  >  project picker  >  New Project'
    Write-Host '    Name it anything (e.g. M365 Migration). Note the Project ID.'
    Write-Host ''

    Write-Host '  STEP 2 - Enable these four APIs' -ForegroundColor Cyan
    Write-Host '    console.cloud.google.com/apis/library  >  search each  >  Enable'
    foreach ($a in $script:RequiredApis) {
        Write-Host ("      - {0,-22} {1}" -f $a.Name, $a.Id)
    }
    Write-Host '    Contacts API is easy to miss and contact migration fails without it.' -ForegroundColor Yellow
    Write-Host ''

    Write-Host '  STEP 3 - Create a service account' -ForegroundColor Cyan
    Write-Host '    console.cloud.google.com/iam-admin/serviceaccounts  >  Create service account'
    Write-Host '    Name it, click Create, skip the optional role and access steps, click Done.'
    Write-Host ''

    Write-Host '  STEP 4 - Create a JSON key' -ForegroundColor Cyan
    Write-Host '    Click the service account  >  Keys tab  >  Add key  >  Create new key  >  JSON'
    Write-Host '    The file downloads. That file is the -KeyPath for this script.'
    Write-Host '    It is a password-equivalent secret. This script deletes it after use.' -ForegroundColor Yellow
    Write-Host ''

    Write-Host '  STEP 5 - Copy the NUMERIC Client ID' -ForegroundColor Cyan
    Write-Host '    On the service account Details tab, copy Unique ID (a long number).'
    Write-Host '    It is also the client_id field inside the JSON key.'
    Write-Host '    Using the service account EMAIL here is the single most common mistake.' -ForegroundColor Yellow
    Write-Host ''

    Write-Host '  STEP 6 - Authorise domain-wide delegation' -ForegroundColor Cyan
    Write-Host '    admin.google.com  >  Security  >  Access and data control  >  API controls'
    Write-Host '    >  Manage Domain Wide Delegation  >  Add new'
    Write-Host '    Client ID: the NUMERIC id from step 5'
    Write-Host '    OAuth scopes: paste this ENTIRE line, exactly, no spaces:' -ForegroundColor White
    Write-Host ''
    Write-Host "      $($script:ScopeString)" -ForegroundColor Green
    Write-Host ''
    Write-Host '    All five scopes are required. An extra or missing scope breaks the' -ForegroundColor Yellow
    Write-Host '    match and the migration fails AFTER the batch starts, not before.' -ForegroundColor Yellow
    Write-Host '    Propagation takes minutes, occasionally up to 24 hours.' -ForegroundColor Yellow
    Write-Host ''

    Write-Head 'MICROSOFT 365 SIDE - REQUIRED BEFORE ANY BATCH'

    Write-Host '  1. Routing subdomains (both, verified and Active)' -ForegroundColor Cyan
    Write-Host '     o365.<yourdomain>    added in Google Admin as a User alias domain,'
    Write-Host '                          MX pointing to Microsoft 365, accepted in M365.'
    Write-Host '     gsuite.<yourdomain>  added in Google Admin as a User alias domain,'
    Write-Host '                          MX pointing to Google.'
    Write-Host '     Do NOT use tenant.onmicrosoft.com as the target delivery domain.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  2. Every user provisioned as a MAIL USER (not a mailbox)' -ForegroundColor Cyan
    Write-Host '     New-MailUser -Name "Will" -ExternalEmailAddress will@gsuite.contoso.com ...'
    Write-Host '     ExternalEmailAddress must point at the GOOGLE routing domain.'
    Write-Host '     Each user also needs a proxy address at the M365 routing domain.'
    Write-Host ''
    Write-Host '  3. Disable MRM and archive policies until migration completes' -ForegroundColor Cyan
    Write-Host '     Otherwise items are flagged "missing" - perceived data loss that is'
    Write-Host '     very hard to separate from real loss during verification.'
    Write-Host ''
    Write-Host '  4. Automatic Forwarding enabled on the Remote Domain' -ForegroundColor Cyan
    Write-Host '     for the Google routing domain, or M365 to Google mail flow breaks.'
    Write-Host ''
    Write-Host '  Not supported: GCC High and DoD. Default max message size: 35 MB.' -ForegroundColor Gray
    Write-Host ''

    Write-Head 'THEN RUN'
    Write-Host '    .\Invoke-GoogleToM365Migration.ps1 -Mode Preflight `' -ForegroundColor Green
    Write-Host '        -KeyPath .\key.json -GoogleAdminEmail admin@contoso.com `' -ForegroundColor Green
    Write-Host '        -CsvPath .\users.csv -TargetDeliveryDomain o365.contoso.com' -ForegroundColor Green
    Write-Host ''
    Write-Host '  Preflight changes nothing. It verifies both sides and tells you' -ForegroundColor Gray
    Write-Host '  exactly what is wrong before a single mailbox is touched.' -ForegroundColor Gray
    Write-Host ''
}

# ----------------------------------------------------------------------------------
# KEY HANDLING
# ----------------------------------------------------------------------------------
function Import-ServiceAccountKey {
    param([string]$Path)

    if (-not $Path) {
        Add-Finding Critical 'Key' 'No service account key supplied.' '' 'Pass -KeyPath pointing at the JSON key downloaded in step 4 of the guide.'
        return $null
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Add-Finding Critical 'Key' "Service account key not found: $Path" '' 'Check the path. Run -Mode Guide for how to create one.'
        return $null
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json
    }
    catch {
        Add-Finding Critical 'Key' 'The key file is not valid JSON.' $_.Exception.Message 'Re-download the JSON key from the Google Cloud console.'
        return $null
    }

    foreach ($f in @('type','client_email','private_key','client_id','private_key_id','project_id')) {
        if (-not $obj.$f) {
            Add-Finding Critical 'Key' "The key file is missing the '$f' field." '' 'This is not a complete service account key. Create a new JSON key.'
            return $null
        }
    }
    if ($obj.type -ne 'service_account') {
        Add-Finding Critical 'Key' "The key type is '$($obj.type)', not 'service_account'." '' 'Download a SERVICE ACCOUNT key, not an OAuth client secret.'
        return $null
    }

    Add-Action 'Load service account key' 'Verified' "$($obj.client_email) (project $($obj.project_id))"
    Write-Info "Numeric Client ID for domain-wide delegation: $($obj.client_id)"
    return $obj
}

function Remove-KeyFileSecurely {
    param([string]$Path)
    if ($KeepKeyFile) {
        Add-Action 'Remove key file' 'Skipped' 'Kept at your request - it is a password-equivalent secret'
        return
    }
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return }
    try {
        $len = (Get-Item -LiteralPath $Path).Length
        $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
        $buf = New-Object -TypeName 'byte[]' -ArgumentList $len
        $rng.GetBytes($buf)
        [IO.File]::WriteAllBytes($Path, $buf)
        $rng.Dispose()
        Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
        Add-Action 'Remove key file' 'Applied' 'Overwritten then deleted from disk'
    }
    catch {
        Add-Action 'Remove key file' 'Failed' "Delete it manually: $Path"
    }
}

# ----------------------------------------------------------------------------------
# CSV - validated, de-duplicated, and written WITHOUT a BOM
# ----------------------------------------------------------------------------------
function Import-MigrationCsv {
    param([string]$Path)

    if (-not $Path) {
        Add-Finding Critical 'CSV' 'No user list supplied.' '' 'Pass -CsvPath. Columns: EmailAddress,Username'
        return @()
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Add-Finding Critical 'CSV' "User list not found: $Path" '' 'Check the path.'
        return @()
    }

    try { $rows = @(Import-Csv -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop) }
    catch {
        Add-Finding Critical 'CSV' 'The user list could not be parsed.' $_.Exception.Message 'Save it as a plain CSV with a header row.'
        return @()
    }
    if ($rows.Count -eq 0) {
        Add-Finding Critical 'CSV' 'The user list has no rows.' '' 'Add at least one user.'
        return @()
    }

    $cols = $rows[0].PSObject.Properties.Name
    # Import-Csv leaves a BOM attached to the first header name.
    $emailCol = @($cols | Where-Object { ($_ -replace "^\xEF\xBB\xBF|^\uFEFF", '') -eq 'EmailAddress' })
    if ($emailCol.Count -eq 0) {
        Add-Finding Critical 'CSV' 'The user list has no EmailAddress column.' "Columns found: $($cols -join ', ')" 'Use headers: EmailAddress,Username'
        return @()
    }
    $emailCol = $emailCol[0]
    $userCol = @($cols | Where-Object { ($_ -replace "^\xEF\xBB\xBF|^\uFEFF", '') -eq 'Username' })

    $rx = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    $mappings = New-Object System.Collections.ArrayList
    $seen = @{}
    $rowNum = 1

    foreach ($r in $rows) {
        $rowNum++
        $target = "$($r.$emailCol)".Trim()
        $source = $target
        if ($userCol.Count -gt 0) {
            $v = "$($r.($userCol[0]))".Trim()
            if ($v) { $source = $v }
        }

        if (-not $target) { Add-Finding Medium 'CSV' "Row $rowNum has an empty EmailAddress - skipped."; continue }
        if ($target -notmatch $rx) { Add-Finding High 'CSV' "Row $rowNum : '$target' is not a valid email address." '' 'Correct the row.'; continue }
        if ($source -notmatch $rx) { Add-Finding High 'CSV' "Row $rowNum : '$source' is not a valid email address." '' 'Correct the Username column.'; continue }
        if ($seen.ContainsKey($target.ToLower())) { Add-Finding High 'CSV' "Row $rowNum : '$target' is listed more than once." '' 'Remove the duplicate.'; continue }

        $seen[$target.ToLower()] = $true
        if ($source -like '*@gmail.com') {
            Add-Finding Medium 'CSV' "Row $rowNum : source '$source' is a consumer gmail.com address." '' 'Google Workspace sources use your managed domain, not gmail.com. Verify this is intended.'
        }
        $null = $mappings.Add([PSCustomObject]@{ EmailAddress = $target; Username = $source })
    }

    if ($mappings.Count -eq 0) {
        Add-Finding Critical 'CSV' 'No valid rows survived validation.' '' 'Fix the errors above.'
        return @()
    }
    if ($mappings.Count -gt 2000) {
        Add-Finding High 'CSV' "$($mappings.Count) users in one batch." '' 'Split into batches of 2000 or fewer.'
    }

    Add-Action 'Validate user list' 'Verified' "$($mappings.Count) unique user(s)"
    return @($mappings)
}

function Export-MigrationCsvNoBom {
    <#
        Windows PowerShell 5.1 writes a BOM with -Encoding UTF8. Those three bytes
        travel into New-MigrationBatch -CSVData and the header parses as
        \uFEFFEmailAddress, so the batch is rejected. This writes clean bytes.
    #>
    param([object[]]$Mappings)
    if (-not $Mappings -or $Mappings.Count -eq 0) { return $null }

    $path = Join-Path $script:RunFolder 'ValidatedMigrationUsers.csv'
    $lines = New-Object System.Collections.ArrayList
    $null = $lines.Add('EmailAddress,Username')
    foreach ($m in $Mappings) { $null = $lines.Add("$($m.EmailAddress),$($m.Username)") }

    $enc = New-Object System.Text.UTF8Encoding($false)   # $false = NO byte order mark
    [System.IO.File]::WriteAllLines($path, [string[]]$lines, $enc)

    $all = [System.IO.File]::ReadAllBytes($path)
    if ($all.Length -lt 3) {
        Add-Finding Critical 'CSV' 'The generated CSV is empty or truncated.' "$($all.Length) byte(s) written." 'Script defect - do not proceed; report this.'
        return $null
    }
    $head = $all[0..2]
    if ($head[0] -eq 0xEF -and $head[1] -eq 0xBB -and $head[2] -eq 0xBF) {
        Add-Finding Critical 'CSV' 'The generated CSV still carries a byte order mark.' '' 'Script defect - do not proceed; report this.'
        return $null
    }

    Add-Action 'Write migration CSV' 'Applied' "$path (BOM-free, verified)"
    return $path
}

# ----------------------------------------------------------------------------------
# EXCHANGE ONLINE - connection with backoff, and full preflight
# ----------------------------------------------------------------------------------
function Invoke-ExoWithRetry {
    <#  Exchange Online throttles. 429 and 5xx are transient; everything else is an answer. #>
    param(
        [Parameter(Mandatory)][scriptblock]$Script,
        [string]$Label = 'Exchange operation',
        [int]$MaxAttempts = 5
    )
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try { return (& $Script) }
        catch {
            $msg = $_.Exception.Message
            $transient = ($msg -match 'throttl|429|503|502|500|timed out|timeout|temporarily|connection was closed|Server is busy')
            if (-not $transient -or $i -eq $MaxAttempts) { throw }
            $delay = [math]::Min(60, 3 * [math]::Pow(2, $i - 1))
            Write-Step "$Label throttled or transient (attempt $i). Waiting $delay s."
            Start-Sleep -Seconds $delay
        }
    }
}

function Connect-ExchangeSafely {
    if ($script:Connected) { return $true }

    $mod = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
           Sort-Object Version -Descending | Select-Object -First 1
    if (-not $mod -or $mod.Version.Major -lt 3) {
        $why = 'is not installed'
        if ($mod) { $why = "is version $($mod.Version), and 3.0.0 or later is required" }
        Write-Warn2 "ExchangeOnlineManagement $why."
        if (Install-ExchangeModuleResilient) {
            $mod = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
                   Sort-Object Version -Descending | Select-Object -First 1
        }
        else {
            Add-Finding Critical 'Exchange' "ExchangeOnlineManagement $why." 'Every installation strategy failed.' 'Install it manually in an elevated session: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force'
            return $false
        }
    }
    try { Import-Module ExchangeOnlineManagement -ErrorAction Stop }
    catch {
        Add-Finding Critical 'Exchange' 'The ExchangeOnlineManagement module failed to load.' $_.Exception.Message 'Reinstall the module.'
        return $false
    }

    $existing = $null
    try { $existing = Get-ConnectionInformation -ErrorAction SilentlyContinue } catch { }
    if ($existing) {
        $script:Connected = $true
        Add-Action 'Connect to Exchange Online' 'Verified' "Existing session as $($existing[0].UserPrincipalName)"
        return $true
    }

    for ($i = 1; $i -le 3; $i++) {
        try {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $script:Connected = $true
            $ci = Get-ConnectionInformation -ErrorAction SilentlyContinue
            $who = 'connected'
            if ($ci) { $who = $ci[0].UserPrincipalName }
            Add-Action 'Connect to Exchange Online' 'Applied' $who
            return $true
        }
        catch {
            Write-Warn2 "Connection attempt $i failed: $($_.Exception.Message)"
            if ($i -lt 3) { Start-Sleep -Seconds (5 * $i) }
        }
    }
    Add-Finding Critical 'Exchange' 'Could not connect to Exchange Online after 3 attempts.' '' 'Check credentials, MFA, and access to outlook.office365.com.'
    return $false
}

function Select-DeliveryDomainInteractive {
    <#
        Offers the tenant's accepted domains when the supplied one is missing,
        blocked, or unknown. Applies the choice at SCRIPT scope immediately -
        setting a variable that is only read earlier in the run would leave the
        migration using the domain that was just rejected.
    #>
    param([string]$Reason = '')

    if ($NonInteractive) { return $false }
    if (-not $script:Connected) { return $false }

    $cands = @(Get-CandidateDeliveryDomains | Select-Object -First 8 | ForEach-Object { $_.DomainName })
    if ($cands.Count -eq 0) {
        Write-Warn2 'No usable accepted domains found - every one is an onmicrosoft.com domain.'
        Write-Info  'Add and verify a routing subdomain such as o365.contoso.com, then re-run.'
        return $false
    }

    Write-Host ''
    if ($Reason) { Write-Info $Reason }
    Write-Host '  Accepted domains in this tenant:' -ForegroundColor Cyan
    for ($i = 0; $i -lt $cands.Count; $i++) { Write-Host ("   {0}  {1}" -f ($i + 1), $cands[$i]) }
    Write-Host ''
    Write-Host '  A routing subdomain (o365.yourdomain.com) is the documented choice.' -ForegroundColor DarkGray
    Write-Host ''

    $pick = "$(Read-Host '  Use one of these? number, or Enter to skip')".Trim()
    if (-not $pick) { return $false }

    $n = 0
    if (-not [int]::TryParse($pick, [ref]$n) -or $n -lt 1 -or $n -gt $cands.Count) {
        Write-Warn2 'That was not one of the listed numbers - keeping the original value.'
        return $false
    }

    $chosen = $cands[$n - 1]
    $script:ResolvedTargetDomain = $chosen
    Set-Variable -Name TargetDeliveryDomain -Value $chosen -Scope Script
    Write-Good "Using $chosen"

    # Clear only the routing findings this choice actually resolves.
    $stale = @($script:Findings | Where-Object {
        $_.Area -eq 'Routing' -and $_.Severity -eq 'Critical'
    })
    foreach ($f in $stale) { $null = $script:Findings.Remove($f) }
    Add-Action 'Target delivery domain' 'Verified' "$chosen selected from accepted domains"
    return $true
}

function Test-ExchangePrerequisites {
    param([object[]]$Mappings)

    if (-not $script:Connected) {
        # A preflight that checked nothing must never report success.
        Add-Finding Critical 'Exchange' 'Exchange checks did not run - no session.' '' 'Resolve the connection failure above and re-run.'
        return
    }

    Write-Step 'Checking the target delivery domain'
    $accepted = @()
    try { $accepted = @(Invoke-ExoWithRetry -Label 'Get-AcceptedDomain' -Script { Get-AcceptedDomain -ErrorAction Stop }) }
    catch { Add-Finding High 'Exchange' 'Could not read accepted domains.' $_.Exception.Message 'Confirm the Recipient Management role.' }

    if (-not $TargetDeliveryDomain) {
        Add-Finding Critical 'Routing' 'No target delivery domain supplied.' '' 'Pass -TargetDeliveryDomain, e.g. o365.contoso.com. Do not use onmicrosoft.com.'
        $null = Select-DeliveryDomainInteractive -Reason 'No target delivery domain was supplied.'
    }
    elseif ($TargetDeliveryDomain -like '*.onmicrosoft.com') {
        Add-Finding Critical 'Routing' "'$TargetDeliveryDomain' is an onmicrosoft.com domain." 'Microsoft warns this causes issues it cannot assist with.' 'Use a verified routing subdomain such as o365.contoso.com.'
        $null = Select-DeliveryDomainInteractive -Reason "'$TargetDeliveryDomain' is an onmicrosoft.com domain, which Microsoft advises against."
    }
    elseif ($accepted.Count -gt 0) {
        $hit = @($accepted | Where-Object { $_.DomainName -eq $TargetDeliveryDomain })
        if ($hit.Count -eq 0) {
            $cands = @(Get-CandidateDeliveryDomains | Select-Object -First 6 | ForEach-Object { $_.DomainName })
            $hint = 'Add and verify the routing subdomain in the Microsoft 365 admin center.'
            if ($cands.Count -gt 0) { $hint = "Accepted domains you could use instead: $($cands -join ', ')" }
            Add-Finding Critical 'Routing' "'$TargetDeliveryDomain' is not an accepted domain in this tenant." "Checked $($accepted.Count) accepted domain(s)." $hint
            $null = Select-DeliveryDomainInteractive -Reason "'$TargetDeliveryDomain' is not an accepted domain in this tenant."
        }
        else {
            Add-Action 'Target delivery domain' 'Verified' "$TargetDeliveryDomain is accepted"
        }
    }

    Write-Step 'Checking recipients are MailUsers'
    $notFound = New-Object System.Collections.ArrayList
    $wrongType = New-Object System.Collections.ArrayList
    $noExternal = New-Object System.Collections.ArrayList

    foreach ($m in $Mappings) {
        $r = $null
        try { $r = Invoke-ExoWithRetry -Label 'Get-Recipient' -Script { Get-Recipient -Identity $m.EmailAddress -ErrorAction Stop } }
        catch { $null = $notFound.Add($m.EmailAddress); continue }

        if ($r.RecipientTypeDetails -ne 'MailUser') {
            $null = $wrongType.Add("$($m.EmailAddress) is $($r.RecipientTypeDetails)")
        }
        else {
            $ext = "$($r.ExternalEmailAddress)"
            if (-not $ext) { $null = $noExternal.Add($m.EmailAddress) }
        }
    }

    if ($notFound.Count -gt 0) {
        $script:MissingRecipients = @($Mappings | Where-Object { $notFound -contains $_.EmailAddress })

        $created = 0
        if (-not $script:AllowMutation) {
            # Read-only mode. Say exactly what would be done, then do nothing.
            Write-Host ''
            Write-Info "Preflight is read-only, so nothing was created."
            if ($script:AutoCreateMailUsers -and $script:GoogleRouting) {
                Write-Info "In setup mode I would mail-enable $($notFound.Count) account(s), routed to <user>@$($script:GoogleRouting)."
            }
        }
        elseif ($script:AutoCreateMailUsers -and $script:GoogleRouting) {
            $created = New-MailUsersForMissing -Missing $script:MissingRecipients -GoogleRoutingDomain $script:GoogleRouting
        }

        if ($created -ge $notFound.Count) {
            Add-Action 'Provision mail users' 'Applied' "$created created - all targets now exist"
        }
        else {
            $left = $notFound.Count - $created
            Add-Finding Critical 'Recipients' "$left user(s) do not exist in Exchange Online." (@($notFound | Select-Object -First 8) -join ', ') "Create them as mail users: New-MailUser -Name <n> -ExternalEmailAddress <user>@$($script:GoogleRouting) -MicrosoftOnlineServicesID <user>@yourdomain.com"
        }
    }
    if ($wrongType.Count -gt 0) {
        Add-Finding Critical 'Recipients' "$($wrongType.Count) target(s) are not MailUser objects." (@($wrongType | Select-Object -First 8) -join '; ') 'Google Workspace migration requires MailUser targets, not existing mailboxes.'
    }
    if ($noExternal.Count -gt 0) {
        Add-Finding High 'Recipients' "$($noExternal.Count) mail user(s) have no ExternalEmailAddress." (@($noExternal | Select-Object -First 8) -join ', ') 'Set it to the user at the Google routing domain, e.g. will@gsuite.contoso.com.'
    }
    if ($notFound.Count -eq 0 -and $wrongType.Count -eq 0) {
        Add-Action 'Recipient type check' 'Verified' "$($Mappings.Count) target(s) are MailUsers"
    }

    Write-Step 'Checking retention and archive policies'
    try {
        $mrm = @(Invoke-ExoWithRetry -Label 'Get-RetentionPolicy' -Script { Get-RetentionPolicy -ErrorAction Stop })
        $applied = @($mrm | Where-Object { $_.Name -match 'Default MRM|Arbitration' })
        if ($applied.Count -gt 0) {
            Add-Finding Medium 'Retention' 'Retention or archive policies exist in this tenant.' (@($applied | ForEach-Object { $_.Name }) -join ', ') 'Microsoft strongly recommends disabling MRM and archive policies for migrating users until the migration completes. Otherwise items are flagged "missing" - perceived data loss that is hard to separate from real loss.'
        }
    }
    catch { Write-Step 'Retention policies could not be read (non-blocking).' }
}

function Test-GooglePrerequisites {
    param($KeyObject, [object[]]$Mappings)

    if ($SkipGoogleTest) {
        Add-Finding High 'Google' 'The Google delegation test was skipped.' '' 'Domain-wide delegation and scopes are therefore UNVERIFIED. Re-run without -SkipGoogleTest before starting a batch.'
        return
    }
    if (-not $KeyObject) { return }

    Write-Step 'Proving domain-wide delegation with a real token request'

    $targets = New-Object System.Collections.ArrayList
    if ($GoogleAdminEmail) { $null = $targets.Add($GoogleAdminEmail) }
    $userSample = @($Mappings | ForEach-Object { $_.Username })
    if ($DelegationTestCount -gt 0) { $userSample = @($userSample | Select-Object -First $DelegationTestCount) }
    foreach ($u in $userSample) {
        if (-not $u) { continue }
        $dup = @($targets | Where-Object { $_ -and $_.ToLower() -eq $u.ToLower() })
        if ($dup.Count -eq 0) { $null = $targets.Add($u) }
    }

    if ($targets.Count -eq 0) {
        Add-Finding High 'Google' 'No user available to test delegation against.' '' 'Supply -GoogleAdminEmail or a CSV with a Username column.'
        return
    }

    $ok = 0
    $failed = New-Object System.Collections.ArrayList
    foreach ($t in $targets) {
        $r = Test-GoogleDelegation -KeyObject $KeyObject -UserEmail $t
        if ($r.Success) { $ok++; Write-Good "Delegated token issued for $t" }
        else {
            $null = $failed.Add($r)
            Write-Bad "Delegation FAILED for $t  [$($r.ErrorCode)]"
        }
    }

    if ($failed.Count -eq 0) {
        Add-Action 'Google domain-wide delegation' 'Verified' "$ok of $($targets.Count) user(s) - scopes and key are correct"
        return
    }

    $first = $failed[0]
    $sev = 'Critical'
    # A single bad user is a data problem; a systemic error is a configuration problem.
    if ($first.ErrorCode -eq 'invalid_grant' -and $ok -gt 0) { $sev = 'High' }

    Add-Finding $sev 'Google' "Domain-wide delegation failed for $($failed.Count) of $($targets.Count) user(s) [$($first.ErrorCode)]" $first.Cause $first.Remedy

    if ($first.ErrorCode -in @('unauthorized_client','access_denied')) {
        Write-Host ''
        Write-Host '  Paste this EXACT scope string in Manage Domain Wide Delegation:' -ForegroundColor Cyan
        Write-Host "    $($script:ScopeString)" -ForegroundColor Green
        Write-Host "  Client ID (numeric, NOT the email): $($KeyObject.client_id)" -ForegroundColor Cyan
        Write-Host ''

        # Propagation, not misconfiguration, is the usual cause. Offer to wait.
        # Retry the user that actually FAILED - retrying one that already worked
        # would return an instant false success.
        if ($script:WaitForDelegation) {
            $probe = $first.User
            Write-Info "Retrying the account that failed: $probe"
            $retry = Wait-ForGoogleDelegation -KeyObject $KeyObject -UserEmail $probe -MaxWaitMinutes 30

            if ($retry.Success) {
                # One success is not proof. Re-test every user that failed before
                # declaring the delegation healthy.
                $stillBad = New-Object System.Collections.ArrayList
                foreach ($f2 in $failed) {
                    if ($f2.User -eq $probe) { continue }
                    $again = Test-GoogleDelegation -KeyObject $KeyObject -UserEmail $f2.User
                    if ($again.Success) { Write-Good "Delegated token issued for $($f2.User)" }
                    else {
                        $null = $stillBad.Add($again)
                        Write-Bad "Still failing for $($f2.User) [$($again.ErrorCode)]"
                    }
                }

                if ($stillBad.Count -eq 0) {
                    # Remove only the delegation finding, not every Google finding.
                    $stale = @($script:Findings | Where-Object {
                        $_.Area -eq 'Google' -and $_.Title -like 'Domain-wide delegation failed*'
                    })
                    foreach ($f3 in $stale) { $null = $script:Findings.Remove($f3) }
                    Add-Action 'Google domain-wide delegation' 'Verified' "All $($failed.Count) previously failing user(s) now succeed"
                }
                else {
                    Add-Finding High 'Google' "$($stillBad.Count) user(s) still fail after the wait." (@($stillBad | ForEach-Object { "$($_.User) [$($_.ErrorCode)]" }) -join ', ') $stillBad[0].Remedy
                }
            }
        }
    }
}

function Test-MigrationEndpointHealth {
    param($KeyObject)
    if (-not $script:Connected -or -not $KeyObject) { return }
    if (-not $GoogleAdminEmail) {
        Add-Finding High 'Endpoint' 'No Google admin email supplied.' '' 'Pass -GoogleAdminEmail - the endpoint is created against it.'
        return
    }
    Write-Step 'Testing migration server availability'
    try {
        $r = Invoke-ExoWithRetry -Label 'Test-MigrationServerAvailability' -Script {
            Test-MigrationServerAvailability -Gmail -ServiceAccountKeyFileData $script:KeyBytes `
                -EmailAddress $GoogleAdminEmail -ErrorAction Stop
        }
        if ($r.Result -eq 'Success') { Add-Action 'Migration server availability' 'Verified' 'Exchange can reach Google with this key' }
        else {
            Add-Finding Critical 'Endpoint' "Migration server availability returned '$($r.Result)'." "$($r.Message)" 'Usually the delegation or scope problem reported above. Fix that first.'
        }
    }
    catch {
        Add-Finding Critical 'Endpoint' 'Test-MigrationServerAvailability failed.' $_.Exception.Message 'Confirm domain-wide delegation, the scope string, and the admin email.'
    }
}

# ----------------------------------------------------------------------------------
# INTERACTIVE WIZARD
#   Asks only what it cannot discover, offers what it can, and defaults to the
#   safe answer every time.
# ----------------------------------------------------------------------------------
function Read-YesNoSafe {
    param([string]$Prompt, [bool]$DefaultYes = $true)
    if ($NonInteractive) { return $DefaultYes }
    $hint = '[Y/n]'
    if (-not $DefaultYes) { $hint = '[y/N]' }
    while ($true) {
        $a = "$(Read-Host "  $Prompt $hint")".Trim().ToLower()
        if (-not $a) { return $DefaultYes }
        if ($a -in @('y','yes')) { return $true }
        if ($a -in @('n','no'))  { return $false }
        Write-Host '  Please answer y or n.' -ForegroundColor Yellow
    }
}

function Read-TextSafe {
    param([string]$Prompt, [string]$Default = '', [string]$Pattern = '', [switch]$AllowEmpty)
    if ($NonInteractive) { return $Default }
    while ($true) {
        $shown = $Prompt
        if ($Default) { $shown = "$Prompt [$Default]" }
        $a = "$(Read-Host "  $shown")".Trim()
        if (-not $a -and $Default) { return $Default }
        if (-not $a -and $AllowEmpty) { return '' }
        if (-not $a) { Write-Host '  This one is required.' -ForegroundColor Yellow; continue }
        if ($Pattern -and $a -notmatch $Pattern) {
            Write-Host '  That does not look right - try again.' -ForegroundColor Yellow
            continue
        }
        return $a
    }
}

function Read-PathSafe {
    param([string]$Prompt, [string]$Filter = '*')
    if ($NonInteractive) { return '' }
    while ($true) {
        $a = "$(Read-Host "  $Prompt")".Trim().Trim('"')
        if (-not $a) { return '' }
        if (Test-Path -LiteralPath $a -PathType Leaf) { return (Resolve-Path -LiteralPath $a).Path }
        Write-Host "  Not found: $a" -ForegroundColor Yellow
        if (-not (Read-YesNoSafe 'Try again?' $true)) { return '' }
    }
}

function Find-ServiceAccountKeys {
    <#  Looks where the browser actually puts the downloaded key.  #>
    $dirs = @(
        (Join-Path $env:USERPROFILE 'Downloads')
        (Join-Path $env:USERPROFILE 'Desktop')
        (Get-Location).Path
    ) | Select-Object -Unique

    $found = New-Object System.Collections.ArrayList
    foreach ($d in $dirs) {
        if (-not (Test-Path -LiteralPath $d)) { continue }
        $files = @(Get-ChildItem -LiteralPath $d -Filter '*.json' -File -ErrorAction SilentlyContinue |
                   Sort-Object LastWriteTime -Descending | Select-Object -First 25)
        foreach ($f in $files) {
            try {
                $j = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
                if ($j.type -eq 'service_account' -and $j.client_email -and $j.private_key) {
                    $null = $found.Add([PSCustomObject]@{
                        Path = $f.FullName; Email = $j.client_email
                        Project = $j.project_id; Modified = $f.LastWriteTime
                    })
                }
            } catch { }
        }
    }
    return @($found | Sort-Object Modified -Descending)
}

function Invoke-Wizard {
    Clear-Host
    Write-Host ''
    Write-Host '  ============================================================' -ForegroundColor DarkCyan
    Write-Host '   GOOGLE WORKSPACE  ->  MICROSOFT 365' -ForegroundColor Cyan
    Write-Host '   Mailbox migration assistant' -ForegroundColor Cyan
    Write-Host '  ============================================================' -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host '   No Google SDK needed. Nothing is changed without your approval.' -ForegroundColor DarkGray
    Write-Host ''

    $prev = Get-ResumeState
    if ($prev) {
        Write-Host "  Last run: $($prev.When)  mode $($prev.Mode)" -ForegroundColor DarkGray
        if ($prev.BatchName) { Write-Host "  Batch: $($prev.BatchName)" -ForegroundColor DarkGray }
        Write-Host ''
    }

    Write-Host '  WHAT DO YOU WANT TO DO?' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '   1  Show me the Google setup steps          (read only)'
    Write-Host '   2  Check everything is ready               (changes nothing)'
    Write-Host '   3  Set up the migration and create a batch (does not start it)'
    Write-Host '   4  Start a batch I already created'
    Write-Host '   5  Watch a running migration'
    Write-Host '   6  Finish the migration - cutover'
    Write-Host ''

    $choice = ''
    while ($choice -notin @('1','2','3','4','5','6')) {
        $choice = "$(Read-Host '  Choose 1-6 [2]')".Trim()
        if (-not $choice) { $choice = '2' }
        if ($choice -notin @('1','2','3','4','5','6')) { Write-Host '  Enter a number from 1 to 6.' -ForegroundColor Yellow }
    }

    $s = [PSCustomObject]@{
        Mode = 'Preflight'; KeyPath = $KeyPath; Admin = $GoogleAdminEmail
        Csv = $CsvPath; Target = $TargetDeliveryDomain
        Endpoint = $EndpointName; Batch = $BatchName
        GoogleRouting = ''; WaitForDelegation = $true; AutoCreateMailUsers = $false
        Cutover = $false
    }
    switch ($choice) {
        '1' { $s.Mode = 'Guide';    return $s }
        '2' { $s.Mode = 'Preflight' }
        '3' { $s.Mode = 'Run' }
        '4' { $s.Mode = 'StartBatch' }
        '5' { $s.Mode = 'Monitor' }
        '6' { $s.Mode = 'Complete'; $s.Cutover = $true }
    }

    # --- Batch name, for the modes that only need that -----------------------------
    if ($s.Mode -in @('StartBatch','Monitor','Complete')) {
        $def = $s.Batch
        if ($prev -and $prev.BatchName) { $def = $prev.BatchName }
        $s.Batch = Read-TextSafe -Prompt 'Batch name' -Default $def
        if ($s.Mode -eq 'Complete') {
            Write-Host ''
            Write-Warn2 'Completing the batch is the CUTOVER. It redirects live mail flow.'
            $s.Cutover = Read-YesNoSafe 'Are you ready to cut over?' $false
        }
        return $s
    }

    # --- Service account key -------------------------------------------------------
    Write-Head 'GOOGLE SERVICE ACCOUNT KEY'
    if (-not $s.KeyPath) {
        $keys = Find-ServiceAccountKeys
        if ($keys.Count -gt 0) {
            Write-Host '  I found these service account keys:' -ForegroundColor Cyan
            Write-Host ''
            for ($i = 0; $i -lt [math]::Min(5, $keys.Count); $i++) {
                Write-Host ("   {0}  {1}" -f ($i+1), $keys[$i].Email)
                Write-Host ("      project {0}  |  {1}  |  {2}" -f $keys[$i].Project, $keys[$i].Modified, (Split-Path $keys[$i].Path -Leaf)) -ForegroundColor DarkGray
            }
            Write-Host ''
            Write-Host '   0  None of these - let me type a path'
            Write-Host ''
            $pick = "$(Read-Host '  Which key? [1]')".Trim()
            if (-not $pick) { $pick = '1' }
            $parsed = 0
            if ($pick -ne '0' -and [int]::TryParse($pick, [ref]$parsed)) {
                $idx = $parsed - 1
                if ($idx -ge 0 -and $idx -lt $keys.Count) { $s.KeyPath = $keys[$idx].Path }
                else { Write-Warn2 'That number was not in the list.' }
            }
        }
        if (-not $s.KeyPath) {
            Write-Host '  If you do not have one yet, choose option 1 from the main menu first.' -ForegroundColor DarkGray
            $s.KeyPath = Read-PathSafe -Prompt 'Full path to the JSON key'
        }
    }
    if ($s.KeyPath) { Write-Good "Using $(Split-Path $s.KeyPath -Leaf)" }

    # --- Google admin --------------------------------------------------------------
    Write-Host ''
    $emailRx = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    $defAdmin = $s.Admin
    if (-not $defAdmin -and $prev) { $defAdmin = "$($prev.GoogleAdminEmail)" }
    $s.Admin = Read-TextSafe -Prompt 'Google Workspace admin email' -Default $defAdmin -Pattern $emailRx

    # --- Domains -------------------------------------------------------------------
    Write-Head 'MAIL ROUTING'
    Write-Host '  Migration needs two routing subdomains. Example for contoso.com:' -ForegroundColor Gray
    Write-Host '    o365.contoso.com     routes to Microsoft 365   (target delivery domain)'
    Write-Host '    gsuite.contoso.com   routes to Google          (mail user external address)'
    Write-Host ''

    $defTarget = $s.Target
    if (-not $defTarget -and $prev) { $defTarget = "$($prev.TargetDeliveryDomain)" }
    $s.Target = Read-TextSafe -Prompt 'Target delivery domain (M365 routing subdomain)' -Default $defTarget -Pattern '^[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$'

    if ($s.Target -like '*.onmicrosoft.com') {
        Write-Host ''
        Write-Warn2 'Microsoft advises against onmicrosoft.com here - it causes issues they cannot assist with.'
        if (-not (Read-YesNoSafe 'Use it anyway?' $false)) {
            $s.Target = Read-TextSafe -Prompt 'Target delivery domain' -Pattern '^[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$'
        }
    }

    $guess = ''
    if ($s.Admin -match '@(.+)$') { $guess = "gsuite.$($matches[1])" }
    $s.GoogleRouting = Read-TextSafe -Prompt 'Google routing subdomain' -Default $guess -Pattern '^[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$'

    # --- User list -----------------------------------------------------------------
    Write-Head 'WHO IS MIGRATING'
    if (-not $s.Csv) {
        Write-Host '   1  I have a CSV file'
        Write-Host '   2  Build the list from mail users already in Microsoft 365'
        Write-Host ''
        $src = "$(Read-Host '  Choose 1-2 [1]')".Trim()
        if (-not $src) { $src = '1' }
        if ($src -eq '1') {
            Write-Host '  Columns: EmailAddress,Username   (target M365, source Google)' -ForegroundColor DarkGray
            $s.Csv = Read-PathSafe -Prompt 'Path to the CSV'
        }
        else { $s.Csv = '<auto>' }
    }

    # --- Resilience choices ---------------------------------------------------------
    Write-Head 'IF SOMETHING IS NOT READY'
    Write-Host '  Delegation can take minutes, occasionally up to 24 hours, to apply.' -ForegroundColor Gray
    $s.WaitForDelegation = Read-YesNoSafe 'Wait and keep retrying instead of failing?' $true

    Write-Host ''
    Write-Host '  Migration requires each user to exist as a mail user in Microsoft 365.' -ForegroundColor Gray
    Write-Host '  A mail user is a directory entry only - no mailbox, no data, reversible.' -ForegroundColor Gray
    $s.AutoCreateMailUsers = Read-YesNoSafe 'Offer to create any that are missing?' $true

    # --- Confirm ---------------------------------------------------------------------
    Write-Head 'PLAN'
    Write-Host "    Action:        $($s.Mode)"
    Write-Host "    Key:           $(if($s.KeyPath){Split-Path $s.KeyPath -Leaf}else{'none'})"
    Write-Host "    Google admin:  $($s.Admin)"
    Write-Host "    Target domain: $($s.Target)"
    Write-Host "    Google domain: $($s.GoogleRouting)"
    Write-Host "    Users:         $(if($s.Csv -eq '<auto>'){'discover from Microsoft 365'}else{$s.Csv})"
    Write-Host "    Batch:         $($s.Batch)"
    Write-Host ''
    if ($s.Mode -eq 'Run') {
        Write-Warn2 'This creates the endpoint and the batch. The batch is NOT started.'
    }
    else { Write-Info 'Nothing will be changed.' }
    Write-Host ''
    if (-not (Read-YesNoSafe 'Start?' $true)) { return $null }
    return $s
}

# ----------------------------------------------------------------------------------
# RESILIENCE - try every safe path before reporting failure
# ----------------------------------------------------------------------------------
function Install-ExchangeModuleResilient {
    <#
        Installs / repairs ExchangeOnlineManagement through every documented path.
        The three most common causes of Install-Module failing on a clean machine
        are handled BEFORE the first attempt: TLS 1.2 off, no NuGet provider,
        untrusted PSGallery.
    #>
    $have = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
            Sort-Object Version -Descending | Select-Object -First 1
    if ($have -and $have.Version.Major -ge 3) {
        try { Import-Module ExchangeOnlineManagement -ErrorAction Stop; return $true } catch { }
    }

    Write-Step 'Preparing PowerShell Gallery access'
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.SecurityProtocolType]::Tls12 -bor [Net.ServicePointManager]::SecurityProtocol
    } catch { }
    try {
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
        }
    } catch { }

    # Remember the original trust level and restore it afterwards.
    $originalPolicy = $null
    try { $originalPolicy = (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue).InstallationPolicy } catch { }
    try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch { }

    $strategies = @(
        @{ Label = 'CurrentUser';                    Args = @{ Scope='CurrentUser'; Force=$true; AllowClobber=$true } }
        @{ Label = 'CurrentUser, skip publisher';    Args = @{ Scope='CurrentUser'; Force=$true; AllowClobber=$true; SkipPublisherCheck=$true } }
        @{ Label = 'AllUsers (needs elevation)';     Args = @{ Scope='AllUsers';    Force=$true; AllowClobber=$true } }
    )

    $ok = $false
    foreach ($st in $strategies) {
        Write-Step "Installing ExchangeOnlineManagement - $($st.Label)"
        try {
            $p = $st.Args.Clone()
            $p.Name = 'ExchangeOnlineManagement'
            $p.MinimumVersion = '3.0.0'
            $p.ErrorAction = 'Stop'
            Install-Module @p
            Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
            $ok = $true
            break
        }
        catch { Write-Step "  did not work: $($_.Exception.Message)" }
    }

    if ($originalPolicy -and $originalPolicy -ne 'Trusted') {
        try { Set-PSRepository -Name PSGallery -InstallationPolicy $originalPolicy -ErrorAction SilentlyContinue } catch { }
    }

    if ($ok) { Add-Action 'Install Exchange module' 'Applied' 'ExchangeOnlineManagement is ready' }
    return $ok
}

function Wait-ForGoogleDelegation {
    <#
        Domain-wide delegation takes minutes, occasionally up to 24 hours, to
        propagate. Rather than failing, this polls until Google issues a token.
        This single behaviour is what turns "run it again tomorrow and hope" into
        a migration that starts the moment the tenant is genuinely ready.
    #>
    param(
        [Parameter(Mandatory)]$KeyObject,
        [Parameter(Mandatory)][string]$UserEmail,
        [int]$MaxWaitMinutes = 30
    )

    $deadline = (Get-Date).AddMinutes($MaxWaitMinutes)
    $attempt = 0
    $delay = 20

    Write-Host ''
    Write-Info "Waiting for Google to apply the delegation (up to $MaxWaitMinutes min). Ctrl+C is safe."

    while ((Get-Date) -lt $deadline) {
        $attempt++
        $r = Test-GoogleDelegation -KeyObject $KeyObject -UserEmail $UserEmail
        if ($r.Success) {
            Write-Host ''
            Add-Action 'Google delegation' 'Verified' "Token issued for $UserEmail after $attempt attempt(s)"
            return $r
        }

        # Only these mean "not propagated yet". Everything else is a real
        # misconfiguration and waiting cannot fix it - fail fast instead.
        if ($r.ErrorCode -notin @('unauthorized_client','access_denied')) {
            Write-Host ''
            Write-Bad "This will not resolve by waiting: $($r.ErrorCode)"
            return $r
        }

        $left = [int]($deadline - (Get-Date)).TotalMinutes
        Write-Host "    attempt $attempt - not yet authorised, ${left} min left. Retrying in ${delay}s." -ForegroundColor DarkGray
        Start-Sleep -Seconds $delay
        if ($delay -lt 60) { $delay += 10 }
    }

    Write-Host ''
    Write-Warn2 "Delegation was still not active after $MaxWaitMinutes minutes."
    Write-Info  'Google allows up to 24 hours. Re-run Preflight later - nothing has been changed.'
    return (Test-GoogleDelegation -KeyObject $KeyObject -UserEmail $UserEmail)
}

function Get-CandidateDeliveryDomains {
    <#  Suggests valid routing domains, ranked, excluding the blocked ones.  #>
    if (-not $script:Connected) { return @() }
    $all = @()
    try { $all = @(Invoke-ExoWithRetry -Label 'Get-AcceptedDomain' -Script { Get-AcceptedDomain -ErrorAction Stop }) }
    catch { return @() }

    $usable = @($all | Where-Object { $_.DomainName -notlike '*.onmicrosoft.com' })
    # A dedicated routing subdomain is the documented best practice.
    $routing = @($usable | Where-Object { $_.DomainName -match '^(o365|m365|office|mail|migrate)\.' })
    $rest    = @($usable | Where-Object { $_.DomainName -notmatch '^(o365|m365|office|mail|migrate)\.' })
    return @($routing + $rest)
}

function Get-ExistingMailUsers {
    <#  Builds a candidate user list from MailUsers already provisioned.  #>
    if (-not $script:Connected) { return @() }
    try {
        $mu = @(Invoke-ExoWithRetry -Label 'Get-Recipient MailUser' -Script {
            Get-Recipient -RecipientTypeDetails MailUser -ResultSize Unlimited -ErrorAction Stop
        })
        return @($mu | ForEach-Object {
            $ext = "$($_.ExternalEmailAddress)" -replace '^SMTP:', ''
            [PSCustomObject]@{
                EmailAddress = "$($_.PrimarySmtpAddress)"
                Username     = $ext
                DisplayName  = "$($_.DisplayName)"
            }
        })
    }
    catch { return @() }
}

function New-MailUsersForMissing {
    [CmdletBinding(SupportsShouldProcess = $true)]
    <#
        Creates the MailUser objects the migration requires. This is the single
        biggest blocker in practice, and it is safe and reversible: a MailUser is
        a directory object with no mailbox and no data.
    #>
    param([object[]]$Missing, [string]$GoogleRoutingDomain)

    if (-not $Missing -or $Missing.Count -eq 0) { return 0 }
    if (-not $script:AllowMutation) {
        Add-Action 'Provision mail users' 'Skipped' 'Read-only mode - re-run in setup mode to apply'
        return 0
    }
    if (-not $GoogleRoutingDomain) {
        Add-Action 'Create mail users' 'Skipped' 'No Google routing domain supplied'
        return 0
    }

    Write-Host ''
    Write-Warn2 "$($Missing.Count) user(s) are not provisioned in Exchange Online."
    Write-Info  'A mail user is a directory entry only - no mailbox, no data, fully reversible.'
    Write-Info  "ExternalEmailAddress will point at: <user>@$GoogleRoutingDomain"

    if (-not (Read-YesNoSafe "Create $($Missing.Count) mail user(s) now?" $false)) {
        Add-Action 'Create mail users' 'Skipped' 'Declined'
        return 0
    }

    $made = 0
    $noAccount = New-Object System.Collections.ArrayList

    foreach ($m in $Missing) {
        $local = ($m.EmailAddress -split '@')[0]
        $ext = "$local@$GoogleRoutingDomain"

        # Does a directory account already exist? If so we mail-enable it, which
        # needs no password. Creating a brand new account requires one, and this
        # script will not invent a password on an administrator's behalf.
        $user = $null
        try { $user = Invoke-ExoWithRetry -Label 'Get-User' -Script { Get-User -Identity $m.EmailAddress -ErrorAction Stop } }
        catch { $user = $null }

        if (-not $user) {
            $null = $noAccount.Add($m.EmailAddress)
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($m.EmailAddress, "Mail-enable, routed to $ext")) { continue }

        try {
            $null = Invoke-ExoWithRetry -Label 'Enable-MailUser' -Script {
                Enable-MailUser -Identity $m.EmailAddress -ExternalEmailAddress $ext -ErrorAction Stop
            }
            $made++
            Write-Good "mail-enabled $($m.EmailAddress) -> $ext"
        }
        catch { Write-Bad "could not mail-enable $($m.EmailAddress): $($_.Exception.Message)" }
    }

    if ($noAccount.Count -gt 0) {
        Add-Finding Critical 'Recipients' "$($noAccount.Count) user(s) have no account in this tenant at all." (@($noAccount | Select-Object -First 8) -join ', ') "These need an account created first - which requires setting a password, so this script will not do it. Create each one, then re-run:`n               New-MailUser -Name <name> -MicrosoftOnlineServicesID <user>@yourdomain.com -ExternalEmailAddress <user>@$GoogleRoutingDomain -Password (Read-Host -AsSecureString)"
    }
    if ($made -gt 0) { Add-Action 'Provision mail users' 'Applied' "$made account(s) mail-enabled" }
    return $made
}

function Save-ResumeState {
    if (-not $script:RunFolder) { return }
    $root = Split-Path $script:RunFolder -Parent
    $file = Join-Path $root 'last-run.json'
    try {
        [PSCustomObject]@{
            When = (Get-Date).ToString('s'); Mode = $Mode
            GoogleAdminEmail = $GoogleAdminEmail
            TargetDeliveryDomain = $TargetDeliveryDomain
            EndpointName = $EndpointName; BatchName = $BatchName
            CsvPath = $CsvPath; RunFolder = $script:RunFolder
            Blocked = (Test-Blocking)
        } | ConvertTo-Json | Out-File -LiteralPath $file -Encoding UTF8 -Force
    } catch { }
}

function Get-ResumeState {
    $root = $OutputRoot
    if (-not $root) {
        $dl = Join-Path $env:USERPROFILE 'Downloads'
        $root = Join-Path $dl 'GoogleM365Migration'
    }
    $file = Join-Path $root 'last-run.json'
    if (-not (Test-Path -LiteralPath $file)) { return $null }
    try { return (Get-Content -LiteralPath $file -Raw | ConvertFrom-Json) } catch { return $null }
}

# ----------------------------------------------------------------------------------
# MIGRATION OPERATIONS
# ----------------------------------------------------------------------------------
function Invoke-CreateEndpoint {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    Assert-NoBlockers -Because 'create the migration endpoint'

    $existing = $null
    try { $existing = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue } catch { }
    if ($existing) {
        Add-Action 'Create migration endpoint' 'NotNeeded' "'$EndpointName' already exists"
        return $true
    }
    if (-not $PSCmdlet.ShouldProcess($EndpointName, 'Create Gmail migration endpoint')) {
        Add-Action 'Create migration endpoint' 'Skipped' 'Declined'
        return $false
    }
    try {
        $null = Invoke-ExoWithRetry -Label 'New-MigrationEndpoint' -Script {
            New-MigrationEndpoint -Gmail -ServiceAccountKeyFileData $script:KeyBytes `
                -EmailAddress $GoogleAdminEmail -Name $EndpointName -ErrorAction Stop
        }
        $check = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue
        if ($check) { Add-Action 'Create migration endpoint' 'Applied' "'$EndpointName' created and verified"; return $true }
        Add-Action 'Create migration endpoint' 'Failed' 'The endpoint did not appear after creation'
        return $false
    }
    catch {
        Add-Finding Critical 'Endpoint' 'Creating the migration endpoint failed.' $_.Exception.Message 'Almost always domain-wide delegation or the scope string. Re-run Preflight.'
        return $false
    }
}

function Invoke-CreateBatch {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    Assert-NoBlockers -Because 'create the migration batch'

    if (-not $script:ValidatedCsv) {
        Add-Finding Critical 'Batch' 'No validated CSV is available.' '' 'Run Preflight first.'
        return $false
    }
    $existing = $null
    try { $existing = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue } catch { }
    if ($existing) {
        Add-Action 'Create migration batch' 'NotNeeded' "'$BatchName' already exists with status $($existing.Status)"
        return $true
    }

    Write-Host ''
    Write-Warn2 "About to create a batch for $($script:Mappings.Count) user(s) into $TargetDeliveryDomain."
    Write-Info  'It will NOT be started automatically. You review it, then start it.'
    if (-not $PSCmdlet.ShouldProcess($BatchName, "Create migration batch for $($script:Mappings.Count) users")) {
        Add-Action 'Create migration batch' 'Skipped' 'Declined'
        return $false
    }

    try {
        $data = [System.IO.File]::ReadAllBytes($script:ValidatedCsv)
        $null = Invoke-ExoWithRetry -Label 'New-MigrationBatch' -Script {
            New-MigrationBatch -Name $BatchName -SourceEndpoint $EndpointName -CSVData $data `
                -TargetDeliveryDomain $TargetDeliveryDomain -AutoComplete:$false -ErrorAction Stop
        }
        $check = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
        if ($check) {
            Add-Action 'Create migration batch' 'Applied' "'$BatchName' created (status $($check.Status)) - NOT started"
            Write-Host ''
            Write-Info "Review it, then run:  -Mode StartBatch -BatchName $BatchName"
            return $true
        }
        Add-Action 'Create migration batch' 'Failed' 'The batch did not appear after creation'
        return $false
    }
    catch {
        Add-Finding Critical 'Batch' 'Creating the migration batch failed.' $_.Exception.Message 'Check the target delivery domain and that all targets are MailUsers.'
        return $false
    }
}

function Invoke-StartBatch {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $b = $null
    try { $b = Get-MigrationBatch -Identity $BatchName -ErrorAction Stop }
    catch { Add-Finding Critical 'Batch' "Batch '$BatchName' was not found." '' 'Create it first with -Mode CreateBatch.'; return $false }

    if ("$($b.Status)" -notin @('Created','Stopped','Failed')) {
        Add-Action 'Start migration batch' 'NotNeeded' "Status is already $($b.Status)"
        return $true
    }
    if (-not $PSCmdlet.ShouldProcess($BatchName, 'Start the migration batch')) {
        Add-Action 'Start migration batch' 'Skipped' 'Declined'
        return $false
    }
    try {
        Invoke-ExoWithRetry -Label 'Start-MigrationBatch' -Script { Start-MigrationBatch -Identity $BatchName -ErrorAction Stop }
        Add-Action 'Start migration batch' 'Applied' "'$BatchName' started"
        return $true
    }
    catch {
        Add-Finding Critical 'Batch' 'Starting the batch failed.' $_.Exception.Message ''
        return $false
    }
}

function Invoke-MonitorBatch {
    $terminal = @('Completed','CompletedWithErrors','Synced','SyncedWithErrors','Failed','Stopped','Corrupted','Removed')
    $deadline = (Get-Date).AddMinutes($MaxMonitorMinutes)
    $statsPath = Join-Path $script:RunFolder 'MigrationUserStatistics.csv'
    $reachedTerminal = $false
    $lastStatus = ''
    $cycle = 0

    Write-Head "MONITORING '$BatchName'"
    Write-Info "Refresh: ${MonitorIntervalSeconds}s   Stops after: ${MaxMonitorMinutes} min   Ctrl+C is safe"

    while ((Get-Date) -lt $deadline) {
        $cycle++
        $batch = $null
        try { $batch = Invoke-ExoWithRetry -Label 'Get-MigrationBatch' -Script { Get-MigrationBatch -Identity $BatchName -ErrorAction Stop } }
        catch {
            # Do not exit silently. Distinguish "gone" from "call failed".
            $msg = $_.Exception.Message
            if ($msg -match "couldn't be found|does not exist|NotFound") {
                Add-Finding High 'Batch' "Batch '$BatchName' no longer exists." $msg 'It was removed. Nothing further to monitor.'
                break
            }
            Write-Warn2 "Status read failed: $msg - reconnecting"
            $script:Connected = $false
            if (-not (Connect-ExchangeSafely)) {
                Add-Finding High 'Batch' 'Lost the Exchange session and could not reconnect.' '' 'Re-run with -Mode Monitor.'
                break
            }
            continue
        }

        if ("$($batch.Status)" -ne $lastStatus) {
            $lastStatus = "$($batch.Status)"
            Write-Host ''
            Write-Info "Status: $lastStatus   (cycle $cycle)"
        }

        $users = @()
        try { $users = @(Invoke-ExoWithRetry -Label 'Get-MigrationUser' -Script { Get-MigrationUser -BatchId $BatchName -ErrorAction Stop }) }
        catch { Write-Step 'User list unavailable this cycle.' }

        if ($users.Count -gt 0) {
            $g = $users | Group-Object Status
            $summary = (@($g | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join '  ')
            Write-Host "    $summary" -ForegroundColor Gray

            $problem = @($users | Where-Object { "$($_.Status)" -in @('Failed','Corrupted','IssueWarning') })
            foreach ($p in @($problem | Select-Object -First 5)) {
                Write-Warn2 "  $($p.Identity): $($p.Status) - $($p.ErrorSummary)"
            }
        }

        if ("$($batch.Status)" -in $terminal) { $reachedTerminal = $true; break }
        Start-Sleep -Seconds $MonitorIntervalSeconds
    }

    # Per-user statistics once, at the end - not every cycle. Appending keeps history.
    if ($script:Connected) {
        try {
            $users = @(Invoke-ExoWithRetry -Label 'Get-MigrationUser' -Script { Get-MigrationUser -BatchId $BatchName -ErrorAction Stop })
            $rows = New-Object System.Collections.ArrayList
            foreach ($u in $users) {
                $st = $null
                try { $st = Get-MigrationUserStatistics -Identity $u.Identity -ErrorAction SilentlyContinue } catch { }
                $bytes = ''
                if ($st -and $null -ne $st.BytesTransferred) { $bytes = "$($st.BytesTransferred)" }
                $synced = ''
                if ($st -and $null -ne $st.SyncedItemCount) { $synced = "$($st.SyncedItemCount)" }
                $skipped = ''
                if ($st -and $null -ne $st.SkippedItemCount) { $skipped = "$($st.SkippedItemCount)" }
                $null = $rows.Add([PSCustomObject]@{
                    Timestamp = (Get-Date).ToString('s'); Identity = "$($u.Identity)"
                    Status = "$($u.Status)"; SyncedItems = $synced; SkippedItems = $skipped
                    BytesTransferred = $bytes; Error = "$($u.ErrorSummary)"
                })
            }
            if ($rows.Count -gt 0) {
                $rows | Export-Csv -LiteralPath $statsPath -NoTypeInformation -Append -Encoding UTF8
                Add-Action 'Per-user statistics' 'Applied' $statsPath
            }
        }
        catch { Write-Step 'Final statistics could not be collected.' }
    }

    if ($reachedTerminal) {
        Add-Action 'Monitor batch' 'Applied' "Reached terminal status: $lastStatus"
        if ($lastStatus -in @('CompletedWithErrors','SyncedWithErrors','Failed','Corrupted')) {
            Add-Finding High 'Batch' "The batch finished as '$lastStatus'." '' "Review $statsPath for the per-user reason before completing."
        }
    }
    else {
        Add-Action 'Monitor batch' 'Skipped' "Monitoring window of ${MaxMonitorMinutes} min elapsed - batch still '$lastStatus'"
        Write-Warn2 'The batch did NOT finish. It is still running server-side.'
        Write-Info  "Resume with: -Mode Monitor -BatchName $BatchName"
    }
}

function Invoke-CompleteBatch {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    if (-not $ApproveCutover) {
        Add-Action 'Complete migration batch' 'Skipped' 'Requires -ApproveCutover - completion redirects live mail flow'
        Write-Warn2 'Completion is the cutover. Re-run with -ApproveCutover when ready.'
        return $false
    }
    $b = $null
    try { $b = Get-MigrationBatch -Identity $BatchName -ErrorAction Stop }
    catch { Add-Finding Critical 'Batch' "Batch '$BatchName' was not found." '' ''; return $false }

    if ("$($b.Status)" -notin @('Synced','SyncedWithErrors')) {
        Add-Finding High 'Batch' "Batch status is '$($b.Status)', not Synced." '' 'Only complete a batch once it has fully synced.'
        return $false
    }
    if ("$($b.Status)" -eq 'SyncedWithErrors') {
        Write-Warn2 'This batch synced WITH ERRORS. Some users may be incomplete.'
    }
    if (-not $PSCmdlet.ShouldProcess($BatchName, 'COMPLETE the batch - this redirects live mail flow')) {
        Add-Action 'Complete migration batch' 'Skipped' 'Declined'
        return $false
    }
    try {
        Invoke-ExoWithRetry -Label 'Complete-MigrationBatch' -Script { Complete-MigrationBatch -Identity $BatchName -ErrorAction Stop }
        Add-Action 'Complete migration batch' 'Applied' 'Cutover started'
        return $true
    }
    catch {
        Add-Finding Critical 'Batch' 'Completing the batch failed.' $_.Exception.Message ''
        return $false
    }
}

# ----------------------------------------------------------------------------------
# REPORTING
# ----------------------------------------------------------------------------------
function Write-Report {
    if (-not $script:RunFolder) { return }

    if ($script:Findings.Count -gt 0) {
        try { $script:Findings | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'Findings.csv') -NoTypeInformation -Encoding UTF8 } catch { }
    }
    if ($script:Actions.Count -gt 0) {
        try { $script:Actions | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'Actions.csv') -NoTypeInformation -Encoding UTF8 } catch { }
    }

    Write-Head 'RESULT'

    $crit = @($script:Findings | Where-Object { $_.Severity -eq 'Critical' })
    $high = @($script:Findings | Where-Object { $_.Severity -eq 'High' })
    $med  = @($script:Findings | Where-Object { $_.Severity -eq 'Medium' })

    if ($crit.Count -eq 0 -and $high.Count -eq 0) {
        Write-Host '  No blocking problems found.' -ForegroundColor Green
    }
    else {
        Write-Host "  $($crit.Count) critical, $($high.Count) high - these BLOCK the migration:" -ForegroundColor Red
        Write-Host ''
        foreach ($f in @($crit + $high)) {
            Write-Host "   * [$($f.Area)] $($f.Title)" -ForegroundColor Red
            if ($f.Remedy) { Write-Host "     FIX: $($f.Remedy)" -ForegroundColor Gray }
        }
    }
    if ($med.Count -gt 0) {
        Write-Host ''
        Write-Host "  $($med.Count) warning(s) - review but not blocking:" -ForegroundColor Yellow
        foreach ($f in $med) { Write-Host "   * [$($f.Area)] $($f.Title)" -ForegroundColor Yellow }
    }

    $applied = @($script:Actions | Where-Object { $_.Result -in @('Applied','Verified') })
    if ($applied.Count -gt 0) {
        Write-Host ''
        Write-Host '  VERIFIED / APPLIED' -ForegroundColor Cyan
        foreach ($a in $applied) { Write-Host "   - $($a.Name): $($a.Detail)" -ForegroundColor Green }
    }

    Write-Host ''
    Write-Host '  NEXT STEP' -ForegroundColor Cyan
    if ($crit.Count -gt 0 -or $high.Count -gt 0) {
        Write-Host '   Fix the blocking items above, then re-run Preflight.' -ForegroundColor Yellow
    }
    else {
        switch ($Mode) {
            'Preflight'      { Write-Host '   Both sides verified. Run -Mode CreateEndpoint, then -Mode CreateBatch.' -ForegroundColor Green }
            'CreateEndpoint' { Write-Host "   Run: -Mode CreateBatch -BatchName $BatchName" -ForegroundColor Green }
            'CreateBatch'    { Write-Host "   Review the batch, then: -Mode StartBatch -BatchName $BatchName" -ForegroundColor Green }
            'StartBatch'     { Write-Host "   Run: -Mode Monitor -BatchName $BatchName" -ForegroundColor Green }
            'Monitor'        { Write-Host "   When Synced: -Mode Complete -BatchName $BatchName -ApproveCutover" -ForegroundColor Green }
            'Complete'       { Write-Host '   Cutover started. Verify mail flow and re-enable retention policies.' -ForegroundColor Green }
            default          { Write-Host '   See the actions above.' -ForegroundColor Green }
        }
    }

    Write-Host ''
    Write-Host "  Report folder: $($script:RunFolder)" -ForegroundColor Cyan
    $mins = [math]::Round(((Get-Date) - $script:StartTime).TotalMinutes, 1)
    Write-Host "  Elapsed: $mins min" -ForegroundColor DarkGray
}

# ----------------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------------
function Invoke-Main {
    # A wizard is only possible with a real console attached.
    $interactive = (-not $NonInteractive) -and
                   ([Environment]::UserInteractive) -and
                   (-not [Console]::IsInputRedirected) -and
                   (-not $script:ModeWasSpecified)

    if ($interactive) {
        $w = Invoke-Wizard
        if (-not $w) { Write-Warn2 'Cancelled - nothing was changed.'; return }
        # Write to SCRIPT scope. Assigning bare names here would create locals and
        # leave every called function reading the original parameter values.
        Set-Variable -Name Mode                 -Value $w.Mode          -Scope Script
        Set-Variable -Name KeyPath              -Value $w.KeyPath       -Scope Script
        Set-Variable -Name GoogleAdminEmail     -Value $w.Admin         -Scope Script
        Set-Variable -Name CsvPath              -Value $w.Csv           -Scope Script
        Set-Variable -Name TargetDeliveryDomain -Value $w.Target        -Scope Script
        Set-Variable -Name EndpointName         -Value $w.Endpoint      -Scope Script
        Set-Variable -Name BatchName            -Value $w.Batch         -Scope Script
        Set-Variable -Name ApproveCutover       -Value ([bool]$w.Cutover) -Scope Script
        $script:GoogleRouting       = $w.GoogleRouting
        $script:WaitForDelegation   = $w.WaitForDelegation
        $script:AutoCreateMailUsers = $w.AutoCreateMailUsers
        $Mode = $w.Mode
    }

    # Only these modes are permitted to change anything in the tenant.
    $script:AllowMutation = ($Mode -in @('Run','CreateEndpoint','CreateBatch','StartBatch','Complete'))

    Write-Host ''
    Write-Host '  Google Workspace to Microsoft 365 migration' -ForegroundColor Cyan
    Write-Host "  Version $($script:Version)  |  Mode: $Mode  |  No Google SDK required" -ForegroundColor DarkGray

    if ($Mode -eq 'Guide') { Show-GoogleSetupGuide; return }

    $run = New-RunFolder
    Write-Info "Report folder: $run"
    try { Start-Transcript -Path (Join-Path $run 'transcript.txt') -Force | Out-Null } catch { }

    $keyObj = $null
    try {
        # ---- Load and validate inputs ------------------------------------------
        if ($Mode -in @('Preflight','CreateEndpoint','CreateBatch','Run')) {
            Write-Head 'VALIDATING INPUTS'
            $keyObj = Import-ServiceAccountKey -Path $KeyPath
            if ($keyObj) {
                try { $script:KeyBytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $KeyPath).Path) }
                catch { Add-Finding Critical 'Key' 'The key file could not be read as bytes.' $_.Exception.Message '' }
            }
            if ($CsvPath -eq '<auto>') {
                Write-Step 'Building the user list from mail users in Microsoft 365'
                if (Connect-ExchangeSafely) {
                    $discovered = @(Get-ExistingMailUsers)
                    if ($discovered.Count -eq 0) {
                        Add-Finding Critical 'CSV' 'No mail users were found in Exchange Online.' '' 'Provision mail users first, or supply a CSV with -CsvPath.'
                    }
                    else {
                        Write-Good "$($discovered.Count) mail user(s) discovered"
                        $script:Mappings = $discovered
                        $CsvPath = $null
                    }
                }
            }
            else {
                $script:Mappings = @(Import-MigrationCsv -Path $CsvPath)
            }
            if ($script:Mappings.Count -gt 0) {
                $script:ValidatedCsv = Export-MigrationCsvNoBom -Mappings $script:Mappings
            }
        }

        # ---- Google side --------------------------------------------------------
        if ($Mode -in @('Preflight','Run')) {
            Write-Head 'GOOGLE SIDE'
            Test-GooglePrerequisites -KeyObject $keyObj -Mappings $script:Mappings
        }

        # ---- Exchange side ------------------------------------------------------
        if ($Mode -ne 'Guide') {
            Write-Head 'MICROSOFT 365 SIDE'
            $null = Connect-ExchangeSafely
        }

        switch ($Mode) {
            'Preflight' {
                Test-ExchangePrerequisites -Mappings $script:Mappings
                Test-MigrationEndpointHealth -KeyObject $keyObj
            }
            'Run' {
                Test-ExchangePrerequisites -Mappings $script:Mappings
                Test-MigrationEndpointHealth -KeyObject $keyObj
                Assert-NoBlockers -Because 'continue past preflight'
                if (Invoke-CreateEndpoint) { $null = Invoke-CreateBatch }
            }
            'CreateEndpoint' {
                Test-ExchangePrerequisites -Mappings $script:Mappings
                Test-MigrationEndpointHealth -KeyObject $keyObj
                $null = Invoke-CreateEndpoint
            }
            'CreateBatch' {
                Test-ExchangePrerequisites -Mappings $script:Mappings
                $null = Invoke-CreateBatch
            }
            'StartBatch' { $null = Invoke-StartBatch }
            'Monitor'    { Invoke-MonitorBatch }
            'Complete'   { $null = Invoke-CompleteBatch }
        }
    }
    catch {
        Add-Finding Critical 'Run' 'The run stopped.' $_.Exception.Message 'See the transcript in the report folder.'
    }
    finally {
        # The private key is password-equivalent. It never outlives the run.
        if ($script:KeyBytes) {
            [Array]::Clear($script:KeyBytes, 0, $script:KeyBytes.Length)
            $script:KeyBytes = $null
        }
        # Only shred the key once it has served its purpose - after the endpoint
        # exists. Deleting it after Preflight would break the next step.
        if ($Mode -in @('CreateEndpoint','Run') -and $KeyPath -and -not (Test-Blocking)) {
            Remove-KeyFileSecurely -Path $KeyPath
        }
        elseif ($Mode -eq 'Preflight' -and $KeyPath -and (Test-Path -LiteralPath $KeyPath)) {
            Write-Host ''
            Write-Warn2 'The service account key is still on disk. It is a password-equivalent secret.'
            Write-Info  "It is deleted automatically once the endpoint is created: $KeyPath"
        }
        if ($script:Connected) {
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
            $script:Connected = $false
        }
        Save-ResumeState
        Write-Report
        try { Stop-Transcript | Out-Null } catch { }
    }
}

# Entry point. A pasted script cannot answer its own prompts, so stop early.
try {
    if (-not $PSCommandPath) {
        Write-Host ''
        Write-Host '  Save this as a .ps1 file and run it - pasting into the console does not work.' -ForegroundColor Yellow
        Write-Host '    powershell -ExecutionPolicy Bypass -File .\Invoke-GoogleToM365Migration.ps1 -Mode Guide' -ForegroundColor Green
        Write-Host ''
    }
    else {
        Invoke-Main
        if ([Environment]::UserInteractive -and -not $NonInteractive -and -not [Console]::IsInputRedirected) {
            Write-Host ''
            Write-Host '  Press Enter to close.' -ForegroundColor DarkGray
            $null = Read-Host
        }
    }
}
catch {
    Write-Host ''
    Write-Host "  Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
    if ($script:RunFolder) { Write-Host "  Diagnostics: $($script:RunFolder)" -ForegroundColor Gray }
    try { Stop-Transcript | Out-Null } catch { }
}
