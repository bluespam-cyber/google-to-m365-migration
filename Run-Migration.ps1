#Requires -Version 5.1
<#
.SYNOPSIS
    Auto-locating launcher for GoogleToM365Migration.ps1 (Exchange Online migration).

.DESCRIPTION
    Automatically finds the migration controller script regardless of the current
    working directory. It locates itself via $PSScriptRoot and dispatches to
    scripts/GoogleToM365Migration.ps1, passing through all arguments unchanged.

.EXAMPLE
    .\Run-Migration.ps1 -Mode Preflight -CsvPath .\examples\migration-users.csv

.EXAMPLE
    .\Run-Migration.ps1 -Mode CreateEndpoint -ServiceAccountKeyPath C:\keys\gws.json -GoogleAdminEmail admin@contoso.com -ApproveGooglePrerequisites

.EXAMPLE
    .\Run-Migration.ps1 -Mode CreateBatch -CsvPath .\examples\migration-users.csv -TargetDeliveryDomain o365.contoso.com -ApproveGooglePrerequisites
#>
param()

$ErrorActionPreference = 'Stop'

# Auto-locate the scripts folder relative to this launcher.
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = Join-Path $scriptDir 'scripts\GoogleToM365Migration.ps1'

if (-not (Test-Path -LiteralPath $target)) {
    throw "Script not found: $target"
}

Write-Host "[INFO] Launcher located script at: $target" -ForegroundColor Cyan

# Pass through all arguments to the target script (preserves named parameters).
& $target @args
