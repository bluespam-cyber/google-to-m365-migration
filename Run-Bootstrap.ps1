#Requires -Version 5.1
<#
.SYNOPSIS
    Auto-locating launcher for GoogleCloudMigrationBootstrap.ps1 (Google Cloud prep).

.DESCRIPTION
    Automatically finds the Google Cloud bootstrap script regardless of the current
    working directory. It locates itself via $PSScriptRoot and dispatches to
    scripts/GoogleCloudMigrationBootstrap.ps1, passing through all arguments unchanged.

.EXAMPLE
    .\Run-Bootstrap.ps1 -Mode Inspect

.EXAMPLE
    .\Run-Bootstrap.ps1 -Mode Inspect -ProjectId my-gcp-project

.EXAMPLE
    .\Run-Bootstrap.ps1 -Mode CreateServiceAccountKey -ProjectId my-gcp-project -ApproveKeyCreation
#>
param()

$ErrorActionPreference = 'Stop'

# Auto-locate the scripts folder relative to this launcher.
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = Join-Path $scriptDir 'scripts\GoogleCloudMigrationBootstrap.ps1'

if (-not (Test-Path -LiteralPath $target)) {
    throw "Script not found: $target"
}

Write-Host "[INFO] Launcher located script at: $target" -ForegroundColor Cyan

# Pass through all arguments to the target script (preserves named parameters).
& $target @args
