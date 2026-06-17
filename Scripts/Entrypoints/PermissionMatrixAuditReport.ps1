#Requires -Version 5.1

<#
.SYNOPSIS
    Entry point for the monthly permission matrix audit report.

.DESCRIPTION
    Thin wrapper that loads the PermissionMatrix module and runs the public
    orchestrator 'Invoke-PermissionMatrixAuditReport'. It is meant to run on a
    slow schedule (e.g. once a month) and is completely separate from the main
    'PermissionMatrix.ps1' entry point that applies ACLs every few minutes.

    All audit configuration lives in its own JSON file (see
    Examples\PermissionMatrixAuditReport.json), including a full
    'Settings.SendMail' block with its own From / Smtp / AssemblyPath.

.PARAMETER ConfigurationJsonFile
    Path to the audit report JSON configuration file.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigurationJsonFile
)

#region Load the module
# INTEGRATION: this mirrors the local Import-PermissionMatrixModuleHC used by
# PermissionMatrix.ps1. Use the exact same definition as the main entry point
# (here it simply imports the module manifest/root module).
function Import-PermissionMatrixModuleHC {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ModulePath)
    Import-Module -Name $ModulePath -Force -ErrorAction Stop
}

$root = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$modulePath = Join-Path $root 'Modules\PermissionMatrix\PermissionMatrix.psm1'

Import-PermissionMatrixModuleHC -ModulePath $modulePath
#endregion

#region Run the audit
# Only 'PermissionMatrixModule' is required: the audit runs the Begin stage
# only, and Begin uses this single path to dot-source the private functions
# inside its parallel runspaces. SetPermissions / UpdateServiceNow /
# TestRequirements are not needed because Process and End never run.
$scriptPath = @{
    PermissionMatrixModule = $modulePath
}

$systemErrors = [System.Collections.Generic.List[object]]::new()

Invoke-PermissionMatrixAuditReport `
    -ConfigurationJsonFile $ConfigurationJsonFile `
    -ScriptPath $scriptPath `
    -SystemErrors ([ref]$systemErrors)
#endregion

if ($systemErrors.Count -gt 0) {
    # The admin was already notified by the orchestrator; surface a non-zero
    # exit so the scheduled task reports the failure too.
    throw "Audit report initialization failed: $(@($systemErrors | ForEach-Object { $_.Message }) -join '; ')"
}