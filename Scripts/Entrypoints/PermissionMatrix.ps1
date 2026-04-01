<#
.SYNOPSIS
    Permission Matrix Orchestrator Script
.DESCRIPTION
    This script loads the Toolbox.PermissionMatrixHC module safely (preferring the
    local copy in ./Modules) and invokes the main Permission Matrix operation using
    the public entrypoint: Invoke-PermissionMatrix.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigurationJsonFile,

    [Parameter(Mandatory)]
    [hashtable]$ScriptPath
)

begin {
    try {
        function Import-PermissionMatrixModuleHC {
            param(
                [Parameter(Mandatory)][string]$ScriptRoot,
                [Parameter()][ref]$SystemErrors
            )

            try {
                $localModuleRoot = Join-Path $ScriptRoot '..\..\Modules\PermissionMatrix'

                if (Test-Path $localModuleRoot) {
                    Write-Verbose "Importing local module: $localModuleRoot"
                    Import-Module $localModuleRoot -Force -ErrorAction Stop
                    return
                }
            }
            catch {
                $msg = "Failed to import PermissionMatrixHC module: $_"
                if ($SystemErrors) {
                    $SystemErrors.Value.Add([pscustomobject]@{
                            DateTime = Get-Date
                            Message  = $msg
                            Type     = 'FatalError'
                            Category = 'Bootstrap'
                        })
                }
                throw $msg
            }
        }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Import-PermissionMatrixModuleHC `
            -ScriptRoot $PSScriptRoot `
            -SystemErrors ([ref]$systemErrors)
    }
    catch {
        Write-Error "BEGIN stage crashed before orchestrator could run: $_"
        exit 1
    }
}

process { }

end {
    try {
        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $ConfigurationJsonFile `
            -ScriptPath $ScriptPath
    }
    catch {
        Write-Error "Unhandled fatal error: $_"
        exit 1
    }
}