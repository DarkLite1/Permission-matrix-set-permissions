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
    function Import-PermissionMatrixModuleHC {
        param(
            [Parameter(Mandatory)][string]$ScriptRoot,
            [Parameter()][ref]$SystemErrors
        )

        try {
            $localModuleRoot = Join-Path $ScriptRoot 'Modules\Toolbox.PermissionMatrixHC'

            if (Test-Path $localModuleRoot) {
                Write-Verbose "Importing local module: $localModuleRoot"
                Import-Module $localModuleRoot -Force -ErrorAction Stop
                return
            }
        }
        catch {
            $msg = "Failed to import Toolbox.PermissionMatrixHC module: $_"

            if ($SystemErrors) {
                $SystemErrors.Value.Add([pscustomobject]@{
                        DateTime = Get-Date
                        Message  = $msg
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
# process not used — preserved to support pipelined input in the future
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