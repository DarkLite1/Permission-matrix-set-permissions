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
                [Parameter(Mandatory)][string]$Path,
                [Parameter()][ref]$SystemErrors
            )

            if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
                $SystemErrors.Value.Add(
                    [pscustomobject]@{
                        DateTime = Get-Date
                        Message  = "PermissionMatrix module not found at '$Path'."
                        Type     = 'FatalError'
                        Category = 'Bootstrap'
                    }
                )
                return
            }

            try {
                Write-Verbose "Importing PermissionMatrix module: $Path"
                Import-Module $Path -Force -ErrorAction Stop
            }
            catch {
                $SystemErrors.Value.Add(
                    [pscustomobject]@{
                        DateTime = Get-Date
                        Message  = "Failed to import PermissionMatrix module: $_"
                        Type     = 'FatalError'
                        Category = 'Bootstrap'
                    }
                )
            }
        }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        $modulePath = Join-Path $PSScriptRoot '..\..\Modules\PermissionMatrix\PermissionMatrix.psm1'

        $ScriptPath.PermissionMatrixModule = $modulePath

        Import-PermissionMatrixModuleHC `
            -Path $modulePath `
            -SystemErrors ([ref]$systemErrors)
    }
    catch {
        Write-Warning "BEGIN stage crashed before orchestrator could run: $_"
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
        Write-Warning "Unhandled fatal error: $_"
        exit 1
    }
}