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
    [string]$ConfigurationJsonFile
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

        # Errors raised before the orchestrator can run (module load failures).
        # These cannot flow through Invoke-PermissionMatrixEndHC because that
        # function is defined inside the module that just failed to load.
        $bootstrapErrors = [System.Collections.Generic.List[object]]::new()

        # Errors raised by the orchestrator itself (config, AD, remote jobs).
        # Passed by ref so Begin/Process/End all write to the same list.
        $runtimeErrors = [System.Collections.Generic.List[object]]::new()

        $projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $modulePath = Join-Path $projectRoot 'Modules\PermissionMatrix\PermissionMatrix.psm1'
        $opsRoot = Join-Path $projectRoot 'Operations'

        $ScriptPath = @{
            PermissionMatrixModule = $modulePath
            SetPermissions         = Join-Path $opsRoot 'SetPermissions.ps1'
            TestRequirements       = Join-Path $opsRoot 'TestRequirements.ps1'
            UpdateServiceNow       = Join-Path $opsRoot 'UpdateServiceNow.ps1'
        }

        Import-PermissionMatrixModuleHC `
            -Path $modulePath `
            -SystemErrors ([ref]$bootstrapErrors)
    }
    catch {
        Write-Warning "BEGIN stage crashed before orchestrator could run: $_"
        exit 1
    }
}

process { }

end {
    # Bootstrap failed — report locally and bail
    # orchestrator functions are not available
    if ($bootstrapErrors.Count -gt 0) {
        foreach ($err in $bootstrapErrors) {
            Write-Warning "[$($err.Category)] $($err.Message)"
        }
        exit 1
    }
    
    try {
        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $ConfigurationJsonFile `
            -ScriptPath $ScriptPath `
            -SystemErrors ([ref]$runtimeErrors)
    }
    catch {
        Write-Warning "Unhandled fatal error: $_"
        exit 1
    }

    if ($runtimeErrors.Count -gt 0) {
        foreach ($err in $runtimeErrors) {
            Write-Warning "[$($err.Type)] $($err.Name): $($err.Message)"
        }

        if ($runtimeErrors | Where-Object { $_.Type -eq 'FatalError' }) {
            exit 1
        }
    }
}