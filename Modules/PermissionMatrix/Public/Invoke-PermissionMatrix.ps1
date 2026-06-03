function Invoke-PermissionMatrix {
    <#
    .SYNOPSIS
        Main orchestrator for the Permission Matrix pipeline.

    .DESCRIPTION
        Executes the Permission Matrix pipeline in a controlled, fault-tolerant 
        manner across three distinct stages:
        
        - BEGIN
            Validates the JSON configuration, imports and archives Excel 
            matrices, resolves Active Directory objects, and safely merges 
            global default permissions.

        - PROCESS
            Evaluates target machine requirements and executes multi-threaded, 
            remote permission assignments across specified servers and folders.

        - END
            Generates detailed HTML/Excel reports, synchronizes data with 
            ServiceNow (if configured), and dispatches automated SMTP email 
            notifications.

        If a catastrophic error prevents the pipeline from creating an 
        execution context (e.g., missing configuration files), the script 
        gracefully falls back to writing terminating errors directly to the 
        Windows Event Log ('Application' log, Source 'Permission Matrix').

    .PARAMETER ConfigurationJsonFile
        The absolute path to the main JSON configuration file governing the 
        execution (e.g., 'C:\PermissionMatrix\Config.json').

    .PARAMETER ScriptPath
        A hashtable containing the absolute file paths to the required 
        execution scripts. 
        Must contain the following keys: 
        - TestRequirementsFile
        - SetPermissionFile
        - UpdateServiceNow
        - PermissionMatrixModule

    .PARAMETER SystemErrors
        A reference variable ([ref]) containing a List[pscustomobject]. Used to 
        capture and bubble up terminating pipeline errors across all stages 
        without halting the primary orchestrator loop.

    .EXAMPLE
        $sysErrors = [System.Collections.Generic.List[pscustomobject]]::new()
        
        $scriptPaths = @{
            TestRequirementsFile   = 'C:\Scripts\Test requirements.ps1'
            SetPermissionFile      = 'C:\Scripts\Set permissions.ps1'
            UpdateServiceNow       = 'C:\Scripts\Update ServiceNow.ps1'
            PermissionMatrixModule = 'C:\Modules\PermissionMatrix\PermissionMatrix.psm1'
        }

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile 'C:\Config.json' `
            -ScriptPath $scriptPaths `
            -SystemErrors ([ref]$sysErrors)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ConfigurationJsonFile,

        [Parameter(Mandatory)]
        [hashtable]$ScriptPath,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    <# Shared variable to hold the execution context across stages. If this 
    remains null, it indicates a failure in the BEGIN stage that prevented 
    pipeline initialization. #>
    $context = $null

    try {
        #region Check for input errors
        $context = Invoke-PermissionMatrixBeginHC `
            -ConfigurationJsonFile $ConfigurationJsonFile `
            -ScriptPath $ScriptPath `
            -SystemErrors $SystemErrors
        #endregion

        $hasFatal = Test-ItemHasFatalErrorHC -CheckList $SystemErrors.Value

        #region Process matrix files
        if ($context -and $context.FoundMatrices -and -not $hasFatal) {
            $context = Invoke-PermissionMatrixProcessHC `
                -Context $context `
                -SystemErrors $SystemErrors
        }
        elseif ($context -and $hasFatal) {
            foreach ($matrixObj in $context.AllMatrices) {
                if (-not (Test-ItemHasFatalErrorHC -CheckList $matrixObj.Check)) {
                    $matrixObj.Check.Add(
                        [pscustomobject]@{
                            Type        = 'FatalError'
                            Name        = 'Run aborted'
                            Description = 'This matrix was not processed because a system-level error aborted the run. See the system errors for the cause.'
                            Value       = ''
                        }
                    )
                }
            }
        }
        #endregion
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Name 'Unhandled orchestrator failure' `
            -Message "Invoke-PermissionMatrix failed: $_" `
            -Category 'Runtime' `
            -SystemErrors $SystemErrors
    }
    finally {
        #region Report results via email and logs (best effort)
        if ($context) {
            Invoke-PermissionMatrixEndHC `
                -Context $context `
                -SystemErrors $SystemErrors
        }
        elseif ($systemErrors.Value.Count -gt 0) {
            foreach ($err in $systemErrors.Value) {
                $msg = "[$($err.Type)] $($err.Name): $($err.Message) $($err.Description)"
                
                Write-Error $msg
                
                try {
                    if (-not [System.Diagnostics.EventLog]::SourceExists('Permission Matrix')) {
                        New-EventLog `
                            -LogName 'Application' `
                            -Source 'Permission Matrix' `
                            -ErrorAction 'SilentlyContinue'
                    }
                    Write-EventLog `
                        -LogName 'Application' `
                        -Source 'Permission Matrix' `
                        -EntryType Error `
                        -EventId 2 `
                        -Message $msg `
                        -ErrorAction 'SilentlyContinue'
                }
                catch { 
                    <# Swallow error: If the event log fallback fails, 
                    we don't want to crash the finally block #>
                }
            }
        }
        #endregion
    }
}