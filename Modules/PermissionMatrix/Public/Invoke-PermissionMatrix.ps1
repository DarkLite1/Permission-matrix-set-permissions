function Invoke-PermissionMatrix {
    <#
    .SYNOPSIS
        Main orchestrator for the Permission Matrix pipeline.
    .DESCRIPTION
        Executes the Permission Matrix pipeline in a controlled manner:
        - BEGIN: Parses config, reads/archives matrices, checks AD.
        - PROCESS: Executes remote commands.
        - END: Exports files, logs, and sends email.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ConfigurationJsonFile,

        [Parameter(Mandatory)]
        [hashtable]$ScriptPath
    )

    # ---------------------------------------------------------------------
    # 0. Shared state
    # ---------------------------------------------------------------------
    $systemErrors = [System.Collections.Generic.List[object]]::new()
    $context = $null

    try {
        #region Check for input errors
        $context = Invoke-PermissionMatrixBeginHC `
            -ConfigurationJsonFile $ConfigurationJsonFile `
            -ScriptPath $ScriptPath `
            -SystemErrors ([ref]$systemErrors)
        #endregion

        $hasFatal = Test-HasFatalErrorsHC ([ref]$systemErrors)

        #region Process matrix files
        if ($context -and $context.FoundMatrices -and -not $hasFatal) {
            $context = Invoke-PermissionMatrixProcessHC `
                -Context $context `
                -SystemErrors ([ref]$systemErrors)
        }
        #endregion
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Name 'Unhandled orchestrator failure' `
            -Message "Invoke-PermissionMatrix failed: $_" `
            -Category 'Runtime' `
            -SystemErrors ([ref]$systemErrors)
    }
    finally {
        #region Report results via email and logs (best effort)
        if ($context) {
            Invoke-PermissionMatrixEndHC `
                -Context $context `
                -SystemErrors ([ref]$systemErrors)
        }
        elseif ($systemErrors.Count -gt 0) {
            foreach ($err in $systemErrors) {
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
                    # Swallow error: If the event log fallback fails, we don't want to crash the finally block
                }
            }
        }
        #endregion

        #region Final fatal error check and exit code
        if ($systemErrors.Count -gt 0) {
            $systemErrors | ForEach-Object {
                Write-Warning "Logged System Error: [$($_.Type)] $($_.Name) - $($_.Message)"
            }
        }

        if (Test-HasFatalErrorsHC ([ref]$systemErrors)) {
            Write-Warning 'Exit script with error code 1'
            
            throw 'Permission Matrix execution completed with fatal errors.'
        }
        else {
            Write-Verbose 'Script finished successfully'
        }
        #endregion
    }
}