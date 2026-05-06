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
        [hashtable]$ScriptPath,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    # ---------------------------------------------------------------------
    # 0. Shared state
    # ---------------------------------------------------------------------
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
                    # Swallow error: If the event log fallback fails, we don't want to crash the finally block
                }
            }
        }
        #endregion
    }
}