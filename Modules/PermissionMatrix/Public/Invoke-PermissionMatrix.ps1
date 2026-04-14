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

    $result = [ordered]@{
        Success = $false
        Files   = @{}
        Errors  = $systemErrors
    }

    try {
        # ================================================================
        # 1. BEGIN STAGE
        # ================================================================
        $context = Invoke-PermissionMatrixBeginHC `
            -ConfigurationJsonFile $ConfigurationJsonFile `
            -ScriptPath $ScriptPath `
            -SystemErrors ([ref]$systemErrors)

        $hasFatal = Test-HasFatalErrorsHC ([ref]$systemErrors)

        # ================================================================
        # 2. PROCESS STAGE
        # ================================================================
        if ($context -and $context.FoundMatrices -and -not $hasFatal) {
            
            $context = Invoke-PermissionMatrixProcessHC `
                -Context $context `
                -SystemErrors ([ref]$systemErrors)
        }
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
        # ================================================================
        # 3. END STAGE (Always runs to guarantee reporting/cleanup)
        # ================================================================
        if ($context) {
            # We have at least a partial configuration! 
            # Run the full END stage (Emails, Logs, ServiceNow)
            Invoke-PermissionMatrixEndHC `
                -Context $context `
                -SystemErrors ([ref]$systemErrors)
        }
        elseif ($systemErrors.Count -gt 0) {
            # FATAL EARLY ERROR (e.g., JSON file missing/corrupt). 
            # We have NO email or log folder configuration, so we must fall back to the Host/EventLog.
            foreach ($err in $systemErrors) {
                $msg = "[$($err.Type)] $($err.Name): $($err.Message) $($err.Description)"
                
                # Output to the PowerShell Host (Standard Error Stream) 
                Write-Error $msg
                
                # Fallback Event Log write (using a hardcoded source since we couldn't read the config) 
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

        # ================================================================
        # 4. FINAL RESULT
        # ================================================================
        if ($context -and $context.ExportedFiles.Count -gt 0) {
            $result.Files = $context.ExportedFiles
        }

        $result.Success = -not (Test-HasFatalErrorsHC ([ref]$systemErrors))
        $result
    }
}