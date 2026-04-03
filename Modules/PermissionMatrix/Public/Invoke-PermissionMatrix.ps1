function Invoke-PermissionMatrix {
    <#
    .SYNOPSIS
        Internal orchestrator for the Permission Matrix pipeline.
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
            -Message "Invoke-PermissionMatrixInternalHC failed: $_" `
            -Category 'Runtime' `
            -SystemErrors ([ref]$systemErrors)
    }
    finally {
        # ================================================================
        # 3. END STAGE (Always runs to guarantee reporting/cleanup)
        # ================================================================
        if ($context) {
            Invoke-PermissionMatrixEndHC `
                -Context $context `
                -SystemErrors ([ref]$systemErrors)
        }

        # ================================================================
        # 4. FINAL RESULT
        # ================================================================
        if ($context -and $context.ExportedFiles) {
            $result.Files = $context.ExportedFiles
        }

        $result.Success = -not (Test-HasFatalErrorsHC ([ref]$systemErrors))
        $result
    }
}