function Invoke-PermissionMatrixInternalHC {
    <#
    .SYNOPSIS
        Internal orchestrator for the Permission Matrix pipeline.

    .DESCRIPTION
        Executes the Permission Matrix pipeline in a controlled manner.
        Supports multiple execution modes:
            - Full        : BEGIN → PROCESS → END (HTML, mail, logging)
            - ExportOnly  : BEGIN → PROCESS → EXPORT only

        This function owns:
            - SystemErrors lifecycle
            - Fatal error short-circuiting
            - Mode-based execution control

        Public functions must call this function and do nothing else.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ConfigurationJsonFile,

        [Parameter(Mandatory)]
        [hashtable]$ScriptPath,

        [Parameter(Mandatory)]
        [ValidateSet('Full', 'ExportOnly')]
        [string]$Mode
    )

    # ---------------------------------------------------------------------
    # 0. Shared state
    # ---------------------------------------------------------------------
    $systemErrors = [System.Collections.Generic.List[object]]::new()
    $result = [ordered]@{
        Success = $false
        Mode    = $Mode
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

        if (Test-HasFatalErrorsHC ([ref]$systemErrors)) {
            return $result
        }


        # ================================================================
        # 2. PROCESS STAGE
        # ================================================================
        $importedMatrix = Invoke-PermissionMatrixProcessHC `
            -Context $context `
            -SystemErrors ([ref]$systemErrors)

        if (-not $importedMatrix) {
            Add-ErrorHC `
                -Type 'Warning' `
                -Name 'No matrix files processed' `
                -Message 'No matrix files were found or successfully processed.' `
                -Category 'Matrix' `
                -SystemErrors ([ref]$systemErrors)
        }

        if (Test-HasFatalErrorsHC ([ref]$systemErrors)) {
            return $result
        }


        # ================================================================
        # 3. EXPORT-ONLY MODE
        # ================================================================
        if ($Mode -eq 'ExportOnly') {

            # Determine if export is actually configured
            $exportWanted =
            $context.Export.PermissionsExcelFile -or
            $context.Export.ServiceNowFormDataExcelFile -or
            $context.Export.OverviewHtmlFile

            if (-not $exportWanted) {
                Add-ErrorHC `
                    -Type 'Information' `
                    -Name 'No export configured' `
                    -Message 'No export targets configured. Nothing to export.' `
                    -Category 'RuntimeSettings' `
                    -SystemErrors ([ref]$systemErrors)

                $result.Success = -not (Test-HasFatalErrorsHC ([ref]$systemErrors))
                return $result
            }

            # Run export pipeline
            $exportedFiles = Export-FilesHC `
                -ImportedMatrix $importedMatrix `
                -ExportSettings $context.Export `
                -HtmlOverview $null `
                -Counters $context.Counter

            $result.Files = $exportedFiles
            $result.Success = -not (Test-HasFatalErrorsHC ([ref]$systemErrors))
            return $result
        }


        # ================================================================
        # 4. END STAGE (FULL PIPELINE ONLY)
        # ================================================================
        Invoke-PermissionMatrixEnd `
            -Context $context `
            -ImportedMatrix $importedMatrix `
            -SystemErrors ([ref]$systemErrors)


        # ================================================================
        # 5. FINAL RESULT
        # ================================================================
        $result.Files = $context.ExportedFiles ?? @{}
        $result.Success = -not (Test-HasFatalErrorsHC ([ref]$systemErrors))
        return $result
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Name 'Unhandled orchestrator failure' `
            -Message "Invoke-PermissionMatrixInternalHC failed: $_" `
            -Category 'Runtime' `
            -SystemErrors ([ref]$systemErrors)

        return $result
    }
}

function Invoke-PermissionMatrix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConfigurationJsonFile,
        [Parameter(Mandatory)][hashtable]$ScriptPath
    )

    Invoke-PermissionMatrixInternalHC `
        -ConfigurationJsonFile $ConfigurationJsonFile `
        -ScriptPath $ScriptPath `
        -Mode 'Full'
}

function Export-PermissionMatrix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConfigurationJsonFile,
        [Parameter(Mandatory)][hashtable]$ScriptPath
    )

    Invoke-PermissionMatrixInternalHC `
        -ConfigurationJsonFile $ConfigurationJsonFile `
        -ScriptPath $ScriptPath `
        -Mode 'ExportOnly'
}