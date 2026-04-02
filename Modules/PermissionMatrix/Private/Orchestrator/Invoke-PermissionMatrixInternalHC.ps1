function Invoke-PermissionMatrixInternalHC {
    <#
    .SYNOPSIS
        Internal orchestrator for the Permission Matrix pipeline.

    .DESCRIPTION
        Executes the Permission Matrix pipeline in a controlled manner.

        Rules:
        - BEGIN may block PROCESS but must not suppress END
        - END runs only if there is something to report:
            * errors exist OR
            * matrix files were found
        - Idle polling runs are silent
        - PROCESS may run in parallel
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
    $importedMatrix = @()

    $canProcess = $true
    $foundMatrices = $false
    $mustReport = $false

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
        
        if ($systemErrors.Count -gt 0) {
            $canProcess = $false
            $mustReport = $true
        }

        # ================================================================
        # 2. PROCESS STAGE
        # ================================================================
        if ($canProcess) {
            $processResult = Invoke-PermissionMatrixProcessHC `
                -Context $context `
                -SystemErrors ([ref]$systemErrors)

            if ($processResult) {
                $foundMatrices = $processResult.FoundMatrices
                $importedMatrix = $processResult.Imported
            }

            if ($foundMatrices) {
                $mustReport = $true
            }

            if (Test-HasFatalErrorsHC ([ref]$systemErrors)) {
                $canProcess = $false
                $mustReport = $true
            }
        }

        # ================================================================
        # 3. EXPORT-ONLY MODE (conditional)
        # ================================================================
        if ($canProcess -and $foundMatrices) {

            $exportWanted =
            $context.Export.PermissionsExcelFile -or
            $context.Export.ServiceNowFormDataExcelFile -or
            $context.Export.OverviewHtmlFile

            if ($exportWanted) {
                $result.Files = Export-FilesHC `
                    -ImportedMatrix $importedMatrix `
                    -ExportSettings $context.Export `
                    -HtmlOverview $null `
                    -Counters $context.Counter
            }
        }
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Name 'Unhandled orchestrator failure' `
            -Message "Invoke-PermissionMatrixInternalHC failed: $_" `
            -Category 'Runtime' `
            -SystemErrors ([ref]$systemErrors)

        $mustReport = $true
    }
    finally {
        # ================================================================
        # 4. END STAGE (conditional but guaranteed for errors)
        # ================================================================
        if ($mustReport) {
            try {
                Invoke-PermissionMatrixEndHC `
                    -Context $context `
                    -ImportedMatrix $importedMatrix `
                    -SystemErrors ([ref]$systemErrors)
            }
            catch {
                Add-ErrorHC `
                    -Type 'FatalError' `
                    -Name 'END stage failure' `
                    -Message $_ `
                    -Category 'Runtime' `
                    -SystemErrors ([ref]$systemErrors)
            }
        }

        # ================================================================
        # 5. FINAL RESULT
        # ================================================================
        if ($context -and $context.ExportedFiles) {
            $result.Files = $context.ExportedFiles
        }

        $result.Success = -not (Test-HasFatalErrorsHC ([ref]$systemErrors))
        $result
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
        -ScriptPath $ScriptPath
}
