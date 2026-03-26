function Export-PermissionMatrix {
    <#
    .SYNOPSIS
        Runs the Permission Matrix export pipeline only.

    .DESCRIPTION
        This function loads and validates the configuration, loads matrix files,
        builds export data, and runs Export-Files. It does *not* generate HTML or 
        send email. Returns a hashtable with export results.

    .PARAMETER ConfigurationJsonFile
        Path to the JSON configuration file.

    .PARAMETER ScriptPath
        Hashtable containing script component paths (TestRequirementsFile, 
        SetPermissionFile, UpdateServiceNow).
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ConfigurationJsonFile,

        [Parameter(Mandatory)]
        [hashtable]$ScriptPath
    )

    $systemErrors = [System.Collections.Generic.List[object]]::new()
    $result = @{}

    try {
        # ------------------------------------------------------------------
        # 1. BEGIN stage (reuse internal private logic)
        # ------------------------------------------------------------------
        $context = Invoke-PermissionMatrixBegin `
            -ConfigurationJsonFile $ConfigurationJsonFile `
            -ScriptPath $ScriptPath `
            -SystemErrors ([ref]$systemErrors)

        if ($systemErrors.Count -gt 0) {
            Write-Warning "BEGIN stage produced $($systemErrors.Count) errors. Export aborted."
            return @{ Success = $false; Errors = $systemErrors }
        }

        # ------------------------------------------------------------------
        # 2. Load matrix files (PROCESS without HTML generation)
        # ------------------------------------------------------------------
        $importedMatrix = Invoke-PermissionMatrixProcess `
            -Context $context `
            -SystemErrors ([ref]$systemErrors)

        if (-not $importedMatrix) {
            Write-Warning 'No matrix files found to process.'
            return @{ Success = $false; Errors = $systemErrors }
        }

        # ------------------------------------------------------------------
        # 3. Validate configuration (cannot export without valid config)
        # ------------------------------------------------------------------
        $validation = Validate-Settings `
            -Settings $context.Settings `
            -Matrix $context.Matrix `
            -Export $context.Export `
            -ServiceNow $context.ServiceNow `
            -MaxConcurrent $context.MaxConcurrent

        foreach ($err in $validation.Errors) {
            $systemErrors.Add($err)
        }

        if (-not $validation.IsValid) {
            Write-Warning 'Configuration validation failed. Export aborted.'
            return @{ Success = $false; Errors = $systemErrors }
        }

        # ------------------------------------------------------------------
        # 4. Determine whether export is enabled
        # ------------------------------------------------------------------
        $exportWanted =
        $context.Export.ServiceNowFormDataExcelFile -or
        $context.Export.PermissionsExcelFile -or
        $context.Export.OverviewHtmlFile

        if (-not $exportWanted) {
            Write-Warning 'No export targets configured. Nothing to export.'
            return @{ Success = $true; Files = @{} }
        }

        # ------------------------------------------------------------------
        # 5. Build export data
        # ------------------------------------------------------------------
        $dataToExport = Build-ExportData `
            -ImportedMatrix $importedMatrix `
            -AdObjectHash $context.AdObjectHash `
            -GroupManagerHash $context.GroupManagerHash

        if (-not $dataToExport) {
            Write-Warning 'No export data produced.'
            return @{ Success = $true; Files = @{} }
        }

        # ------------------------------------------------------------------
        # 6. Run export
        # ------------------------------------------------------------------
        $exportLogFolder = Join-Path (Get-DatedLogFolderPathHC) 'Export'

        if (-not (Test-Path $exportLogFolder)) {
            New-Item -ItemType Directory -Path $exportLogFolder | Out-Null
        }

        $exportedFiles = Export-Files `
            -DataToExport $dataToExport `
            -ExportConfig $context.Export `
            -ServiceNowConfig $context.ServiceNow `
            -ExportLogFolder $exportLogFolder `
            -ScriptPathItem $context.ScriptPathItem `
            -SystemErrors ([ref]$systemErrors)

        # ------------------------------------------------------------------
        # 7. Final result structure
        # ------------------------------------------------------------------
        return @{
            Success = ($systemErrors.Count -eq 0)
            Files   = $exportedFiles
            Errors  = $systemErrors
        }
    }
    catch {
        $systemErrors.Add([pscustomobject]@{
                DateTime = Get-Date
                Message  = "Export-PermissionMatrix failed: $_"
            })

        return @{
            Success = $false
            Files   = @{}
            Errors  = $systemErrors
        }
    }
}