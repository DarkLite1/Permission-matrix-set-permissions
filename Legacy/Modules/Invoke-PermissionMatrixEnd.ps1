function Invoke-PermissionMatrixEnd {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter()][array]$ImportedMatrix,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    try {
        # ------------------------------------------------------------
        # 1. VALIDATION
        # ------------------------------------------------------------
        Validate-RuntimeSettings `
            -Settings $Context.Settings `
            -Matrix $Context.Matrix `
            -Export $Context.Export `
            -ServiceNow $Context.ServiceNow `
            -MaxConcurrent $Context.MaxConcurrent
        
        if ($SystemErrors.Value.Type -contains 'FatalError') {
            Write-Warning 'Runtime settings validation failed. Aborting END block.'
            return
        }


        # ------------------------------------------------------------
        # 2. HTML STRUCTURE
        # ------------------------------------------------------------
        $html = Initialize-HtmlStructure


        # ------------------------------------------------------------
        # 3. MATRIX PROCESSING (per-file HTML + troubleshooting)
        # ------------------------------------------------------------
        if ($ImportedMatrix) {
            $ImportedMatrix = Process-MatrixObjects `
                -ImportedMatrix $ImportedMatrix `
                -Html $html

            $html.MatrixTables = Build-MatrixEmailHtml `
                -ImportedMatrix $ImportedMatrix `
                -Html $html
        }


        # ------------------------------------------------------------
        # 4. EXPORT DATA
        # ------------------------------------------------------------
        $dataToExport = $null
        $exportWanted =
        $Context.Export.ServiceNowFormDataExcelFile -or
        $Context.Export.PermissionsExcelFile -or
        $Context.Export.OverviewHtmlFile

        if ($exportWanted -and $ImportedMatrix) {
            $dataToExport = Build-ExportData `
                -ImportedMatrix $ImportedMatrix `
                -AdObjectHash $Context.AdObjectHash `
                -GroupManagerHash $Context.GroupManagerHash
        }


        # ------------------------------------------------------------
        # 5. EXPORT FILES
        # ------------------------------------------------------------
        $exportedFiles = @{}

        if ($dataToExport -and $SystemErrors.Value.Count -eq 0) {

            $exportLogFolder = Join-Path (Get-DatedLogFolderPathHC) 'Export'
            if (-not (Test-Path $exportLogFolder)) {
                New-Item -ItemType Directory -Path $exportLogFolder | Out-Null
            }

            $exportedFiles = Export-Files `
                -DataToExport $dataToExport `
                -ExportConfig $Context.Export `
                -ServiceNowConfig $Context.ServiceNow `
                -ExportLogFolder $exportLogFolder `
                -ScriptPathItem $Context.ScriptPathItem `
                -SystemErrors $SystemErrors
        }


        # ------------------------------------------------------------
        # 6. COUNTERS + SUMMARY TABLE
        # ------------------------------------------------------------
        $counter = Build-Counters `
            -ImportedMatrix $ImportedMatrix `
            -SystemErrors $SystemErrors

        $html.ErrorWarningTable = Build-ErrorWarningTable `
            -CounterData $counter `
            -SystemErrors $SystemErrors


        # ------------------------------------------------------------
        # 7. EVENTS
        # ------------------------------------------------------------
        Write-EventLogSafe `
            -EventLogData $Context.EventLogData `
            -ScriptName $Context.Settings.ScriptName `
            -Settings $Context.Settings `
            -SystemErrors $SystemErrors


        # ------------------------------------------------------------
        # 8. LOG CLEANUP + ERROR LOG
        # ------------------------------------------------------------
        Cleanup-OldLogs `
            -LogFolder $Context.LogFolder `
            -RetentionDays $Context.Settings.SaveLogFiles.DeleteLogsAfterDays `
            -SystemErrors $SystemErrors

        Write-SystemErrorLog `
            -SystemErrors $SystemErrors `
            -LogFolder $Context.LogFolder `
            -MailParams ([ref]$Context.MailParams)


        # ------------------------------------------------------------
        # 9. MAIL
        # ------------------------------------------------------------
        $Context.MailParams = Build-MailParameters `
            -Settings $Context.Settings `
            -Html $html `
            -ExportedFiles $exportedFiles `
            -Counter $counter `
            -SystemErrors $SystemErrors `
            -MatrixCount @($ImportedMatrix).Count `
            -ExistingMailParams $Context.MailParams `
            -MailToDefaultsFile $Context.Settings.SendMail.ToDefaults `
            -LogFolder $Context.LogFolder `
            -ScriptStartTime $Context.ScriptStartTime

        Send-MailSafe `
            -MailParams $Context.MailParams `
            -SystemErrors $SystemErrors

        Save-MailBodyToLog `
            -MailParams $Context.MailParams `
            -LogFolder $Context.LogFolder `
            -SystemErrors $SystemErrors

    }
    catch {
        $SystemErrors.Value.Add([pscustomobject]@{
                DateTime = Get-Date
                Message  = "Unhandled error in END: $_"
            })
    }
}