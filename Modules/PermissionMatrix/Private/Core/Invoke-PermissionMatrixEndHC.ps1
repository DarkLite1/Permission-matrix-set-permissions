function Invoke-PermissionMatrixEndHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Context,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    # Tally up all errors and warnings generated across the runspaces
    $Context.Counter = Update-MatrixCounterHC `
        -Context $Context `
        -SystemErrors $SystemErrors.Value

    $hasFatalErrors = Test-HasFatalErrorsHC $SystemErrors
    $htmlTemplates = Initialize-HtmlStructureHC
    $fullHtmlBody = ''

    # =====================================================================
    # 1. BUILD HTML BODY (Best Effort)
    # =====================================================================
    try {
        $matrixHtml = if ($Context.FoundMatrices) { Build-MatrixEmailHtmlHC -ImportedMatrix $Context.Matrices -Html $htmlTemplates } else { '' }
        $fullHtmlBody = Generate-MailBodyHtmlHC `
            -Settings $Context.Settings `
            -Html @{ Style = $htmlTemplates.Style; ErrorWarningTable = (Build-ErrorWarningTableHC -CounterData $Context.Counter -SystemErrors $SystemErrors); MatrixTables = $matrixHtml } `
            -ScriptStartTime $Context.StartTime
    }
    catch {
        Add-ErrorHC -Type 'Warning' -Name 'HTML Generation' -Message "Failed to build HTML body: $_" -Category 'Reporting' -SystemErrors $SystemErrors
    }

    # =====================================================================
    # 2. EXPORTS & SERVICENOW (Skip if Fatal Errors)
    # =====================================================================
    if (-not $hasFatalErrors -and $Context.FoundMatrices) {
        try {
            $Context.ExportedFiles = Export-FilesHC -ImportedMatrix $Context.Matrices -ExportSettings $Context.Export -HtmlOverview $fullHtmlBody -Counters $Context.Counter
            
            if ($Context.Export.ServiceNowFormDataExcelFile -and $Context.ServiceNow.CredentialsFilePath) {
                $snowParams = @{ CredentialsFilePath = $Context.ServiceNow.CredentialsFilePath; Environment = $Context.ServiceNow.Environment; TableName = $Context.ServiceNow.TableName; FormDataExcelFilePath = $Context.Export.ServiceNowFormDataExcelFile; ExcelFileWorksheetName = 'SnowFormData' }
                & $Context.ScriptPath.UpdateServiceNow @snowParams
            }
        }
        catch {
            Add-ErrorHC -Type 'Warning' -Name 'Exports/ServiceNow' -Message "Failed during export phase: $_" -Category 'Reporting' -SystemErrors $SystemErrors
        }
    }

    # =====================================================================
    # 3. WRITE LOGS (Best Effort)
    # =====================================================================
    $logFolder = $Context.Settings.SaveLogFiles.Where.Folder

    if (-not $logFolder) {
        $fallBackLogFolder = Join-Path $env:TEMP 'PermissionMatrixLogs'

        if (-not (Test-Path -LiteralPath $fallBackLogFolder -PathType Container)) {
            $null = New-Item -ItemType Directory -Path $fallBackLogFolder -Force -ErrorAction Stop
        }

        $logFolder = $fallBackLogFolder
    }

    if ($logFolder) {
        try {
            if ($Context.FoundMatrices) {
                foreach ($matrix in $Context.Matrices) {
                    $null = Write-MatrixTroubleshootingLogHC -Matrix $matrix -Html $htmlTemplates
                }
            }
            # Only attempt to write system errors to disk if there is a log folder defined
            if ($SystemErrors.Value.Count -gt 0) {
                # We will hold the generated attachments path here to pass to MailKit
                $sysErrAttachments = @() 
                Write-SystemErrorLogHC -SystemErrors $SystemErrors.Value -LogFolder $logFolder -MailParams ([ref]@{Attachments = $sysErrAttachments }) -ScriptStartTime $Context.StartTime
            }
        }
        catch {
            Add-ErrorHC -Type 'Warning' -Name 'Logging' -Message "Failed to write local logs to '$logFolder': $_" -Category 'Logging' -SystemErrors $SystemErrors
        }
    }

    # =====================================================================
    # 4. SEND EMAIL (Best Effort)
    # =====================================================================
    if ($Context.Settings.SendMail) {
        try {
            $mailParams = Build-MailParametersHC `
                -Settings $Context.Settings `
                -Html $fullHtmlBody `
                -ExportedFiles $Context.ExportedFiles `
                -Counter $Context.Counter `
                -SystemErrors $SystemErrors.Value `
                -MatrixCount $Context.Matrices.Count `
                -MailToDefaultsFile $Context.Defaults.MailTo `
                -LogFolder $logFolder `
                -ScriptStartTime $Context.StartTime
            
            # Re-attach the JSON error log if it was successfully created in Step 3
            if ($sysErrAttachments) { $mailParams.Attachments = $sysErrAttachments }

            Send-MailKitMessageHC @mailParams
            
            if ($logFolder) { $null = Save-MailBodyToLogHC -MailParams $mailParams -LogFolder $logFolder }
        }
        catch {
            Add-ErrorHC -Type 'Warning' -Name 'Email Failed' -Message "Failed to send summary email: $_" -Category 'Reporting' -SystemErrors $SystemErrors
        }
    }

    # =====================================================================
    # 5. EVENT LOG & CLEANUP (Best Effort)
    # =====================================================================
    try {
        if ($Context.Settings.SaveInEventLog.Save) {
            $eventData = [System.Collections.Generic.List[PSObject]]::new()
            Write-EventLogSafe -EventLogData $eventData -ScriptName ($Context.Settings.ScriptName ?? 'Permission Matrix') -Settings $Context.Settings -SystemErrors $SystemErrors
        }
        if ($Context.Settings.SaveLogFiles.DeleteLogsAfterDays -gt 0 -and $logFolder) {
            Cleanup-OldLogsHC -LogFolder $logFolder -RetentionDays $Context.Settings.SaveLogFiles.DeleteLogsAfterDays -SystemErrors $SystemErrors
        }
    }
    catch {
        # Final catch, we don't need to log this failure anywhere else since we're tearing down
    }
}