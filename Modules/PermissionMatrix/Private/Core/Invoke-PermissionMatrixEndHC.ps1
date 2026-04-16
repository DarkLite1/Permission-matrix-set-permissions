function Invoke-PermissionMatrixEndHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Context,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    # Tally up all errors and warnings generated across the runspaces
    $Context.Counter = Update-MatrixCounterHC `
        -Context $Context `
        -SystemErrors $SystemErrors

    $hasFatalErrors = Test-HasFatalErrorsHC $SystemErrors
    $htmlTemplates = Initialize-HtmlStructureHC
    $fullHtmlBody = ''

    # =====================================================================
    # 1. BUILD HTML BODY (Best Effort)
    # =====================================================================
    try {
        $matrixHtml = if ($Context.FoundMatrices) { 
            Build-MatrixEmailHtmlHC `
                -ImportedMatrix $Context.Matrices `
                -Html $htmlTemplates 
        }
        else { '' }
        $fullHtmlBody = Generate-MailBodyHtmlHC `
            -Settings $Context.Config.Settings `
            -ScriptStartTime $Context.StartTime `
            -Html @{ 
            Style             = $htmlTemplates.Style 
            MatrixTables      = $matrixHtml 
            ErrorWarningTable = (
                Build-ErrorWarningTableHC `
                    -CounterData $Context.Counter `
                    -SystemErrors $SystemErrors
            )
        }
    }
    catch {
        Add-ErrorHC -Type 'Warning' -Name 'HTML Generation' -Message "Failed to build HTML body: $_" -Category 'Reporting' -SystemErrors $SystemErrors
    }

    # =====================================================================
    # 2. EXPORTS & SERVICENOW (Skip if Fatal Errors)
    # =====================================================================
    if (-not $hasFatalErrors -and $Context.FoundMatrices) {
        try {
            $Context.ExportedFiles = Export-FilesHC -ImportedMatrix $Context.Matrices -ExportSettings $Context.Config.Export -HtmlOverview $fullHtmlBody -Counters $Context.Counter
            
            if ($Context.Config.Export.ServiceNowFormDataExcelFile -and $Context.Config.ServiceNow.CredentialsFilePath) {
                $snowParams = @{ CredentialsFilePath = $Context.Config.ServiceNow.CredentialsFilePath; Environment = $Context.Config.ServiceNow.Environment; TableName = $Context.Config.ServiceNow.TableName; FormDataExcelFilePath = $Context.Config.Export.ServiceNowFormDataExcelFile; ExcelFileWorksheetName = 'SnowFormData' }
                & $Context.ScriptPath.UpdateServiceNow @snowParams
            }
        }
        catch {
            Add-ErrorHC -Type 'Warning' -Name 'Exports/ServiceNow' -Message "Failed during export phase: $_" -Category 'Reporting' -SystemErrors $SystemErrors
        }
    }

    #region Create log folder
    $logFolder = $Context.Config.Settings.SaveLogFiles.Where.Folder
    $tempLogFolder = Join-Path $env:TEMP 'PermissionMatrixLogs'
    
    # Use temp folder if no log folder is specified
    if ([string]::IsNullOrWhiteSpace($logFolder)) {
        $logFolder = $tempLogFolder
    }

    try {
        if (-not (Test-Path -LiteralPath $logFolder -PathType Container)) {
            $null = New-Item -ItemType Directory -Path $logFolder -Force -ErrorAction Stop
        }
    }
    catch {
        if ($logFolder -ne $tempLogFolder) {
            Add-ErrorHC `
                -Type 'Warning' `
                -Name 'Log Folder Fallback' `
                -Message "Failed to create configured log folder '$logFolder': $_" `
                -Description "Falling back to temporary log folder '$tempLogFolder'." `
                -Category 'Logging' `
                -SystemErrors $SystemErrors

            $logFolder = $tempLogFolder

            # Try to create the temp folder
            try {
                if (-not (Test-Path -LiteralPath $logFolder -PathType Container)) {
                    $null = New-Item -ItemType Directory -Path $logFolder -Force -ErrorAction Stop
                }
            }
            catch { $logFolder = $null }
        } 
        else { $logFolder = $null }
    }
    #endregion

    # 3. Normal / Write the logs to the verified folder
    if ($logFolder) {
        try {
            if ($Context.FoundMatrices) {
                #region Create dated log folder
                $dateStr = $Context.StartTime.ToString('yyyy_MM_dd_HHmmss')

                $datedLogFolder = New-Item -ItemType Directory `
                    -Path (Join-Path `
                        -Path $logFolder `
                        -ChildPath "$dateStr - $($Context.JsonFileName)" ) `
                    -Force -ErrorAction Stop
                #endregion

                foreach ($matrix in $Context.Matrices) {
                    #region Create matrix-specific log folder
                    $matrix.LogFolder = New-Item -ItemType Directory `
                        -Path (Join-Path `
                            -Path $datedLogFolder.FullName `
                            -ChildPath $matrix.File.Item.BaseName) `
                        -Force -ErrorAction Stop
                    #endregion
                    
                    Write-MatrixTroubleshootingLogHC `
                        -Matrix $matrix `
                        -Html $htmlTemplates `
                        -LogFolderPath $matrix.LogFolder.FullName
                }
            }
            
            if ($SystemErrors.Value.Count -gt 0) {
                $sysErrAttachments = @() 
                Write-SystemErrorLogHC `
                    -SystemErrors $SystemErrors.Value `
                    -LogFolder $logFolder `
                    -MailParams ([ref]@{Attachments = $sysErrAttachments }) `
                    -ScriptStartTime $Context.StartTime
            }
        }
        catch {
            Add-ErrorHC `
                -Type 'Warning' `
                -Name 'Logging' `
                -Message "Failed to write local logs to '$logFolder': $_" `
                -Category 'Logging' `
                -SystemErrors $SystemErrors
        }
    }
    else {
        Add-ErrorHC `
            -Type 'Warning' `
            -Name 'Log Folder Unavailable' `
            -Message "No valid log folder available. Logs will not be saved to disk: $_" `
            -Category 'Logging' `
            -SystemErrors $SystemErrors
    }

    # =====================================================================
    # 4. SEND EMAIL (Best Effort)
    # =====================================================================
    if ($Context.Config.Settings.SendMail) {
        try {
            $mailParams = Build-MailParametersHC `
                -Settings $Context.Config.Settings `
                -Html $fullHtmlBody `
                -ExportedFiles $Context.ExportedFiles `
                -Counter $Context.Counter `
                -SystemErrors $SystemErrors.Value `
                -MatrixCount $Context.Matrices.Count `
                -MailToDefaultsFile $Context.Defaults.MailTo `
                -LogFolder $logFolder `
                -ScriptStartTime $Context.StartTime
            
            # Re-attach the JSON error log if it was successfully created in Step 3
            if ($sysErrAttachments) { 
                $mailParams.Attachments = $sysErrAttachments 
            }

            Send-MailKitMessageHC @mailParams
            
            if ($logFolder) { 
                $null = Save-MailBodyToLogHC `
                    -MailParams $mailParams `
                    -LogFolder $logFolder 
            }
        }
        catch {
            Add-ErrorHC `
                -Type 'Warning' `
                -Name 'Email Failed' `
                -Message "Failed to send summary email: $_" `
                -Category 'Reporting' `
                -SystemErrors $SystemErrors
        }
    }

    # =====================================================================
    # 5. EVENT LOG & CLEANUP (Best Effort)
    # =====================================================================
    try {
        if ($Context.Config.Settings.SaveInEventLog.Save) {
            $eventData = [System.Collections.Generic.List[PSObject]]::new()

            Write-EventLogSafe `
                -EventLogData $eventData `
                -ScriptName (
                $Context.Config.Settings.ScriptName ?? 'Permission Matrix') `
                -Settings $Context.Config.Settings `
                -SystemErrors $SystemErrors
        }
        if ($Context.Config.Settings.SaveLogFiles.DeleteLogsAfterDays -gt 0 -and $logFolder) {
            Cleanup-OldLogsHC `
                -LogFolder $logFolder `
                -RetentionDays $Context.Config.Settings.SaveLogFiles.DeleteLogsAfterDays `
                -SystemErrors $SystemErrors
        }
    }
    catch {
        # Final catch, we don't need to log this failure anywhere else since we're tearing down
    }
}