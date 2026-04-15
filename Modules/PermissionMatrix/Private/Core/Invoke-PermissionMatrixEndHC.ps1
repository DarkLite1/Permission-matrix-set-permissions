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
            -Settings $Context.Settings `
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
    $tempLogFolder = Join-Path $env:TEMP 'PermissionMatrixLogs'
    
    # 1. When settings has no log folder, we use the temp folder
    if ([string]::IsNullOrWhiteSpace($logFolder)) {
        $logFolder = $tempLogFolder
    }

    # Attempt to create/validate the chosen log folder
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

    # Update the Context so the Email and Cleanup logic know the correct, working path
    if ($Context.Settings.SaveLogFiles.Where) {
        $Context.Settings.SaveLogFiles.Where.Folder = $logFolder
    }

    # 3. Normal / Write the logs to the verified folder
    if ($logFolder) {
        try {
            if ($Context.FoundMatrices) {
                $dateStr = $Context.StartTime.ToString('yyyy_MM_dd_HHmmss')

                foreach ($matrix in $Context.Matrices) {
                    #region Create matrix-specific log folder
                    $matrixBaseName = [System.IO.Path]::GetFileNameWithoutExtension($matrix.File.Name)

                    $specificFolder = Join-Path `
                        -Path $logFolder `
                        -ChildPath "$dateStr - $matrixBaseName"
                    
                    if (-not (Test-Path -LiteralPath $specificFolder -PathType Container)) {
                        $null = New-Item -ItemType Directory -Path $specificFolder -Force -ErrorAction Stop
                    }
                    #endregion
                    
                    $matrix.File.LogFolder = $specificFolder

                    $null = Write-MatrixTroubleshootingLogHC `
                        -Matrix $matrix `
                        -Html $htmlTemplates 
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
        if ($Context.Settings.SaveInEventLog.Save) {
            $eventData = [System.Collections.Generic.List[PSObject]]::new()

            Write-EventLogSafe `
                -EventLogData $eventData `
                -ScriptName (
                $Context.Settings.ScriptName ?? 'Permission Matrix') `
                -Settings $Context.Settings `
                -SystemErrors $SystemErrors
        }
        if ($Context.Settings.SaveLogFiles.DeleteLogsAfterDays -gt 0 -and $logFolder) {
            Cleanup-OldLogsHC `
                -LogFolder $logFolder `
                -RetentionDays $Context.Settings.SaveLogFiles.DeleteLogsAfterDays `
                -SystemErrors $SystemErrors
        }
    }
    catch {
        # Final catch, we don't need to log this failure anywhere else since we're tearing down
    }
}