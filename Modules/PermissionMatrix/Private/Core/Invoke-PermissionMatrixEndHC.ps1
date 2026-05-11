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

    $hasFatalErrors = Test-ItemHasFatalErrorHC -CheckList $SystemErrors.Value
    $htmlTemplates = Initialize-HtmlStructureHC
    $fullHtmlBody = ''
    $sysErrAttachments = @()

    # =====================================================================
    # 1. BUILD HTML BODY (Best Effort)
    # =====================================================================
    try {
        $matrixHtml = if (
            $Context.FileResults -and $Context.FileResults.Count -gt 0
        ) { 
            Build-MatrixEmailHtmlHC `
                -FileResults $Context.FileResults `
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
    if (-not $hasFatalErrors -and $Context.AllMatrices) {
        try {
            $Context.ExportedFiles = Export-FilesHC `
                -ImportedMatrix $Context.AllMatrices `
                -ExportSettings $Context.Config.Export `
                -HtmlOverview $fullHtmlBody `
                -Counters $Context.Counter
            
            if (
                $Context.Config.Export.ServiceNowFormDataExcelFile -and $Context.Config.ServiceNow.CredentialsFilePath
            ) {
                $snowParams = @{ 
                    CredentialsFilePath    = $Context.Config.ServiceNow.CredentialsFilePath
                    Environment            = $Context.Config.ServiceNow.Environment 
                    TableName              = $Context.Config.ServiceNow.TableName 
                    FormDataExcelFilePath  = $Context.Config.Export.ServiceNowFormDataExcelFile 
                    ExcelFileWorksheetName = 'SnowFormData' 
                }
                & $Context.ScriptPath.UpdateServiceNow @snowParams
            }
        }
        catch {
            Add-ErrorHC `
                -Type 'Warning' `
                -Name 'Exports/ServiceNow' `
                -Message "Failed during export phase: $_" `
                -Category 'Reporting' `
                -SystemErrors $SystemErrors
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

    #region Create log files
    if ($logFolder) {
        try {
            if ($Context.FoundMatrices) {
                #region Create dated log folder
                $dateStr = $Context.StartTime.ToString('yyyy_MM_dd_HHmmss')

                $datedLogFolder = New-Item -ItemType Directory `
                    -Path (Join-Path `
                        -Path $logFolder `
                        -ChildPath "$dateStr ($($Context.JsonFileName))" ) `
                    -Force -ErrorAction Stop
                #endregion

                foreach ($fileResult in $Context.FileResults) {
                    $baseName = $fileResult.Item.BaseName
                    
                    $fileLogFolder = New-Item -ItemType Directory `
                        -Path (Join-Path $datedLogFolder.FullName $baseName) `
                        -Force -ErrorAction Stop

                    $fileResult.LogFolder = $fileLogFolder.FullName
                    $fileResult.ReportFilePath = Join-Path `
                        -Path $fileLogFolder.FullName `
                        -ChildPath $fileResult.ReportFileName

                    #region Create JSON files for file-level checks
                    $checkIndex = 0

                    foreach ($fc in $fileResult.Check) {
                        $checkIndex++
                        $checkFileName = "File - Detail $checkIndex.json"
                        
                        $fc | Add-Member -NotePropertyMembers @{
                            JsonFileName = $checkFileName
                            JsonFilePath = Join-Path -Path $fileLogFolder.FullName -ChildPath $checkFileName  
                        } -Force
                        
                        if ($fc.Value) {
                            try {
                                $fc | Select-Object -ExcludeProperty JsonFilePath, JsonFileName | 
                                ConvertTo-Json -Depth 10 | 
                                Out-File -FilePath $fc.JsonFilePath -Encoding UTF8 -Force
                            }
                            catch {
                                $fc.Description += "[Detailed JSON log failed to generate: $($_)]"
                                $fc.JsonFileName = $null
                                $fc.JsonFilePath = $null   
                            }
                        }
                        else {
                            $fc.JsonFileName = $null
                            $fc.JsonFilePath = $null
                        }
                    }
                    #endregion

                    #region Create JSON file for matrix-level checks
                    $checkIndex = 0

                    foreach ($m in $fileResult.Matrices) {
                        foreach ($c in $m.Check) {
                            $checkIndex++
                            $checkFileName = "ID $($m.ID) - Detail $checkIndex.json"
                            
                            $c | Add-Member -NotePropertyMembers @{
                                JsonFileName = $checkFileName
                                JsonFilePath = Join-Path -Path $fileLogFolder.FullName -ChildPath $checkFileName  
                            } -Force

                            if ($c.Value) {
                                try {
                                    $c | Select-Object -ExcludeProperty JsonFilePath, JsonFileName | 
                                    ConvertTo-Json -Depth 10 | 
                                    Out-File -FilePath $c.JsonFilePath -Encoding UTF8 -Force
                                }
                                catch {
                                    $c.Description += "[Detailed JSON log failed to generate: $($_)]"
                                    $c.JsonFileName = $null
                                    $c.JsonFilePath = $null
                                }
                            } 
                            else {
                                $c.JsonFileName = $null
                                $c.JsonFilePath = $null
                            }
                        }
                    }
                    #endregion

                    <# 
                    start (ls $context.Config.Settings.SaveLogFiles.Where.Folder -Recurse -file | select -First 1).FullName

                    (ls $context.Config.Settings.SaveLogFiles.Where.Folder -Recurse -file).FullName | ForEach-Object {start $_}
                    #>

                    Write-MatrixExecutionReportHC `
                        -FileResult $fileResult `
                        -Html $htmlTemplates `
                        -LogFolder $fileLogFolder.FullName
                }
            }
            
            if ($SystemErrors.Value.Count -gt 0) {
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
    #endregion

    #region Send Summary Email
    if ($Context.Config.Settings.SendMail) {
        try {
            $sendMail = $Context.Config.Settings.SendMail

            $recipients = Generate-MailRecipientListHC `
                -SendMailSettings $sendMail `
                -MailToDefaultsFile $Context.Defaults.MailTo

            $subject = Generate-MailSubjectHC `
                -SystemErrors $SystemErrors.Value `
                -Counter $Context.Counter `
                -MatrixCount $Context.AllMatrices.Count `
                -CustomSubject $sendMail.Subject

            $priority = if (
                $SystemErrors.Value.Count -gt 0 -or
                $Context.Counter.TotalErrors -gt 0 -or
                $Context.Counter.TotalWarnings -gt 0
            ) { 'High' } else { 'Normal' }

            if ([string]::IsNullOrEmpty($fullHtmlBody)) {
                $fullHtmlBody = '<html><body>Email body unavailable due to upstream error.</body></html>'
            }

            $mailParams = @{
                To                  = $recipients
                From                = Get-StringValueHC $sendMail.From
                FromDisplayName     = Get-StringValueHC $sendMail.FromDisplayName
                SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
                SmtpPort            = Get-StringValueHC $sendMail.Smtp.Port
                SmtpConnectionType  = Get-StringValueHC $sendMail.Smtp.ConnectionType
                MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
                MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
                Subject             = $subject
                Body                = $fullHtmlBody
                Priority            = $priority
                Attachments         = $sysErrAttachments
            }

            if ($sendMail.Bcc) {
                $mailParams.Bcc = $sendMail.Bcc
            }

            # SMTP credential (only if both username and password supplied)
            $smtpUser = Get-StringValueHC $sendMail.Smtp.UserName
            $smtpPass = Get-StringValueHC $sendMail.Smtp.Password
            if ($smtpUser -and $smtpPass) {
                $secure = ConvertTo-SecureString -String $smtpPass -AsPlainText -Force
                $mailParams.Credential = New-Object System.Management.Automation.PSCredential($smtpUser, $secure)
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
    #endregion

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