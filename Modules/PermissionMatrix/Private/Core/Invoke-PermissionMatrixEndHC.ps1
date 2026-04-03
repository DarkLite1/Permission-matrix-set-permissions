function Invoke-PermissionMatrixEndHC {
    <#
    .SYNOPSIS
        END stage for the Permission Matrix pipeline.
    .DESCRIPTION
        1. Sequential: Generates Excel and HTML exports.
        2. Sequential: Executes ServiceNow update if configured.
        3. Sequential: Writes per-matrix troubleshooting logs and sends the summary email.
        4. Sequential: Cleans up old logs and writes to the Windows Event Log.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Context,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    try {
        $hasFatalErrors = Test-HasFatalErrorsHC $SystemErrors

        # =====================================================================
        # 1. EXPORTS & SERVICENOW (Skip if Fatal Errors exist)
        # =====================================================================
        if (-not $hasFatalErrors -and $Context.FoundMatrices) {
            
            # Generate HTML Overview for Export and Email
            $htmlTemplates = Initialize-HtmlStructureHC
            $matrixHtml = Build-MatrixEmailHtmlHC -ImportedMatrix $Context.Matrices -Html $htmlTemplates
            $fullHtmlBody = Generate-MailBodyHtmlHC `
                -Settings $Context.Settings `
                -Html @{ Style = $htmlTemplates.Style; ErrorWarningTable = (Build-ErrorWarningTableHC -CounterData $Context.Counter -SystemErrors $SystemErrors); MatrixTables = $matrixHtml } `
                -ScriptStartTime $Context.StartTime

            # Execute File Exports
            $Context.ExportedFiles = Export-FilesHC `
                -ImportedMatrix $Context.Matrices `
                -ExportSettings $Context.Export `
                -HtmlOverview $fullHtmlBody `
                -Counters $Context.Counter

            # Update ServiceNow 
            if ($Context.Export.ServiceNowFormDataExcelFile -and $Context.ServiceNow.CredentialsFilePath) {
                try {
                    $snowParams = @{
                        CredentialsFilePath    = $Context.ServiceNow.CredentialsFilePath
                        Environment            = $Context.ServiceNow.Environment
                        TableName              = $Context.ServiceNow.TableName
                        FormDataExcelFilePath  = $Context.Export.ServiceNowFormDataExcelFile
                        ExcelFileWorksheetName = 'SnowFormData'
                    }
                    & $Context.ScriptPath.UpdateServiceNow @snowParams
                }
                catch {
                    Add-ErrorHC -Type 'Warning' -Name 'ServiceNow Update Failed' -Message "Failed executing ServiceNow script: $_" -Category 'ServiceNow' -SystemErrors $SystemErrors
                }
            }
        }

        # =====================================================================
        # 2. LOGGING & EMAIL
        # =====================================================================
        if ($Context.FoundMatrices -or $SystemErrors.Value.Count -gt 0) {
            
            # Write Individual Matrix Troubleshooting Logs
            if ($Context.FoundMatrices) {
                foreach ($matrix in $Context.Matrices) {
                    $null = Write-MatrixTroubleshootingLogHC -Matrix $matrix -Html (Initialize-HtmlStructureHC)
                }
            }

            # Send Summary Email
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
                        -LogFolder $Context.Settings.SaveLogFiles.Where.Folder `
                        -ScriptStartTime $Context.StartTime

                    Write-SystemErrorLogHC -SystemErrors $SystemErrors.Value -LogFolder $Context.Settings.SaveLogFiles.Where.Folder -MailParams ([ref]$mailParams) -ScriptStartTime $Context.StartTime
                    
                    Send-MailKitMessageHC @mailParams
                    $null = Save-MailBodyToLogHC -MailParams $mailParams -LogFolder $Context.Settings.SaveLogFiles.Where.Folder
                }
                catch {
                    Add-ErrorHC -Type 'Warning' -Name 'Email Failed' -Message "Failed to send summary email: $_" -Category 'Reporting' -SystemErrors $SystemErrors
                }
            }
        }

        # =====================================================================
        # 3. SYSTEM TEARDOWN (Cleanup & Event Log)
        # =====================================================================
        
        # Log Retention Cleanup 
        if ($Context.Settings.SaveLogFiles.DeleteLogsAfterDays -gt 0) {
            Cleanup-OldLogsHC `
                -LogFolder $Context.Settings.SaveLogFiles.Where.Folder `
                -RetentionDays $Context.Settings.SaveLogFiles.DeleteLogsAfterDays `
                -SystemErrors $SystemErrors
        }

        # Write to Windows Event Log 
        if ($Context.Settings.SaveInEventLog.Save) {
            $eventData = [System.Collections.Generic.List[PSObject]]::new()
            Write-EventLogSafe `
                -EventLogData $eventData `
                -ScriptName $Context.Settings.ScriptName `
                -Settings $Context.Settings `
                -SystemErrors $SystemErrors
        }

    }
    catch {
        Add-ErrorHC -Type 'FatalError' -Category 'Runtime' -Name 'END stage failure' -Message "Unhandled exception occurred: $_" -SystemErrors $SystemErrors
    }
}