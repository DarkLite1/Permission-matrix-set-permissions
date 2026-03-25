function Generate-MailRecipientList {
    param(
        [object]$Recipients,
        [array]$Defaults = @()
    )

    #
    # Normalize a single input value into an array of strings
    #
    function Normalize-MailEntry {
        param([object]$Value)

        if (-not $Value) { return @() }
        if ($Value -is [string]) { return , $Value.Trim() }
        if ($Value -is [array]) { return $Value | ForEach-Object { "$_".Trim() } }

        # Unsupported type → ignore silently (same as original loose behavior)
        return @()
    }

    #
    # Normalize primary list + defaults
    #
    $combined =
    (Normalize-MailEntry -Value $Recipients) +
    (Normalize-MailEntry -Value $Defaults)

    #
    # Remove empty values, trim, unique, sorted
    #
    return $combined |
    ForEach-Object { "$_".Trim() } |
    Where-Object { $_ } |
    Sort-Object -Unique
}
function Generate-MailSubject {
    param(
        [int]$MatrixCount,
        [object]$SystemErrors,
        [object]$Counter,
        [string]$CustomSubject
    )

    #
    # Helper: pluralize a word if needed
    #
        

    #
    # Optional custom suffix
    #
    $suffix = if ($CustomSubject) { ", $CustomSubject" } else { '' }

    $subject = $null

    #
    # 1. If system errors exist → priority subject line
    #
    if ($SystemErrors.Count -gt 0) {

        $sysWord = Plural -Count $SystemErrors.Count -Word 'System Error'
        $matWord = Plural -Count $MatrixCount -Word 'matrix file'

        if ($MatrixCount -gt 0) {
            $subject = "$MatrixCount $matWord, $($SystemErrors.Count) $sysWord$suffix"
        }
        else {
            # No matrix, only system failures
            $subject = "$sysWord`: $($SystemErrors.Count) critical failure$(if ($SystemErrors.Count -ne 1) {'s'})$suffix"
        }
    }

    #
    # 2. Otherwise → Matrix subject with error and warning summary
    #
    else {
        $matWord = Plural -Count $MatrixCount -Word 'matrix file'

        $errPart = if ($Counter.TotalErrors) {
            ", $($Counter.TotalErrors) error$(if($Counter.TotalErrors-ne 1){'s'})" 
        }
        else { '' }

        $warnPart = if ($Counter.TotalWarnings) {
            ", $($Counter.TotalWarnings) warning$(if($Counter.TotalWarnings-ne 1){'s'})" 
        }
        else { '' }

        $subject = "$MatrixCount $matWord$errPart$warnPart$suffix"
    }

    #
    # 3. Sanitize for use as filenames (original behavior preserved)
    #
    return [string]::Join(
        '_',
        $subject.Split([System.IO.Path]::GetInvalidFileNameChars())
    )
}
function Build-MailParameters {
    param(
        [Parameter(Mandatory)][object]$Settings,
        [Parameter(Mandatory)][object]$Html,
        [Parameter()][object]$ExportedFiles,
        [Parameter()][object]$Counter,
        [Parameter()][object]$SystemErrors,
        [Parameter()][int]$MatrixCount,
        [Parameter(Mandatory)][hashtable]$ExistingMailParams,
        [Parameter()][array]$MailToDefaultsFile,
        [Parameter()][string]$LogFolder,
        [Parameter()][datetime]$ScriptStartTime
    )

    #
    # 1. Prepare base hashtable
    #
    $mail = $ExistingMailParams
    $sendMail = $Settings.SendMail
    $smtp = $sendMail.Smtp

    #
    # 2. Recipients
    #
    $mail.To = Generate-MailRecipientList `
        -Recipients $sendMail.To `
        -Defaults $MailToDefaultsFile

    if ($sendMail.Bcc) {
        $mail.Bcc = Generate-MailRecipientList -Recipients $sendMail.Bcc
    }

    #
    # 3. Basic metadata
    #
    $mail.From = Get-StringValueHC $sendMail.From
    $mail.FromDisplayName = Get-StringValueHC $sendMail.FromDisplayName
    $mail.SmtpServerName = Get-StringValueHC $smtp.ServerName
    $mail.SmtpPort = Get-StringValueHC $smtp.Port
    $mail.SmtpConnectionType = Get-StringValueHC $smtp.ConnectionType

    $mail.MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
    $mail.MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit

    #
    # 4. Credential (optional)
    #
    if ($smtp.UserName -and $smtp.Password) {
        $sec = ConvertTo-SecureString `
            -String (Get-StringValueHC $smtp.Password) `
            -AsPlainText -Force

        $mail.Credential = New-Object `
            System.Management.Automation.PSCredential `
        (Get-StringValueHC $smtp.UserName), $sec
    }

    #
    # 5. Subject line
    #
    $mail.Subject = Generate-MailSubject `
        -MatrixCount $MatrixCount `
        -SystemErrors $SystemErrors `
        -Counter $Counter `
        -CustomSubject $sendMail.Subject

    #
    # 6. Mail priority
    #
    if (
        $SystemErrors.Count -gt 0 -or
        $Counter.TotalErrors -gt 0 -or
        $Counter.TotalWarnings -gt 0
    ) {
        $mail.Priority = 'High'
    }

    #
    # 7. Build the mail body
    #
    $attachmentNote = if ($mail.Attachments) {
        '<p><i>* Check the attachment(s) for details</i></p>'
    }

    $durationString = $null
    if ($ScriptStartTime) {
        $ts = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
        $durationString = '{0:00}:{1:00}:{2:00}' -f $ts.Hours, $ts.Minutes, $ts.Seconds
    }

    $mail.Body = Generate-MailBodyHtml `
        -Settings $Settings `
        -Html $Html `
        -ExportedFiles $ExportedFiles `
        -AttNote $attachmentNote `
        -DurStr $durationString `
        -ScriptStartTime $ScriptStartTime `
        -LogFolder $LogFolder

    return $mail
}
function Send-MailSafe {
    param(
        [Parameter(Mandatory)][hashtable]$MailParams,
        [Parameter(Mandatory)][ref]$SystemErrors
    )
    try { Send-MailKitMessageHC @MailParams }
    catch {
        $SystemErrors.Value.Add(
            [PSCustomObject]@{
                DateTime = Get-Date 
                Message  = "Failed to send mail: $_" 
            }
        ) 
    }
}
function Save-MailBodyToLog {
    param(
        [Parameter(Mandatory)][hashtable]$MailParams,
        [Parameter(Mandatory)][string]$LogFolder,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    try {
        # No subject → no log file
        if (-not $MailParams.Subject) {
            return
        }

        # Ensure log folder exists
        if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
            return
        }

        # Build final file path
        $fileName = "Mail - $($MailParams.Subject).html"
        $fullPath = Join-Path (Get-DatedLogFolderPathHC) $fileName

        # Save HTML
        $MailParams.Body |
        Out-File -LiteralPath $fullPath -Encoding UTF8 -Force
    }
    catch {
        $SystemErrors.Value.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Failed to save mail HTML: $_"
            }
        )
    }
}