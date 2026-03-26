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
function Send-MailKitMessageHC {
    <#
            .SYNOPSIS
                Send an email using MailKit and MimeKit assemblies.

            .DESCRIPTION
                This function sends an email using the MailKit and MimeKit
                assemblies. It requires the assemblies to be installed before
                calling the function:

                $params = @{
                    Source           = 'https://www.nuget.org/api/v2'
                    SkipDependencies = $true
                    Scope            = 'AllUsers'
                }
                Install-Package @params -Name 'MailKit'
                Install-Package @params -Name 'MimeKit'

            .PARAMETER MailKitAssemblyPath
                The path to the MailKit assembly.

            .PARAMETER MimeKitAssemblyPath
                The path to the MimeKit assembly.

            .PARAMETER SmtpServerName
                The name of the SMTP server.

            .PARAMETER SmtpPort
                The port of the SMTP server.

            .PARAMETER SmtpConnectionType
                The connection type for the SMTP server.

                Valid values are:
                - 'None'
                - 'Auto'
                - 'SslOnConnect'
                - 'StartTlsWhenAvailable'
                - 'StartTls'

            .PARAMETER Credential
                The credential object containing the username and password.

            .PARAMETER From
                The sender's email address.

            .PARAMETER FromDisplayName
            The display name to show for the sender.

            Email clients may display this differently. It is most likely
            to be shown if the sender's email address is not recognized
                (e.g., not in the address book).

            .PARAMETER To
                The recipient's email address.

            .PARAMETER Body
            The body of the email, HTML is supported.

            .PARAMETER Subject
            The subject of the email.

            .PARAMETER Attachments
            An array of file paths to attach to the email.

            .PARAMETER Priority
            The email priority.

            Valid values are:
            - 'Low'
            - 'Normal'
            - 'High'

            .EXAMPLE
            # Send an email with StartTls and credential

            $SmtpUserName = 'smtpUser'
            $SmtpPassword = 'smtpPassword'

            $securePassword = ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($SmtpUserName, $securePassword)

            $params = @{
                SmtpServerName = 'SMT_SERVER@example.com'
                SmtpPort = 587
                SmtpConnectionType = 'StartTls'
                Credential = $credential
                from = 'm@example.com'
                To = '007@example.com'
                Body = '<p>Mission details in attachment</p>'
                Subject = 'For your eyes only'
                Priority = 'High'
                Attachments = @('c:\Mission.ppt', 'c:\ID.pdf')
                MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
            }

            Send-MailKitMessageHC @params

            .EXAMPLE
            # Send an email without authentication

            $params = @{
                SmtpServerName      = 'SMT_SERVER@example.com'
                SmtpPort            = 25
                From                = 'hacker@example.com'
                FromDisplayName     = 'White hat hacker'
                Bcc                 = @('james@example.com', 'mike@example.com')
                Body                = '<h1>You have been hacked</h1>'
                Subject             = 'Oops'
                MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
            }

            Send-MailKitMessageHC @params
            #>

    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [string]$MailKitAssemblyPath,
        [parameter(Mandatory)]
        [string]$MimeKitAssemblyPath,
        [parameter(Mandatory)]
        [string]$SmtpServerName,
        [parameter(Mandatory)]
        [ValidateSet(25, 465, 587, 2525)]
        [int]$SmtpPort,
        [parameter(Mandatory)]
        [string]$Body,
        [parameter(Mandatory)]
        [string]$Subject,
        [parameter(Mandatory)]
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
        [string]$From,
        [string]$FromDisplayName,
        [string[]]$To,
        [string[]]$Bcc,
        [int]$MaxAttachmentSize = 20MB,
        [ValidateSet(
            'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
        )]
        [string]$SmtpConnectionType = 'None',
        [ValidateSet('Normal', 'Low', 'High')]
        [string]$Priority = 'Normal',
        [string[]]$Attachments,
        [PSCredential]$Credential
    )

    begin {
        function Test-IsAssemblyLoaded {
            param (
                [String]$Name
            )
            foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
                if ($assembly.FullName -like "$Name, Version=*") {
                    return $true
                }
            }
            return $false
        }

        function Add-Attachments {
            param (
                [string[]]$Attachments,
                [MimeKit.Multipart]$BodyMultiPart
            )

            $attachmentList = New-Object System.Collections.ArrayList($null)
            $totalSizeAttachments = 0

            foreach (
                $attachmentPath in
                $Attachments | Sort-Object -Unique
            ) {
                try {
                    #region Test if file exists
                    try {
                        $attachmentItem = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop

                        if ($attachmentItem.PSIsContainer) {
                            Write-Warning "Attachment '$attachmentPath' is a folder, not a file"
                            continue
                        }
                    }
                    catch {
                        Write-Warning "Attachment '$attachmentPath' not found"
                        continue
                    }
                    #endregion

                    $totalSizeAttachments += $attachmentItem.Length
                    $null = $attachmentList.Add($attachmentItem)

                    #region Check size of attachments
                    if ($totalSizeAttachments -ge $MaxAttachmentSize) {
                        $M = 'The maximum allowed attachment size of {0} MB has been exceeded ({1} MB). No attachments were added to the email. Check the log folder for details.' -f
                        ([math]::Round(($MaxAttachmentSize / 1MB))),
                        ([math]::Round(($totalSizeAttachments / 1MB), 2))

                        Write-Warning $M

                        return [PSCustomObject]@{
                            AttachmentLimitExceededMessage = $M
                        }
                    }
                }
                catch {
                    Write-Warning "Failed to add attachment '$attachmentPath': $_"
                }
            }
            #endregion

            foreach (
                $attachmentItem in
                $attachmentList
            ) {
                try {
                    Write-Verbose "Add mail attachment '$($attachmentItem.Name)'"

                    $attachment = New-Object MimeKit.MimePart

                    #region Create a MemoryStream to hold the file content
                    $memoryStream = New-Object System.IO.MemoryStream

                    try {
                        $fileStream = [System.IO.File]::OpenRead($attachmentItem.FullName)
                        $fileStream.CopyTo($memoryStream)
                    }
                    finally {
                        if ($fileStream) {
                            $fileStream.Dispose()
                        }
                    }

                    $memoryStream.Position = 0
                    #endregion

                    $attachment.Content = New-Object MimeKit.MimeContent($memoryStream)

                    $attachment.ContentDisposition = New-Object MimeKit.ContentDisposition

                    $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64

                    $attachment.FileName = $attachmentItem.Name

                    $bodyMultiPart.Add($attachment)
                }
                catch {
                    Write-Warning "Failed to add attachment '$attachmentItem': $_"
                }
            }
        }

        try {
            #region Test To or Bcc required
            if (-not ($To -or $Bcc)) {
                throw "Either 'To' to 'Bcc' is required for sending emails"
            }
            #endregion

            #region Test To
            foreach ($email in $To) {
                if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                    throw "To email address '$email' not valid."
                }
            }
            #endregion

            #region Test Bcc
            foreach ($email in $Bcc) {
                if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                    throw "Bcc email address '$email' not valid."
                }
            }
            #endregion

            #region Load MimeKit assembly
            if (-not(Test-IsAssemblyLoaded -Name 'MimeKit')) {
                try {
                    Write-Verbose "Load MimeKit assembly '$MimeKitAssemblyPath'"
                    Add-Type -Path $MimeKitAssemblyPath
                }
                catch {
                    throw "Failed to load MimeKit assembly '$MimeKitAssemblyPath': $_"
                }
            }
            #endregion

            #region Load MailKit assembly
            if (-not(Test-IsAssemblyLoaded -Name 'MailKit')) {
                try {
                    Write-Verbose "Load MailKit assembly '$MailKitAssemblyPath'"
                    Add-Type -Path $MailKitAssemblyPath
                }
                catch {
                    throw "Failed to load MailKit assembly '$MailKitAssemblyPath': $_"
                }
            }
            #endregion
        }
        catch {
            throw "Failed to send email to '$To': $_"
        }
    }

    process {
        try {
            $message = New-Object -TypeName 'MimeKit.MimeMessage'

            #region Create body with attachments
            $bodyPart = New-Object MimeKit.TextPart('html')
            $bodyPart.Text = $Body

            $bodyMultiPart = New-Object MimeKit.Multipart('mixed')
            $bodyMultiPart.Add($bodyPart)

            if ($Attachments) {
                $params = @{
                    Attachments   = $Attachments
                    BodyMultiPart = $bodyMultiPart
                }
                $addAttachments = Add-Attachments @params

                if ($addAttachments.AttachmentLimitExceededMessage) {
                    $bodyPart.Text += '<p><i>{0}</i></p>' -f
                    $addAttachments.AttachmentLimitExceededMessage
                }
            }

            $message.Body = $bodyMultiPart
            #endregion

            $fromAddress = New-Object MimeKit.MailboxAddress(
                $FromDisplayName, $From
            )
            $message.From.Add($fromAddress)

            foreach ($email in $To) {
                $message.To.Add($email)
            }

            foreach ($email in $Bcc) {
                $message.Bcc.Add($email)
            }

            $message.Subject = $Subject

            #region Set priority
            switch ($Priority) {
                'Low' {
                    $message.Headers.Add('X-Priority', '5 (Lowest)')
                    break
                }
                'Normal' {
                    $message.Headers.Add('X-Priority', '3 (Normal)')
                    break
                }
                'High' {
                    $message.Headers.Add('X-Priority', '1 (Highest)')
                    break
                }
                default {
                    throw "Priority type '$_' not supported"
                }
            }
            #endregion

            $smtp = New-Object -TypeName 'MailKit.Net.Smtp.SmtpClient'

            try {
                $smtp.Connect(
                    $SmtpServerName, $SmtpPort,
                    [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType
                )
            }
            catch {
                throw "Failed to connect to SMTP server '$SmtpServerName' on port '$SmtpPort' with connection type '$SmtpConnectionType': $_"
            }

            if ($Credential) {
                try {
                    $smtp.Authenticate(
                        $Credential.UserName,
                        $Credential.GetNetworkCredential().Password
                    )
                }
                catch {
                    throw "Failed to authenticate with user name '$($Credential.UserName)' to SMTP server '$SmtpServerName': $_"
                }
            }

            Write-Verbose "Send mail to '$To' with subject '$Subject'"

            $null = $smtp.Send($message)
        }
        catch {
            throw "Failed to send email to '$To': $_"
        }
        finally {
            if ($smtp) {
                $smtp.Disconnect($true)
                $smtp.Dispose()
            }
            if ($message) {
                $message.Dispose()
            }
        }
    }
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