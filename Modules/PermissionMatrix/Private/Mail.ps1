function Generate-MailRecipientListHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $SendMailSettings,

        $MailToDefaultsFile
    )

    $list = @()
    if ($SendMailSettings.To) { $list += $SendMailSettings.To }
    if ($MailToDefaultsFile) { $list += $MailToDefaultsFile }

    # First Where-Object drops $null / empty entries before .Trim() is called,
    # so a null in the array can no longer throw.
    return (
        $list |
        Where-Object { $_ } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ }
    ) | Sort-Object -Unique
}

function Generate-MailSubjectHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $SystemErrors,

        [Parameter(Mandatory)]
        $Counter,

        [Parameter(Mandatory)]
        $MatrixCount,

        [string]$CustomSubject
    )

    $matrixPlural = if ($MatrixCount -ne 1) { 's' } else { '' }
    $cSuffix = if ($CustomSubject) { ", $CustomSubject" } else { '' }

    # If system errors exist
    if ($SystemErrors.Count -gt 0) {
        $sysPlural = if ($SystemErrors.Count -ne 1) { 's' } else { '' }
        return "$MatrixCount matrix file$matrixPlural, $($SystemErrors.Count) System Error$sysPlural$cSuffix"
    }

    # No system errors: embed matrix counts + warnings/errors
    $err = $Counter.TotalErrors
    $warn = $Counter.TotalWarnings

    $errPart = if ($err -gt 0) { ", $err error$(if ($err -ne 1) {'s'})" } else { '' }
    $warnPart = if ($warn -gt 0) { ", $warn warning$(if ($warn -ne 1) {'s'})" } else { '' }

    return "$MatrixCount matrix file$matrixPlural$errPart$warnPart$cSuffix"
}

function Send-MailKitMessageHC {
    <#
        Sends email using MailKit.
        Expected parameters:
            - MailKitAssemblyPath
            - MimeKitAssemblyPath
            - SmtpServerName
            - SmtpPort
            - SmtpConnectionType
            - Credential (optional)
            - From
            - FromDisplayName
            - To[]
            - Bcc[]
            - Body (HTML)
            - Subject
            - Attachments[]
            - Priority
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$MailKitAssemblyPath,
        [Parameter(Mandatory)][string]$MimeKitAssemblyPath,
        [Parameter(Mandatory)][string]$SmtpServerName,
        [Parameter(Mandatory)][int]$SmtpPort,
        [Parameter(Mandatory)][string]$Body,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$From,
        [string]$FromDisplayName,
        [string[]]$To,
        [string[]]$Bcc,
        [ValidateSet('High', 'Normal', 'Low')]
        [string]$Priority = 'Normal',
        [string[]]$Attachments,
        [ValidateSet('None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable')]
        [string]$SmtpConnectionType = 'None',
        [PSCredential]$Credential
    )

    # Load assemblies (MimeKit first; MailKit depends on it). Wrap Add-Type so a
    # missing, unset, or invalid DLL produces a clear, actionable error instead
    # of a cryptic assembly-load failure.
    foreach ($assembly in @(
            [PSCustomObject]@{ Name = 'MimeKit'; Path = $MimeKitAssemblyPath },
            [PSCustomObject]@{ Name = 'MailKit'; Path = $MailKitAssemblyPath }
        )) {
        if ([string]::IsNullOrWhiteSpace($assembly.Path)) {
            throw "The $($assembly.Name) assembly path is not set. Provide the full path to $($assembly.Name).dll (e.g. via the configuration file or its environment variable)."
        }

        try {
            Add-Type -Path $assembly.Path -ErrorAction Stop
        }
        catch {
            throw "Failed to load the $($assembly.Name) assembly from '$($assembly.Path)'. $($_.Exception.Message) Verify the path is correct and that the MimeKit and MailKit NuGet packages are installed."
        }
    }

    # Streams opened for attachments must stay open until after Send (MimeKit
    # reads them during Send), then be disposed in the finally block.
    $attachmentStreams = [System.Collections.Generic.List[System.IO.Stream]]::new()
    $smtp = $null

    try {
        # Create message
        $message = New-Object MimeKit.MimeMessage
        $fromAddress = New-Object MimeKit.MailboxAddress($FromDisplayName, $From)
        $message.From.Add($fromAddress)

        # InternetAddressList.Add takes an InternetAddress, not a string, so
        # parse each recipient into a MailboxAddress.
        foreach ($t in $To) {
            $message.To.Add([MimeKit.MailboxAddress]::Parse($t))
        }
        foreach ($b in $Bcc) {
            $message.Bcc.Add([MimeKit.MailboxAddress]::Parse($b))
        }

        $message.Subject = $Subject

        # Priority header
        switch ($Priority) {
            'High' { $message.Headers.Add('X-Priority', '1 (Highest)') }
            'Normal' { $message.Headers.Add('X-Priority', '3 (Normal)') }
            'Low' { $message.Headers.Add('X-Priority', '5 (Lowest)') }
        }

        # Body
        $bodyPart = New-Object MimeKit.TextPart('html')
        $bodyPart.Text = $Body
        $bodyContainer = New-Object MimeKit.Multipart 'mixed'
        $bodyContainer.Add($bodyPart)

        # Attachments
        foreach ($path in $Attachments) {
            if (Test-Path $path) {
                $file = New-Object MimeKit.MimePart
                $stream = [System.IO.File]::OpenRead($path)
                $attachmentStreams.Add($stream)
                $content = New-Object MimeKit.MimeContent($stream)
                $file.Content = $content
                $file.FileName = [System.IO.Path]::GetFileName($path)
                $file.ContentDisposition = New-Object MimeKit.ContentDisposition
                $file.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64
                $bodyContainer.Add($file)
            }
        }

        $message.Body = $bodyContainer

        # SMTP client
        $smtp = New-Object MailKit.Net.Smtp.SmtpClient
        $smtp.Connect($SmtpServerName, $SmtpPort, [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType)

        if ($Credential) {
            $smtp.Authenticate($Credential.UserName, $Credential.GetNetworkCredential().Password)
        }

        $smtp.Send($message)
    }
    finally {
        if ($smtp) {
            $smtp.Disconnect($true)
            $smtp.Dispose()
        }

        foreach ($stream in $attachmentStreams) {
            if ($stream) { $stream.Dispose() }
        }
    }
}

function Save-MailBodyToLogHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailParams,

        [Parameter(Mandatory)]
        $LogFolder
    )

    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return }

    # Replace any character that is invalid in a file name with a space.
    # Splitting on the invalid-char set avoids the $OFS-dependent string cast.
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safeSubject = ([string]$MailParams.Subject).Split($invalid) -join ' '

    $path = Join-Path $LogFolder ('Mail - {0}.html' -f $safeSubject)

    $MailParams.Body | Out-File -LiteralPath $path -Encoding utf8 -Force

    return $path
}