function Build-MailParametersHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Settings,

        [Parameter(Mandatory)]
        $Html,

        [Parameter(Mandatory)]
        $ExportedFiles,

        [Parameter(Mandatory)]
        $Counter,

        [Parameter(Mandatory)]
        $SystemErrors,

        [Parameter(Mandatory)]
        $MatrixCount,

        [Parameter(Mandatory)]
        $MailToDefaultsFile,

        [Parameter(Mandatory)]
        $LogFolder,

        [Parameter(Mandatory)]
        $ScriptStartTime
    )

    $params = @{}

    # Recipients
    $params.To = Generate-MailRecipientListHC `
        -SendMailSettings $Settings.SendMail `
        -MailToDefaultsFile $MailToDefaultsFile

    if ($Settings.SendMail.Bcc) {
        $params.Bcc = $Settings.SendMail.Bcc
    }

    # Base mail settings
    $params.From = Get-StringValueHC $Settings.SendMail.From
    $params.FromDisplayName = Get-StringValueHC $Settings.SendMail.FromDisplayName
    $params.SmtpServerName = Get-StringValueHC $Settings.SendMail.Smtp.ServerName
    $params.SmtpPort = Get-StringValueHC $Settings.SendMail.Smtp.Port
    $params.SmtpConnectionType = Get-StringValueHC $Settings.SendMail.Smtp.ConnectionType
    $params.MailKitAssemblyPath = Get-StringValueHC $Settings.SendMail.AssemblyPath.MailKit
    $params.MimeKitAssemblyPath = Get-StringValueHC $Settings.SendMail.AssemblyPath.MimeKit

    # Subject
    $params.Subject = Generate-MailSubjectHC `
        -SystemErrors $SystemErrors `
        -Counter $Counter `
        -MatrixCount $MatrixCount `
        -CustomSubject $Settings.SendMail.Subject

    # Attachments get populated later by Write-SystemErrorLogHC if needed
    $params.Attachments = @()

    # Priority (High if errors or warnings)
    if ($SystemErrors.Count -gt 0 -or $Counter.TotalErrors -gt 0 -or $Counter.TotalWarnings -gt 0) {
        $params.Priority = 'High'
    }
    else {
        $params.Priority = 'Normal'
    }

    # Mail Body
    $params.Body = $Html

    # Create SMTP credential
    $smtpUser = Get-StringValueHC $Settings.SendMail.Smtp.UserName
    $smtpPass = Get-StringValueHC $Settings.SendMail.Smtp.Password

    if ($smtpUser -and $smtpPass) {
        $secure = ConvertTo-SecureString -String $smtpPass -AsPlainText -Force
        $params.Credential = New-Object System.Management.Automation.PSCredential($smtpUser, $secure)
    }

    return $params
}

function Generate-MailRecipientListHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $SendMailSettings,

        [Parameter(Mandatory)]
        $MailToDefaultsFile
    )

    $list = @()
    if ($SendMailSettings.To) { $list += $SendMailSettings.To }
    if ($MailToDefaultsFile) { $list += $MailToDefaultsFile }

    return ($list | ForEach-Object { $_.Trim() } | Where-Object { $_ }) | Sort-Object -Unique
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
        [string]$Priority = 'Normal',
        [string[]]$Attachments,
        [string]$SmtpConnectionType = 'None',
        [PSCredential]$Credential
    )

    # Load assemblies
    Add-Type -Path $MimeKitAssemblyPath
    Add-Type -Path $MailKitAssemblyPath

    # Create message
    $message = New-Object MimeKit.MimeMessage
    $fromAddress = New-Object MimeKit.MailboxAddress($FromDisplayName, $From)
    $message.From.Add($fromAddress)

    foreach ($t in $To) {
        $message.To.Add($t)
    }
    foreach ($b in $Bcc) {
        $message.Bcc.Add($b)
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
    try {
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

    $safeSubject = $MailParams.Subject
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $pattern = [Regex]::Escape(([string]$invalid))
    $safeSubject = $safeSubject -replace "[$pattern]", ' '

    $path = Join-Path $LogFolder ('Mail - {0}.html' -f $safeSubject)

    $MailParams.Body | Out-File -LiteralPath $path -Encoding utf8 -Force

    return $path
}