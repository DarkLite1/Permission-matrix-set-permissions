#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

<#
    Pester 5 tests for Modules\PermissionMatrix\Private\Mail.ps1

    Covered functions:
        - Generate-MailRecipientListHC   (pure)
        - Generate-MailSubjectHC         (pure)
        - Save-MailBodyToLogHC           (file system, uses TestDrive)
        - Send-MailKitMessageHC          (stubbed mail stack, see note below)

    NOTE on Send-MailKitMessageHC:
        The function instantiates concrete MailKit / MimeKit types with
        New-Object and calls instance methods (.Connect/.Authenticate/.Send/
        .Disconnect/.Dispose) directly. There is no command to mock for an
        instance method, so to "mock the send and assert it is called right"
        the whole mail layer is replaced:
            * Add-Type is mocked to a no-op (no real DLLs required).
            * New-Object is mocked to return recording fakes. The fake
              SmtpClient records its Connect/Authenticate/Send/Disconnect/
              Dispose calls so the tests can assert on them.
            * The two enum literals the function references
              ([MailKit.Security.SecureSocketOptions] and
              [MimeKit.ContentEncoding]) are stubbed via Add-Type only when
              the real assemblies are not already loaded.
        A cleaner long-term seam would be to refactor the function to accept
        an injectable SMTP client; that would let the test mock just the send.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    . "$moduleRoot\Private\Mail.ps1"

    # Stub the two enum types the sender references, but only if the real
    # MailKit / MimeKit assemblies are not already present in the session.
    if (-not ('MailKit.Security.SecureSocketOptions' -as [type])) {
        Add-Type -TypeDefinition @'
namespace MailKit.Security {
    public enum SecureSocketOptions {
        None, Auto, SslOnConnect, StartTls, StartTlsWhenAvailable
    }
}
'@
    }

    if (-not ('MimeKit.ContentEncoding' -as [type])) {
        Add-Type -TypeDefinition @'
namespace MimeKit {
    public enum ContentEncoding {
        Default, Binary, EightBit, SevenBit, Base64, QuotedPrintable, UUEncode
    }
}
'@
    }
}

Describe 'Generate-MailRecipientListHC' {
    It 'merges the To list with the defaults file recipients' {
        $settings = [PSCustomObject]@{ To = @('bob@example.com') }

        $result = Generate-MailRecipientListHC -SendMailSettings $settings -MailToDefaultsFile @('amy@example.com')

        @($result) | Should -HaveCount 2
        $result | Should -Contain 'bob@example.com'
        $result | Should -Contain 'amy@example.com'
    }

    It 'trims surrounding whitespace from addresses' {
        $settings = [PSCustomObject]@{ To = @('  bob@example.com  ') }

        $result = Generate-MailRecipientListHC -SendMailSettings $settings

        $result | Should -Be 'bob@example.com'
    }

    It 'drops empty and whitespace-only entries' {
        $settings = [PSCustomObject]@{ To = @('amy@example.com', '', '   ') }

        $result = Generate-MailRecipientListHC -SendMailSettings $settings

        $result | Should -Be 'amy@example.com'
    }

    It 'removes duplicates and returns the list sorted' {
        $settings = [PSCustomObject]@{ To = @('zoe@example.com', 'amy@example.com', 'zoe@example.com') }

        $result = Generate-MailRecipientListHC -SendMailSettings $settings

        @($result) | Should -HaveCount 2
        $result[0] | Should -Be 'amy@example.com'
        $result[1] | Should -Be 'zoe@example.com'
    }

    It 'returns nothing when there are no recipients' {
        $settings = [PSCustomObject]@{ To = @() }

        $result = Generate-MailRecipientListHC -SendMailSettings $settings

        $result | Should -BeNullOrEmpty
    }

    It 'works when only the defaults file recipients are supplied' {
        $settings = [PSCustomObject]@{}

        $result = Generate-MailRecipientListHC -SendMailSettings $settings -MailToDefaultsFile 'amy@example.com'

        $result | Should -Be 'amy@example.com'
    }
}

Describe 'Generate-MailSubjectHC' {
    BeforeAll {
        # A counter with no errors / warnings, reused where the branch ignores it.
        $script:zeroCounter = [PSCustomObject]@{ TotalErrors = 0; TotalWarnings = 0 }

        function script:New-SystemErrors {
            param([int]$Count)
            $list = [System.Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $Count; $i++) { $list.Add([PSCustomObject]@{ N = $i }) }
            , $list
        }
    }

    Context 'when system errors are present (takes priority over the counter)' {
        It 'reports a single matrix file and a single system error' {
            $result = Generate-MailSubjectHC -SystemErrors (New-SystemErrors 1) `
                -Counter $zeroCounter -MatrixCount 1

            $result | Should -Be '1 matrix file, 1 System Error'
        }

        It 'pluralises matrix files and system errors' {
            $result = Generate-MailSubjectHC -SystemErrors (New-SystemErrors 3) `
                -Counter $zeroCounter -MatrixCount 2

            $result | Should -Be '2 matrix files, 3 System Errors'
        }

        It 'ignores counter errors and warnings when system errors exist' {
            $counter = [PSCustomObject]@{ TotalErrors = 9; TotalWarnings = 9 }

            $result = Generate-MailSubjectHC -SystemErrors (New-SystemErrors 1) `
                -Counter $counter -MatrixCount 1

            $result | Should -Be '1 matrix file, 1 System Error'
        }

        It 'appends the custom subject' {
            $result = Generate-MailSubjectHC -SystemErrors (New-SystemErrors 2) `
                -Counter $zeroCounter -MatrixCount 1 -CustomSubject 'Nightly run'

            $result | Should -Be '1 matrix file, 2 System Errors, Nightly run'
        }
    }

    Context 'when there are no system errors' {
        It 'reports only the matrix count when there are no errors or warnings' {
            $result = Generate-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
                -Counter $zeroCounter -MatrixCount 1

            $result | Should -Be '1 matrix file'
        }

        It 'pluralises the matrix count' {
            $result = Generate-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
                -Counter $zeroCounter -MatrixCount 3

            $result | Should -Be '3 matrix files'
        }

        It 'pluralises a zero matrix count' {
            $result = Generate-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
                -Counter $zeroCounter -MatrixCount 0

            $result | Should -Be '0 matrix files'
        }

        It 'appends the custom subject' {
            $result = Generate-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
                -Counter $zeroCounter -MatrixCount 1 -CustomSubject 'All good'

            $result | Should -Be '1 matrix file, All good'
        }

        It 'builds the subject for <Errors> error(s) and <Warnings> warning(s): <Expected>' -TestCases @(
            @{ Errors = 1; Warnings = 0; Expected = '1 matrix file, 1 error' }
            @{ Errors = 2; Warnings = 0; Expected = '1 matrix file, 2 errors' }
            @{ Errors = 0; Warnings = 1; Expected = '1 matrix file, 1 warning' }
            @{ Errors = 0; Warnings = 2; Expected = '1 matrix file, 2 warnings' }
            @{ Errors = 1; Warnings = 1; Expected = '1 matrix file, 1 error, 1 warning' }
            @{ Errors = 2; Warnings = 3; Expected = '1 matrix file, 2 errors, 3 warnings' }
        ) {
            param($Errors, $Warnings, $Expected)

            $counter = [PSCustomObject]@{ TotalErrors = $Errors; TotalWarnings = $Warnings }

            $result = Generate-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
                -Counter $counter -MatrixCount 1

            $result | Should -Be $Expected
        }
    }
}

Describe 'Save-MailBodyToLogHC' {
    It 'writes the body to a Mail - <subject>.html file and returns its path' {
        $params = @{ Subject = 'Daily report'; Body = '<html><body>Hello</body></html>' }

        $result = Save-MailBodyToLogHC -MailParams $params -LogFolder $TestDrive

        $result | Should -Be (Join-Path $TestDrive 'Mail - Daily report.html')
        $result | Should -Exist
        Get-Content -LiteralPath $result -Raw | Should -BeLike '*<html><body>Hello</body></html>*'
    }

    It 'replaces characters that are invalid in a file name with a space' {
        # Pick a real invalid char for the current OS so the test is portable.
        $invalidChar = [System.IO.Path]::GetInvalidFileNameChars() |
            Where-Object { $_ -ne ' ' } | Select-Object -First 1
        $params = @{ Subject = "report${invalidChar}name"; Body = '<p>x</p>' }

        $result = Save-MailBodyToLogHC -MailParams $params -LogFolder $TestDrive

        [System.IO.Path]::GetFileName($result) | Should -Be 'Mail - report name.html'
    }

    It 'returns nothing when the log folder does not exist' {
        $missing = Join-Path $TestDrive 'does-not-exist'
        $params = @{ Subject = 'x'; Body = 'y' }

        $result = Save-MailBodyToLogHC -MailParams $params -LogFolder $missing

        $result | Should -BeNullOrEmpty
    }

    It 'returns nothing when the log folder path is a file, not a directory' {
        $filePath = Join-Path $TestDrive 'a-file.txt'
        'not a folder' | Out-File -LiteralPath $filePath
        $params = @{ Subject = 'x'; Body = 'y' }

        $result = Save-MailBodyToLogHC -MailParams $params -LogFolder $filePath

        $result | Should -BeNullOrEmpty
    }
}

Describe 'Send-MailKitMessageHC' {
    BeforeAll {
        # No real assemblies needed.
        Mock Add-Type {}

        # Replace the whole MailKit / MimeKit object graph with recording fakes.
        # IMPORTANT: never call New-Object inside this mock (it would recurse) —
        # build everything with [PSCustomObject] and Add-Member.
        Mock New-Object {
            # Helper that builds an "addable list" (mimics InternetAddressList).
            $newAddableList = {
                $l = [PSCustomObject]@{ Items = [System.Collections.Generic.List[object]]::new() }
                $l | Add-Member -MemberType ScriptMethod -Name Add `
                    -Value { param($x) $this.Items.Add($x) } -PassThru
            }

            switch ($TypeName) {
                'MimeKit.MimeMessage' {
                    $headers = [PSCustomObject]@{ Items = [System.Collections.Generic.List[object]]::new() }
                    $headers | Add-Member -MemberType ScriptMethod -Name Add `
                        -Value { param($n, $v) $this.Items.Add([PSCustomObject]@{ Name = $n; Value = $v }) }

                    $msg = [PSCustomObject]@{
                        From    = (& $newAddableList)
                        To      = (& $newAddableList)
                        Bcc     = (& $newAddableList)
                        Headers = $headers
                        Subject = $null
                        Body    = $null
                    }
                    $script:LastMessage = $msg
                    $msg
                }
                'MimeKit.MailboxAddress' {
                    [PSCustomObject]@{ Name = $ArgumentList[0]; Address = $ArgumentList[1] }
                }
                'MimeKit.TextPart' {
                    [PSCustomObject]@{ Kind = 'TextPart'; Subtype = $ArgumentList[0]; Text = $null }
                }
                'MimeKit.Multipart' {
                    $c = [PSCustomObject]@{
                        Kind    = 'Multipart'
                        Subtype = $ArgumentList[0]
                        Parts   = [System.Collections.Generic.List[object]]::new()
                    }
                    $c | Add-Member -MemberType ScriptMethod -Name Add `
                        -Value { param($p) $this.Parts.Add($p) }
                    $c
                }
                'MimeKit.MimePart' {
                    [PSCustomObject]@{
                        Kind                    = 'MimePart'
                        FileName                = $null
                        Content                 = $null
                        ContentDisposition      = $null
                        ContentTransferEncoding = $null
                    }
                }
                'MimeKit.MimeContent' {
                    [PSCustomObject]@{ Kind = 'MimeContent'; Stream = $ArgumentList[0] }
                }
                'MimeKit.ContentDisposition' {
                    [PSCustomObject]@{ Kind = 'ContentDisposition' }
                }
                'MailKit.Net.Smtp.SmtpClient' {
                    $client = [PSCustomObject]@{
                        ConnectArgs  = $null
                        AuthArgs     = $null
                        SentMessage  = $null
                        SendCount    = 0
                        Disconnected = $false
                        Disposed     = $false
                    }
                    $client | Add-Member -MemberType ScriptMethod -Name Connect `
                        -Value { param($s, $p, $o) $this.ConnectArgs = [PSCustomObject]@{ Server = $s; Port = $p; Options = $o } }
                    $client | Add-Member -MemberType ScriptMethod -Name Authenticate `
                        -Value { param($u, $pw) $this.AuthArgs = [PSCustomObject]@{ UserName = $u; Password = $pw } }
                    $client | Add-Member -MemberType ScriptMethod -Name Send `
                        -Value { param($m) $this.SentMessage = $m; $this.SendCount++ }
                    $client | Add-Member -MemberType ScriptMethod -Name Disconnect `
                        -Value { param($q) $this.Disconnected = $true }
                    $client | Add-Member -MemberType ScriptMethod -Name Dispose `
                        -Value { $this.Disposed = $true }
                    $script:LastSmtpClient = $client
                    $client
                }
                default {
                    [PSCustomObject]@{ Kind = $TypeName }
                }
            }
        }
    }

    BeforeEach {
        $script:LastMessage = $null
        $script:LastSmtpClient = $null

        $params = @{
            MailKitAssemblyPath = 'X:\MailKit.dll'
            MimeKitAssemblyPath = 'X:\MimeKit.dll'
            SmtpServerName      = 'smtp.example.com'
            SmtpPort            = 587
            Subject             = 'Test subject'
            Body                = '<p>hi</p>'
            From                = 'from@example.com'
            FromDisplayName     = 'Sender Name'
            To                  = @('a@example.com', 'b@example.com')
            Bcc                 = @('c@example.com')
        }
    }

    It 'loads the MimeKit and MailKit assemblies' {
        Send-MailKitMessageHC @params

        Should -Invoke Add-Type -Exactly -Times 2
        Should -Invoke Add-Type -Times 1 -ParameterFilter { $Path -eq 'X:\MimeKit.dll' }
        Should -Invoke Add-Type -Times 1 -ParameterFilter { $Path -eq 'X:\MailKit.dll' }
    }

    It 'connects to the SMTP server with the configured server, port and connection type' {
        Send-MailKitMessageHC @params -SmtpConnectionType 'StartTls'

        $script:LastSmtpClient.ConnectArgs.Server | Should -Be 'smtp.example.com'
        $script:LastSmtpClient.ConnectArgs.Port | Should -Be 587
        $script:LastSmtpClient.ConnectArgs.Options.ToString() | Should -Be 'StartTls'
    }

    It 'defaults the connection type to None' {
        Send-MailKitMessageHC @params

        $script:LastSmtpClient.ConnectArgs.Options.ToString() | Should -Be 'None'
    }

    It 'sends the message exactly once and then disconnects and disposes' {
        Send-MailKitMessageHC @params

        $script:LastSmtpClient.SendCount | Should -Be 1
        $script:LastSmtpClient.SentMessage | Should -Be $script:LastMessage
        $script:LastSmtpClient.Disconnected | Should -BeTrue
        $script:LastSmtpClient.Disposed | Should -BeTrue
    }

    It 'authenticates when a credential is supplied' {
        $cred = [PSCredential]::new('svc-user', (ConvertTo-SecureString 'secret-pw' -AsPlainText -Force))

        Send-MailKitMessageHC @params -Credential $cred

        $script:LastSmtpClient.AuthArgs.UserName | Should -Be 'svc-user'
        $script:LastSmtpClient.AuthArgs.Password | Should -Be 'secret-pw'
    }

    It 'does not authenticate when no credential is supplied' {
        Send-MailKitMessageHC @params

        $script:LastSmtpClient.AuthArgs | Should -BeNullOrEmpty
    }

    It 'sets the sender, recipients, subject and body on the message' {
        Send-MailKitMessageHC @params

        $script:LastMessage.Subject | Should -Be 'Test subject'

        $script:LastMessage.From.Items[0].Address | Should -Be 'from@example.com'
        $script:LastMessage.From.Items[0].Name | Should -Be 'Sender Name'

        $script:LastMessage.To.Items | Should -Contain 'a@example.com'
        $script:LastMessage.To.Items | Should -Contain 'b@example.com'
        $script:LastMessage.Bcc.Items | Should -Contain 'c@example.com'

        $bodyPart = $script:LastMessage.Body.Parts | Where-Object { $_.Kind -eq 'TextPart' }
        $bodyPart.Text | Should -Be '<p>hi</p>'
        $bodyPart.Subtype | Should -Be 'html'
    }

    It 'sets the X-Priority header to <Expected> for priority <Priority>' -TestCases @(
        @{ Priority = 'High'; Expected = '1 (Highest)' }
        @{ Priority = 'Normal'; Expected = '3 (Normal)' }
        @{ Priority = 'Low'; Expected = '5 (Lowest)' }
    ) {
        param($Priority, $Expected)

        Send-MailKitMessageHC @params -Priority $Priority

        $header = $script:LastMessage.Headers.Items | Where-Object { $_.Name -eq 'X-Priority' }
        $header.Value | Should -Be $Expected
    }

    Context 'attachments' {
        # The function opens attachment streams with [System.IO.File]::OpenRead
        # and never disposes them, so the file stays locked on Windows and
        # Pester cannot remove TestDrive. Release any captured stream here.
        AfterEach {
            if ($script:LastMessage -and $script:LastMessage.Body -and $script:LastMessage.Body.Parts) {
                foreach ($part in $script:LastMessage.Body.Parts) {
                    if ($part.Kind -eq 'MimePart' -and $part.Content -and $part.Content.Stream) {
                        try { $part.Content.Stream.Dispose() } catch { }
                    }
                }
            }
        }

        It 'includes an existing attachment and skips a missing one' {
            $existing = Join-Path $TestDrive 'note.txt'
            'attachment body' | Out-File -LiteralPath $existing
            $missing = Join-Path $TestDrive 'gone.txt'

            Send-MailKitMessageHC @params -Attachments @($existing, $missing)

            # Body container holds the HTML body part plus one attachment part.
            $script:LastMessage.Body.Parts | Should -HaveCount 2

            $attachment = $script:LastMessage.Body.Parts | Where-Object { $_.Kind -eq 'MimePart' }
            $attachment.FileName | Should -Be 'note.txt'
            $attachment.ContentTransferEncoding.ToString() | Should -Be 'Base64'
        }

        It 'adds only the HTML body part when there are no attachments' {
            Send-MailKitMessageHC @params

            $script:LastMessage.Body.Parts | Should -HaveCount 1
            $script:LastMessage.Body.Parts[0].Kind | Should -Be 'TextPart'
        }
    }
}