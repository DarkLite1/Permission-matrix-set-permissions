#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

<#
    Pester 5 tests for Modules\PermissionMatrix\Private\Mail.ps1

    Covered functions:
        - Get-MailRecipientListHC   (pure)
        - Get-MailSubjectHC         (pure)
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
            * The enum literals the function references
              ([MailKit.Security.SecureSocketOptions], [MimeKit.ContentEncoding])
              and the MimeKit.MailboxAddress type (whose static ::Parse the
              function now calls for To/Bcc) are stubbed via Add-Type only when
              the real assemblies are not already loaded.
        A cleaner long-term seam would be to refactor the function to accept
        an injectable SMTP client; that would let the test mock just the send.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    . "$moduleRoot\Private\Mail.ps1"

    # Prefer the real MailKit / MimeKit assemblies when their paths are known
    # (the script reads the same env vars). Loading the real types means the
    # tests exercise the actual [MimeKit.MailboxAddress]::Parse contract and,
    # crucially, leave NO conflicting stub type behind. The lightweight stubs
    # below are only a fallback for environments without the DLLs (e.g. CI).
    #
    # The stub MimeKit.MailboxAddress does not derive from InternetAddress, so
    # if it is compiled into a session it will shadow the real type and break
    # any real send done later in the same process. Always run these tests in
    # their own PowerShell process, e.g.:
    #   pwsh -NoProfile -Command "Invoke-Pester -Path .\Tests\Unit\Private\Mail.Tests.ps1"
    foreach ($dll in @($env:MIMEKIT_DLL, $env:MAILKIT_DLL)) {
        if ($dll -and (Test-Path -LiteralPath $dll)) {
            Add-Type -Path $dll -ErrorAction SilentlyContinue
        }
    }

    # Stub the types the sender references, but only if the real MailKit /
    # MimeKit assemblies are not already present in the session.
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

    # The function calls the static [MimeKit.MailboxAddress]::Parse for To/Bcc
    # and the (name, address) constructor for From, so the stub provides both
    # plus the .Name / .Address properties the tests read.
    if (-not ('MimeKit.MailboxAddress' -as [type])) {
        Add-Type -TypeDefinition @'
namespace MimeKit {
    public class MailboxAddress {
        public string Name { get; set; }
        public string Address { get; set; }
        public MailboxAddress(string name, string address) {
            Name = name;
            Address = address;
        }
        public static MailboxAddress Parse(string text) {
            return new MailboxAddress(null, text);
        }
    }
}
'@
    }
}

Describe 'Get-MailRecipientListHC' {
    It 'merges the To list with the defaults file recipients' {
        $settings = [PSCustomObject]@{ To = @('bob@example.com') }

        $result = Get-MailRecipientListHC -SendMailSettings $settings -DefaultsMailTo @('amy@example.com')

        @($result) | Should -HaveCount 2
        $result | Should -Contain 'bob@example.com'
        $result | Should -Contain 'amy@example.com'
    }

    It 'trims surrounding whitespace from addresses' {
        $settings = [PSCustomObject]@{ To = @('  bob@example.com  ') }

        $result = Get-MailRecipientListHC -SendMailSettings $settings

        $result | Should -Be 'bob@example.com'
    }

    It 'drops empty and whitespace-only entries' {
        $settings = [PSCustomObject]@{ To = @('amy@example.com', '', '   ') }

        $result = Get-MailRecipientListHC -SendMailSettings $settings

        $result | Should -Be 'amy@example.com'
    }

    It 'ignores a null entry in the list instead of throwing' {
        $settings = [PSCustomObject]@{ To = @('amy@example.com', $null) }

        $result = Get-MailRecipientListHC -SendMailSettings $settings

        $result | Should -Be 'amy@example.com'
    }

    It 'removes duplicates and returns the list sorted' {
        $settings = [PSCustomObject]@{ To = @('zoe@example.com', 'amy@example.com', 'zoe@example.com') }

        $result = Get-MailRecipientListHC -SendMailSettings $settings

        @($result) | Should -HaveCount 2
        $result[0] | Should -Be 'amy@example.com'
        $result[1] | Should -Be 'zoe@example.com'
    }

    It 'returns nothing when there are no recipients' {
        $settings = [PSCustomObject]@{ To = @() }

        $result = Get-MailRecipientListHC -SendMailSettings $settings

        $result | Should -BeNullOrEmpty
    }

    It 'works when only the defaults file recipients are supplied' {
        $settings = [PSCustomObject]@{}

        $result = Get-MailRecipientListHC -SendMailSettings $settings -DefaultsMailTo 'amy@example.com'

        $result | Should -Be 'amy@example.com'
    }
}

Describe 'Get-MailSubjectHC' {
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
            $result = Get-MailSubjectHC -SystemErrors (New-SystemErrors 1) `
                -Counter $zeroCounter -MatrixCount 1

            $result | Should -Be '1 matrix file, 1 System Error'
        }

        It 'pluralises matrix files and system errors' {
            $result = Get-MailSubjectHC -SystemErrors (New-SystemErrors 3) `
                -Counter $zeroCounter -MatrixCount 2

            $result | Should -Be '2 matrix files, 3 System Errors'
        }

        It 'ignores counter errors and warnings when system errors exist' {
            $counter = [PSCustomObject]@{ TotalErrors = 9; TotalWarnings = 9 }

            $result = Get-MailSubjectHC -SystemErrors (New-SystemErrors 1) `
                -Counter $counter -MatrixCount 1

            $result | Should -Be '1 matrix file, 1 System Error'
        }

        It 'appends the custom subject' {
            $result = Get-MailSubjectHC -SystemErrors (New-SystemErrors 2) `
                -Counter $zeroCounter -MatrixCount 1 -CustomSubject 'Nightly run'

            $result | Should -Be '1 matrix file, 2 System Errors, Nightly run'
        }
    }

    Context 'when there are no system errors' {
        It 'reports only the matrix count when there are no errors or warnings' {
            $result = Get-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
                -Counter $zeroCounter -MatrixCount 1

            $result | Should -Be '1 matrix file'
        }

        It 'pluralises the matrix count' {
            $result = Get-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
                -Counter $zeroCounter -MatrixCount 3

            $result | Should -Be '3 matrix files'
        }

        It 'pluralises a zero matrix count' {
            $result = Get-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
                -Counter $zeroCounter -MatrixCount 0

            $result | Should -Be '0 matrix files'
        }

        It 'appends the custom subject' {
            $result = Get-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
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

            $result = Get-MailSubjectHC -SystemErrors (New-SystemErrors 0) `
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
                    # Use the real (or stubbed) type so these New-Object
                    # instances and the static ::Parse the function calls for
                    # To/Bcc share one shape with .Name and .Address.
                    [MimeKit.MailboxAddress]::new($ArgumentList[0], $ArgumentList[1])
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

        $script:LastMessage.To.Items.Address | Should -Contain 'a@example.com'
        $script:LastMessage.To.Items.Address | Should -Contain 'b@example.com'
        $script:LastMessage.Bcc.Items.Address | Should -Contain 'c@example.com'

        # No attachments here, so the body is the HTML TextPart itself,
        # not a multipart container.
        $script:LastMessage.Body.Kind | Should -Be 'TextPart'
        $script:LastMessage.Body.Text | Should -Be '<p>hi</p>'
        $script:LastMessage.Body.Subtype | Should -Be 'html'
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

    Context 'parameter validation' {
        It 'rejects an invalid priority' {
            { Send-MailKitMessageHC @params -Priority 'Urgent' } | Should -Throw
        }

        It 'rejects an invalid connection type' {
            { Send-MailKitMessageHC @params -SmtpConnectionType 'Bogus' } | Should -Throw
        }
    }

    Context 'assembly loading' {
        It 'throws a clear, assembly-specific error when MimeKit fails to load' {
            Mock Add-Type { throw 'Could not load file or assembly.' } `
                -ParameterFilter { $Path -eq 'X:\MimeKit.dll' }

            { Send-MailKitMessageHC @params } |
            Should -Throw -ExpectedMessage '*MimeKit*X:\MimeKit.dll*'
        }

        It 'throws a clear, assembly-specific error when MailKit fails to load' {
            Mock Add-Type { throw 'Bad image format.' } `
                -ParameterFilter { $Path -eq 'X:\MailKit.dll' }

            { Send-MailKitMessageHC @params } |
            Should -Throw -ExpectedMessage '*MailKit*X:\MailKit.dll*'
        }

        It 'throws a clear error when an assembly path is only whitespace' {
            $params.MimeKitAssemblyPath = '   '

            { Send-MailKitMessageHC @params } |
            Should -Throw -ExpectedMessage '*MimeKit*not set*'
        }
    }

    Context 'attachments' {
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

        It 'uses the HTML body part directly as the message body when there are no attachments' {
            Send-MailKitMessageHC @params

            $script:LastMessage.Body.Kind | Should -Be 'TextPart'
            $script:LastMessage.Body.Text | Should -Be '<p>hi</p>'
        }

        It 'disposes the attachment stream after sending' {
            $existing = Join-Path $TestDrive 'disposed.txt'
            'x' | Out-File -LiteralPath $existing

            Send-MailKitMessageHC @params -Attachments @($existing)

            # The fake MimeContent holds the real FileStream the function opened;
            # a disposed FileStream reports CanRead = $false.
            $attachment = $script:LastMessage.Body.Parts | Where-Object { $_.Kind -eq 'MimePart' }
            $attachment.Content.Stream.CanRead | Should -BeFalse
        }
    }
}