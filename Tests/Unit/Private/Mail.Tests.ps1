#requires -Modules Pester

Describe 'Mail.ps1 - Mail Handling' {

    BeforeAll {
        # Dot-source the Mail.ps1 file
        $modulePath = Split-Path -Parent $MyInvocation.MyCommand.Path
        $mailFile = Join-Path $modulePath '../Modules/Toolbox.PermissionMatrixHC/Private/Mail.ps1'
        . $mailFile
    }

    Context 'Recipient generation' {

        It 'Combines SendMail.To and MailToDefaultsFile uniquely' {
            $settings = @{
                To = @('a@x.com', 'b@x.com')
            }

            $defaults = @('b@x.com', 'c@x.com')

            $result = Generate-MailRecipientListHC -SendMailSettings $settings -MailToDefaultsFile $defaults

            $result | Should -Contain 'a@x.com'
            $result | Should -Contain 'b@x.com'
            $result | Should -Contain 'c@x.com'
            $result.Count | Should -Be 3
        }
    }

    Context 'Mail subject generation' {

        It 'Generates subject with system errors' {
            $sys = @(1, 2)
            $counter = @{
                TotalErrors   = 2
                TotalWarnings = 1
            }

            $result = Generate-MailSubjectHC -SystemErrors $sys -Counter $counter -MatrixCount 5 -CustomSubject 'Test'
            $result | Should -Match '5 matrix file'
            $result | Should -Match '2 System Error'
            $result | Should -Match 'Test'
        }

        It 'Generates subject without system errors' {
            $sys = @()
            $counter = @{
                TotalErrors   = 1
                TotalWarnings = 2
            }

            $result = Generate-MailSubjectHC -SystemErrors $sys -Counter $counter -MatrixCount 3 -CustomSubject $null

            $result | Should -Match '3 matrix file'
            $result | Should -Match '1 error'
            $result | Should -Match '2 warning'
        }
    }

    Context 'Build-MailParametersHC' {

        It 'Creates correct mail parameters' {

            $settings = @{
                SendMail = @{
                    From            = 'a@x.com'
                    FromDisplayName = 'Automated'
                    To              = @('x@y.com')
                    Bcc             = @('z@y.com')
                    Subject         = 'Custom'
                    Body            = '<p>Test</p>'
                    Smtp            = @{
                        ServerName     = 'smtp.server'
                        Port           = 25
                        ConnectionType = 'None'
                        UserName       = 'user'
                        Password       = 'pass'
                    }
                    AssemblyPath    = @{
                        MailKit = 'C:\MailKit.dll'
                        MimeKit = 'C:\MimeKit.dll'
                    }
                }
            }

            $html = '<html>content</html>'
            $exported = @{}
            $counter = @{
                TotalErrors   = 0
                TotalWarnings = 0
            }
            $sysErr = @()
            $mailToDefaults = @('default@y.com')

            $params = Build-MailParametersHC `
                -Settings $settings `
                -Html $html `
                -ExportedFiles $exported `
                -Counter $counter `
                -SystemErrors $sysErr `
                -MatrixCount 4 `
                -MailToDefaultsFile $mailToDefaults `
                -LogFolder 'C:\Logs' `
                -ScriptStartTime (Get-Date)

            $params.From | Should -Be 'a@x.com'
            $params.To | Should -Contain 'x@y.com'
            $params.To | Should -Contain 'default@y.com'
            $params.Subject | Should -Match '4 matrix file'
            $params.Body | Should -Be $html
        }
    }

    Context 'Send-MailKitMessageHC' {

        It 'Calls MailKit but is mocked during test' {
            Mock Add-Type
            Mock New-Object -MockWith {
                # Return dummy object for MimeMessage
                New-Object PSObject -Property @{
                    From    = New-Object System.Collections.ArrayList
                    To      = New-Object System.Collections.ArrayList
                    Bcc     = New-Object System.Collections.ArrayList
                    Headers = @{}
                }
            } -ParameterFilter { $TypeName -eq 'MimeKit.MimeMessage' }

            # Dummy SMTP client
            $smtpMock = New-Object PSObject -Property @{
                Connect      = { param($s, $p, $so) }
                Authenticate = { }
                Send         = { }
                Disconnect   = { }
                Dispose      = { }
            }

            Mock New-Object -MockWith { $smtpMock } -ParameterFilter { $TypeName -eq 'MailKit.Net.Smtp.SmtpClient' }

            { 
                Send-MailKitMessageHC `
                    -MailKitAssemblyPath 'x' `
                    -MimeKitAssemblyPath 'y' `
                    -SmtpServerName 'server' `
                    -SmtpPort 25 `
                    -Body '<p>body</p>' `
                    -Subject 'Test' `
                    -From 'sender@x.com' `
                    -To 'receiver@x.com'
            } | Should -Not -Throw
        }
    }

    Context 'Save-MailBodyToLogHC' {

        It 'Saves mail HTML to the log folder' {
            $folder = Join-Path $TestDrive 'log'
            New-Item -ItemType Directory -Path $folder | Out-Null

            $params = @{
                Subject = 'Test Subject'
                Body    = '<html>body</html>'
            }

            $path = Save-MailBodyToLogHC -MailParams $params -LogFolder $folder
            Test-Path $path | Should -BeTrue
        }
    }
}