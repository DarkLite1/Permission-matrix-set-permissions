#requires -Modules Pester

Describe 'Logging.ps1 - Consolidated Logging' {

    BeforeAll {
        $root = Split-Path -Parent $MyInvocation.MyCommand.Path
        $log = Join-Path $root '../Modules/Toolbox.PermissionMatrixHC/Private/Logging.ps1'
        $utils = Join-Path $root '../Modules/Toolbox.PermissionMatrixHC/Private/Utils.ps1'
        . $utils
        . $log
    }


    # ----------------------------------------------------------------------
    Context 'Cleanup-OldLogsHC' {

        It 'Removes old files and keeps recent ones' {
            $sys = @()

            $folder = Join-Path $TestDrive 'logs'
            New-Item -ItemType Directory -Path $folder | Out-Null

            $oldFile = Join-Path $folder 'old.txt'
            Set-Content $oldFile 'x'
            (Get-Item $oldFile).CreationTime = (Get-Date).AddDays(-50)

            $newFile = Join-Path $folder 'new.txt'
            Set-Content $newFile 'y'
            (Get-Item $newFile).CreationTime = (Get-Date).AddDays(-1)

            Cleanup-OldLogsHC `
                -LogFolder $folder `
                -RetentionDays 30 `
                -SystemErrors ([ref]$sys)

            Test-Path $oldFile | Should -BeFalse
            Test-Path $newFile | Should -BeTrue
        }
    }


    # ----------------------------------------------------------------------
    Context 'Out-LogFileHC - CSV / JSON / TXT / XLSX' {

        BeforeEach {
            Mock Export-Excel
        }

        It 'Exports CSV successfully' {
            $path = Join-Path $TestDrive 'export'
            $rows = @([pscustomobject]@{A = 1 })

            $result = Out-LogFileHC `
                -DataToExport $rows `
                -PartialPath $path `
                -FileExtensions '.csv'

            Test-Path ($result[0]) | Should -BeTrue
        }

        It 'Exports JSON successfully' {
            $path = Join-Path $TestDrive 'log'
            $rows = @([pscustomobject]@{A = 1 })

            $result = Out-LogFileHC `
                -DataToExport $rows `
                -PartialPath $path `
                -FileExtensions '.json'

            Test-Path ($result[0]) | Should -BeTrue
            (Get-Content $result[0] -Raw) | Should -Match '"A":'
        }

        It 'Exports TXT successfully' {
            $path = Join-Path $TestDrive 'txtlog'
            $rows = @([pscustomobject]@{A = 1; B = 2 })

            $result = Out-LogFileHC `
                -DataToExport $rows `
                -PartialPath $path `
                -FileExtensions '.txt'

            Test-Path ($result[0]) | Should -BeTrue
        }

        It 'Exports XLSX via Export-Excel' {
            $path = Join-Path $TestDrive 'excel'
            $rows = @([pscustomobject]@{A = 1 })

            Out-LogFileHC -DataToExport $rows -PartialPath $path -FileExtensions '.xlsx'

            Should -Invoke Export-Excel -Times 1
        }
    }


    # ----------------------------------------------------------------------
    Context 'Remove-FileHC' {

        It 'Deletes file when present' {
            $path = Join-Path $TestDrive 'del.txt'
            Set-Content $path 'test'

            Remove-FileHC -FilePath $path

            Test-Path $path | Should -BeFalse
        }

        It 'Adds error when deletion fails' {
            $sys = @()

            Mock Remove-Item { throw 'Mock failure' }

            $path = Join-Path $TestDrive 'err.txt'
            Set-Content $path 'test'

            Remove-FileHC -FilePath $path -SystemErrors ([ref]$sys)

            $sys.Count | Should -Be 1
            $sys[0].Category | Should -Be 'Logging'
        }
    }


    # ----------------------------------------------------------------------
    Context 'Write-EventLogSafe' {

        It 'Skips event log write when disabled' {
            Mock Write-EventsToEventLogHC

            $settings = @{
                SaveInEventLog = @{
                    Save    = $false
                    LogName = 'TestLog'
                }
            }

            $sysErrors = @()
            $data = [System.Collections.Generic.List[object]]::new()

            Write-EventLogSafe `
                -EventLogData $data `
                -ScriptName 'X' `
                -Settings $settings `
                -SystemErrors ([ref]$sysErrors)

            Should -Not -Invoke Write-EventsToEventLogHC
        }

        It 'Invokes Write-EventsToEventLogHC when enabled' {
            Mock Write-EventsToEventLogHC

            $settings = @{
                SaveInEventLog = @{
                    Save    = $true
                    LogName = 'TestLog'
                }
            }

            $sysErrors = @(
                [pscustomobject]@{ Message = 'E1'; DateTime = Get-Date }
            )
            $data = [System.Collections.Generic.List[object]]::new()

            Write-EventLogSafe `
                -EventLogData $data `
                -ScriptName 'TestScript' `
                -Settings $settings `
                -SystemErrors ([ref]$sysErrors)

            Should -Invoke Write-EventsToEventLogHC -Times 1
        }
    }


    # ----------------------------------------------------------------------
    Context 'Write-SystemErrorLogHC' {

        It 'Creates JSON system error log and adds attachment' {

            $sys = @(
                [pscustomobject]@{ DateTime = Get-Date; Message = 'Err1' }
            )
            $folder = Join-Path $TestDrive 'syslogs'
            New-Item -ItemType Directory -Path $folder | Out-Null

            $mailParams = @{ } | Select-Object -Property * -ExpandProperty * -ErrorAction Ignore
            $mailParams = [ref]@{}

            Write-SystemErrorLogHC `
                -SystemErrors $sys `
                -LogFolder $folder `
                -MailParams $mailParams `
                -ScriptStartTime (Get-Date) `
                -JsonFileItem @{ BaseName = 'Cfg' }

            $mailParams.Value.Attachments.Count | Should -Be 1
            Test-Path $mailParams.Value.Attachments[0] | Should -BeTrue
        }
    }
}