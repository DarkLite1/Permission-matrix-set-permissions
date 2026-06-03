#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
#Requires -Modules ImportExcel

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
    ForEach-Object { . $_.FullName }
}

Describe 'Remove-OldLogsHC' {
    BeforeEach {
        $script:errors = [System.Collections.Generic.List[PSObject]]::new()
    }

    It 'returns silently when RetentionDays is 0 or less' {
        { Remove-OldLogsHC -LogFolder 'TestDrive:\Logs' -RetentionDays 0 `
                -SystemErrors ([ref]$script:errors) } | Should -Not -Throw
    }

    It 'returns silently when the folder does not exist' {
        { Remove-OldLogsHC -LogFolder 'TestDrive:\DoesNotExist' -RetentionDays 30 `
                -SystemErrors ([ref]$script:errors) } | Should -Not -Throw
        $script:errors.Count | Should -Be 0
    }

    It 'deletes files older than the retention window' {
        $folder = Join-Path $TestDrive 'Logs1'
        New-Item -Path $folder -ItemType Directory -Force | Out-Null

        $old = Join-Path $folder 'old.log'
        $new = Join-Path $folder 'new.log'
        Set-Content -LiteralPath $old -Value 'old'
        Set-Content -LiteralPath $new -Value 'new'
        (Get-Item $old).CreationTime = (Get-Date).AddDays(-40)

        Remove-OldLogsHC -LogFolder $folder -RetentionDays 30 `
            -SystemErrors ([ref]$script:errors)

        Test-Path $old | Should -BeFalse
        Test-Path $new | Should -BeTrue
    }

    It 'keeps files newer than the retention window' {
        $folder = Join-Path $TestDrive 'Logs2'
        New-Item -Path $folder -ItemType Directory -Force | Out-Null

        $recent = Join-Path $folder 'recent.log'
        Set-Content -LiteralPath $recent -Value 'x'
        (Get-Item $recent).CreationTime = (Get-Date).AddDays(-5)

        Remove-OldLogsHC -LogFolder $folder -RetentionDays 30 `
            -SystemErrors ([ref]$script:errors)

        Test-Path $recent | Should -BeTrue
    }

    It 'removes empty sub-folders left behind after file deletion' {
        $folder = Join-Path $TestDrive 'Logs3'
        $sub = Join-Path $folder 'sub'
        New-Item -Path $sub -ItemType Directory -Force | Out-Null

        $old = Join-Path $sub 'old.log'
        Set-Content -LiteralPath $old -Value 'x'
        (Get-Item $old).CreationTime = (Get-Date).AddDays(-90)

        Remove-OldLogsHC -LogFolder $folder -RetentionDays 30 `
            -SystemErrors ([ref]$script:errors)

        Test-Path $old | Should -BeFalse
        Test-Path $sub | Should -BeFalse
    }

    It 'leaves non-empty sub-folders intact' {
        $folder = Join-Path $TestDrive 'Logs4'
        $sub = Join-Path $folder 'sub'
        New-Item -Path $sub -ItemType Directory -Force | Out-Null

        $keep = Join-Path $sub 'keep.log'
        Set-Content -LiteralPath $keep -Value 'x'
        (Get-Item $keep).CreationTime = (Get-Date).AddDays(-1)

        Remove-OldLogsHC -LogFolder $folder -RetentionDays 30 `
            -SystemErrors ([ref]$script:errors)

        Test-Path $sub | Should -BeTrue
    }
}

Describe 'Remove-FileHC' {
    BeforeEach {
        $script:errors = [System.Collections.Generic.List[PSObject]]::new()
    }

    It 'removes an existing file' {
        $file = Join-Path $TestDrive 'remove-me.txt'
        Set-Content -LiteralPath $file -Value 'x'

        Remove-FileHC -FilePath $file -SystemErrors ([ref]$script:errors)

        Test-Path $file | Should -BeFalse
    }

    It 'returns silently when the file does not exist' {
        { Remove-FileHC -FilePath (Join-Path $TestDrive 'nope.txt') `
                -SystemErrors ([ref]$script:errors) } | Should -Not -Throw
        $script:errors.Count | Should -Be 0
    }

    It 'returns silently when given a directory path (PathType Leaf)' {
        $dir = Join-Path $TestDrive 'adir'
        New-Item -Path $dir -ItemType Directory -Force | Out-Null

        { Remove-FileHC -FilePath $dir -SystemErrors ([ref]$script:errors) } |
        Should -Not -Throw
        Test-Path $dir | Should -BeTrue
    }

    It 'records a warning when deletion fails and SystemErrors is supplied' {
        $file = Join-Path $TestDrive 'locked.txt'
        Set-Content -LiteralPath $file -Value 'x'

        Mock Remove-Item { throw 'access denied' } -ParameterFilter {
            $LiteralPath -eq $file
        }

        Remove-FileHC -FilePath $file -SystemErrors ([ref]$script:errors)

        $script:errors.Count | Should -Be 1
        $script:errors[0].Name | Should -Be 'Failed to remove file'
    }

    It 'falls back to Write-Warning when no SystemErrors ref is given' {
        $file = Join-Path $TestDrive 'locked2.txt'
        Set-Content -LiteralPath $file -Value 'x'

        Mock Remove-Item { throw 'access denied' } -ParameterFilter {
            $LiteralPath -eq $file
        }

        { Remove-FileHC -FilePath $file -WarningAction SilentlyContinue } |
        Should -Not -Throw
    }
}

Describe 'Out-LogFileHC' {
    BeforeAll {
        $script:data = @(
            [PSCustomObject]@{ Name = 'Alice'; Score = 1 }
            [PSCustomObject]@{ Name = 'Bob'; Score = 2 }
        )
    }

    It 'creates a semicolon-delimited CSV' {
        $partial = Join-Path $TestDrive 'out-csv'

        $paths = Out-LogFileHC -DataToExport $script:data `
            -PartialPath $partial -FileExtensions '.csv'

        $paths | Should -Be "$partial.csv"
        Test-Path "$partial.csv" | Should -BeTrue
        (Get-Content "$partial.csv" -Raw) | Should -Match ';'
    }

    It 'creates a JSON file with all rows' {
        $partial = Join-Path $TestDrive 'out-json'

        Out-LogFileHC -DataToExport $script:data `
            -PartialPath $partial -FileExtensions '.json' | Out-Null

        $obj = Get-Content "$partial.json" -Raw | ConvertFrom-Json
        $obj.Count | Should -Be 2
        $obj[0].Name | Should -Be 'Alice'
    }

    It 'creates a TXT file' {
        $partial = Join-Path $TestDrive 'out-txt'

        Out-LogFileHC -DataToExport $script:data `
            -PartialPath $partial -FileExtensions '.txt' | Out-Null

        Test-Path "$partial.txt" | Should -BeTrue
        (Get-Content "$partial.txt" -Raw) | Should -Match 'Alice'
    }

    It 'creates an XLSX file' {
        $partial = Join-Path $TestDrive 'out-xlsx'

        $paths = Out-LogFileHC -DataToExport $script:data `
            -PartialPath $partial -FileExtensions '.xlsx'

        $paths | Should -Be "$partial.xlsx"
        Test-Path "$partial.xlsx" | Should -BeTrue
    }

    It 'creates multiple files when several extensions are given' {
        $partial = Join-Path $TestDrive 'out-multi'

        $paths = Out-LogFileHC -DataToExport $script:data `
            -PartialPath $partial -FileExtensions '.csv', '.json'

        $paths.Count | Should -Be 2
        Test-Path "$partial.csv" | Should -BeTrue
        Test-Path "$partial.json" | Should -BeTrue
    }

    It 'de-duplicates repeated extensions' {
        $partial = Join-Path $TestDrive 'out-dupe'

        $paths = Out-LogFileHC -DataToExport $script:data `
            -PartialPath $partial -FileExtensions '.csv', '.csv'

        @($paths).Count | Should -Be 1
    }

    It 'warns and skips an unsupported extension instead of throwing' {
        $partial = Join-Path $TestDrive 'out-bad'

        $paths = Out-LogFileHC -DataToExport $script:data `
            -PartialPath $partial -FileExtensions '.xyz' `
            -WarningAction SilentlyContinue

        @($paths).Count | Should -Be 0
    }

    It 'appends rows to an existing JSON file' {
        $partial = Join-Path $TestDrive 'out-append'

        Out-LogFileHC -DataToExport $script:data `
            -PartialPath $partial -FileExtensions '.json' | Out-Null
        Out-LogFileHC -DataToExport @([PSCustomObject]@{ Name = 'Carol'; Score = 3 }) `
            -PartialPath $partial -FileExtensions '.json' -Append | Out-Null

        $obj = Get-Content "$partial.json" -Raw | ConvertFrom-Json
        $obj.Count | Should -Be 3
        ($obj.Name) | Should -Contain 'Carol'
    }

    It 'converts ErrorRecord properties to their message text in JSON' {
        $partial = Join-Path $TestDrive 'out-errrec'
        $rec = try { throw 'boom' } catch { $_ }
        $row = [PSCustomObject]@{ Name = 'X'; Error = $rec }

        Out-LogFileHC -DataToExport @($row) `
            -PartialPath $partial -FileExtensions '.json' | Out-Null

        $obj = Get-Content "$partial.json" -Raw | ConvertFrom-Json
        $obj.Error | Should -Be 'boom'
    }
}

Describe 'Write-EventsToEventLogHC' {
    BeforeAll {
        Mock New-EventLog {}
        Mock Write-EventLog {}
    }

    It 'does not throw for a basic event when source handling succeeds' {
        Mock New-EventLog {}
        Mock Write-EventLog {}

        $events = @(
            [PSCustomObject]@{ EntryType = 'Information'; EventID = 4; Message = 'hi' }
        )

        { Write-EventsToEventLogHC -Source 'PesterFakeSource' `
                -LogName 'Application' -Events $events } | Should -Not -Throw
    }

    It 'writes one entry per event' {
        Mock Write-EventLog {}
        Mock New-EventLog {}

        $events = @(
            [PSCustomObject]@{ EntryType = 'Information'; EventID = 4; Message = 'a' }
            [PSCustomObject]@{ EntryType = 'Error'; EventID = 2; Message = 'b' }
        )

        Write-EventsToEventLogHC -Source 'PesterFakeSource' `
            -LogName 'Application' -Events $events

        Should -Invoke Write-EventLog -Times 2 -Exactly
    }

    It 'defaults EntryType to Information and EventID to 4 when missing' {
        Mock New-EventLog {}
        Mock Write-EventLog {} -Verifiable -ParameterFilter {
            $EntryType -eq 'Information' -and $EventID -eq 4
        }

        $events = @([PSCustomObject]@{ Message = 'no type' })

        Write-EventsToEventLogHC -Source 'PesterFakeSource' `
            -LogName 'Application' -Events $events

        Should -InvokeVerifiable
    }

    It 're-throws a wrapped error when Write-EventLog fails' {
        Mock New-EventLog {}
        Mock Write-EventLog { throw 'nope' }

        $events = @([PSCustomObject]@{ EntryType = 'Information'; EventID = 4; Message = 'x' })

        { Write-EventsToEventLogHC -Source 'PesterFakeSource' `
                -LogName 'Application' -Events $events } |
        Should -Throw '*Failed writing events*'
    }
}

Describe 'Write-EventLogSafeHC' {
    BeforeEach {
        $script:errors = [System.Collections.Generic.List[PSObject]]::new()
        $script:eventData = [System.Collections.Generic.List[PSObject]]::new()
    }

    It 'returns without writing when Save is false' {
        Mock Write-EventsToEventLogHC {}

        $settings = [PSCustomObject]@{
            SaveInEventLog = [PSCustomObject]@{ Save = $false; LogName = 'App' }
        }

        Write-EventLogSafeHC -EventLogData $script:eventData -ScriptName 'S' `
            -Settings $settings -SystemErrors ([ref]$script:errors)

        Should -Invoke Write-EventsToEventLogHC -Times 0 -Exactly
    }

    It 'returns without writing when LogName is blank' {
        Mock Write-EventsToEventLogHC {}

        $settings = [PSCustomObject]@{
            SaveInEventLog = [PSCustomObject]@{ Save = $true; LogName = '' }
        }

        Write-EventLogSafeHC -EventLogData $script:eventData -ScriptName 'S' `
            -Settings $settings -SystemErrors ([ref]$script:errors)

        Should -Invoke Write-EventsToEventLogHC -Times 0 -Exactly
    }

    It 'appends a System error entry plus a script-ended entry, then writes' {
        Mock Write-EventsToEventLogHC {}

        $script:errors.Add([PSCustomObject]@{ Message = 'bad'; DateTime = (Get-Date) })

        $settings = [PSCustomObject]@{
            SaveInEventLog = [PSCustomObject]@{ Save = $true; LogName = 'Application' }
        }

        Write-EventLogSafeHC -EventLogData $script:eventData -ScriptName 'S' `
            -Settings $settings -SystemErrors ([ref]$script:errors)

        # 1 error entry + 1 "Script ended" entry
        $script:eventData.Count | Should -Be 2
        $script:eventData[0].EntryType | Should -Be 'Error'
        $script:eventData[1].Message | Should -Be 'Script ended'
        Should -Invoke Write-EventsToEventLogHC -Times 1 -Exactly
    }

    It 'truncates messages longer than the 31000 char limit' {
        Mock Write-EventsToEventLogHC {}

        $script:eventData.Add([PSCustomObject]@{
                Message   = ('x' * 32000)
                DateTime  = (Get-Date)
                EntryType = 'Information'
                EventID   = '1'
            })

        $settings = [PSCustomObject]@{
            SaveInEventLog = [PSCustomObject]@{ Save = $true; LogName = 'Application' }
        }

        Write-EventLogSafeHC -EventLogData $script:eventData -ScriptName 'S' `
            -Settings $settings -SystemErrors ([ref]$script:errors)

        $longEntry = $script:eventData | Where-Object { $_.EventID -eq '1' }
        $longEntry.Message | Should -Match 'TRUNCATED'
        $longEntry.Message.Length | Should -BeLessThan 31100
    }

    It 'records a warning when the underlying write throws' {
        Mock Write-EventsToEventLogHC { throw 'boom' }

        $settings = [PSCustomObject]@{
            SaveInEventLog = [PSCustomObject]@{ Save = $true; LogName = 'Application' }
        }

        Write-EventLogSafeHC -EventLogData $script:eventData -ScriptName 'S' `
            -Settings $settings -SystemErrors ([ref]$script:errors)

        $script:errors.Count | Should -Be 1
        $script:errors[0].Name | Should -Be 'Failed to write to event log'
    }
}

Describe 'Write-SystemErrorLogHC' {
    BeforeEach {
        $script:mailParams = @{}
    }

    It 'returns silently when there are no system errors' {
        Mock Out-LogFileHC {}

        $errs = [System.Collections.Generic.List[PSObject]]::new()

        Write-SystemErrorLogHC -SystemErrors $errs -LogFolder $TestDrive `
            -MailParams ([ref]$script:mailParams)

        Should -Invoke Out-LogFileHC -Times 0 -Exactly
        $script:mailParams.ContainsKey('Attachments') | Should -BeFalse
    }

    It 'returns silently when the log folder is missing' {
        Mock Out-LogFileHC {}

        $errs = [System.Collections.Generic.List[PSObject]]::new()
        $errs.Add([PSCustomObject]@{ Message = 'x' })

        Write-SystemErrorLogHC -SystemErrors $errs `
            -LogFolder (Join-Path $TestDrive 'missing') `
            -MailParams ([ref]$script:mailParams)

        Should -Invoke Out-LogFileHC -Times 0 -Exactly
    }

    It 'writes a JSON log and attaches it to the mail params' {
        Mock Out-LogFileHC { @("$TestDrive\SystemErrors.json") }

        $errs = [System.Collections.Generic.List[PSObject]]::new()
        $errs.Add([PSCustomObject]@{ Message = 'x' })

        Write-SystemErrorLogHC -SystemErrors $errs -LogFolder $TestDrive `
            -MailParams ([ref]$script:mailParams)

        $script:mailParams['Attachments'] | Should -Contain "$TestDrive\SystemErrors.json"
    }

    It 'appends to existing attachments rather than overwriting' {
        Mock Out-LogFileHC { @("$TestDrive\SystemErrors.json") }

        $script:mailParams['Attachments'] = @('existing.txt')

        $errs = [System.Collections.Generic.List[PSObject]]::new()
        $errs.Add([PSCustomObject]@{ Message = 'x' })

        Write-SystemErrorLogHC -SystemErrors $errs -LogFolder $TestDrive `
            -MailParams ([ref]$script:mailParams)

        $script:mailParams['Attachments'].Count | Should -Be 2
        $script:mailParams['Attachments'] | Should -Contain 'existing.txt'
    }

    It 'does not add Attachments when nothing was written' {
        Mock Out-LogFileHC { $null }

        $errs = [System.Collections.Generic.List[PSObject]]::new()
        $errs.Add([PSCustomObject]@{ Message = 'x' })

        Write-SystemErrorLogHC -SystemErrors $errs -LogFolder $TestDrive `
            -MailParams ([ref]$script:mailParams)

        $script:mailParams.ContainsKey('Attachments') | Should -BeFalse
    }
}