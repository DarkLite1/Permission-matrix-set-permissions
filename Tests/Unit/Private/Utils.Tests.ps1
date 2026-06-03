#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    . "$moduleRoot\Private\Utils.ps1"
}

Describe 'Add-ErrorHC' {
    BeforeEach {
        $script:errors = [System.Collections.Generic.List[PSObject]]::new()
    }

    It 'adds a single object to the SystemErrors collection' {
        Add-ErrorHC -Type 'FatalError' -Name 'Test' -Message 'Boom' `
            -Category 'Matrix' -SystemErrors ([ref]$script:errors)

        $script:errors.Count | Should -Be 1
    }

    It 'populates all provided fields' {
        Add-ErrorHC `
            -Type 'Warning' `
            -Name 'N1' `
            -Message 'M1' `
            -Description 'D1' `
            -Category 'Permissions' `
            -SystemErrors ([ref]$script:errors)

        $e = $script:errors[0]
        $e.Type | Should -Be 'Warning'
        $e.Name | Should -Be 'N1'
        $e.Message | Should -Be 'M1'
        $e.Description | Should -Be 'D1'
        $e.Category | Should -Be 'Permissions'
        $e.DateTime | Should -BeOfType [datetime]
    }

    It 'defaults Description to an empty string when omitted' {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Name 'N' `
            -Message 'M' `
            -Category 'JsonSchema' `
            -SystemErrors ([ref]$script:errors)

        $script:errors[0].Description | Should -Be ''
    }

    It 'appends without clearing existing entries' {
        Add-ErrorHC `
            -Type 'Warning' `
            -Name 'A' `
            -Message 'M' `
            -Category 'C' `
            -SystemErrors ([ref]$script:errors)
        Add-ErrorHC `
            -Type 'Warning' `
            -Name 'B' `
            -Message 'M' `
            -Category 'C' `
            -SystemErrors ([ref]$script:errors)

        $script:errors.Count | Should -Be 2
    }

    It 'throws when a mandatory parameter is missing' {
        # Splat with $null so binding fails as a terminating error instead of
        # prompting interactively for the missing mandatory value.
        $params = @{
            Name         = 'N'
            Message      = 'M'
            Category     = 'C'
            SystemErrors = ([ref]$script:errors)
            Type         = $null
        }
        { Add-ErrorHC @params } | Should -Throw
    }
}

Describe 'Category wrapper functions' {
    BeforeEach {
        $script:errors = [System.Collections.Generic.List[PSObject]]::new()
    }

    It 'Add-MatrixErrorHC sets Category to Matrix' {
        Add-MatrixErrorHC `
            -Type 'FatalError' `
            -Name 'N' `
            -Message 'M' `
            -SystemErrors ([ref]$script:errors)
        $script:errors[0].Category | Should -Be 'Matrix'
    }

    It 'Add-PermissionsErrorHC sets Category to Permissions' {
        Add-PermissionsErrorHC `
            -Type 'FatalError' `
            -Name 'N' `
            -Message 'M' `
            -SystemErrors ([ref]$script:errors)
        $script:errors[0].Category | Should -Be 'Permissions'
    }

    It 'Add-RuntimeErrorHC sets Category to RuntimeSettings' {
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name 'N' `
            -Message 'M' `
            -SystemErrors ([ref]$script:errors)
        $script:errors[0].Category | Should -Be 'RuntimeSettings'
    }

    It 'Add-JsonSchemaErrorHC sets Category to JsonSchema' {
        Add-JsonSchemaErrorHC `
            -Type 'FatalError' `
            -Name 'N' `
            -Message 'M' `
            -SystemErrors ([ref]$script:errors)
        $script:errors[0].Category | Should -Be 'JsonSchema'
    }

    It 'forwards Description through the wrapper' {
        Add-MatrixErrorHC `
            -Type 'Warning' `
            -Name 'N' `
            -Message 'M' `
            -Description 'forwarded' `
            -SystemErrors ([ref]$script:errors)
        $script:errors[0].Description | Should -Be 'forwarded'
    }
}

Describe 'Get-StringValueHC' {
    It 'returns null for null input' {
        Get-StringValueHC -Name $null | Should -BeNullOrEmpty
    }

    It 'returns null for empty string' {
        Get-StringValueHC -Name '' | Should -BeNullOrEmpty
    }

    It 'returns null for whitespace-only input' {
        Get-StringValueHC -Name '   ' | Should -BeNullOrEmpty
    }

    It 'returns the literal value when no ENV: prefix' {
        Get-StringValueHC -Name 'PlainValue' | Should -Be 'PlainValue'
    }

    Context 'ENV: prefix resolution' {
        BeforeAll {
            $env:PESTER_UTILS_VAR = 'resolved-value'
        }
        AfterAll {
            Remove-Item -Path 'Env:\PESTER_UTILS_VAR' -ErrorAction Ignore
        }

        It 'resolves an existing environment variable' {
            Get-StringValueHC -Name 'ENV:PESTER_UTILS_VAR' | Should -Be 'resolved-value'
        }

        It 'is case-insensitive on the ENV: prefix' {
            Get-StringValueHC -Name 'env:PESTER_UTILS_VAR' | Should -Be 'resolved-value'
        }

        It 'trims whitespace around the variable name' {
            Get-StringValueHC -Name 'ENV:  PESTER_UTILS_VAR  ' | Should -Be 'resolved-value'
        }

        It 'throws when the environment variable does not exist' {
            { Get-StringValueHC -Name 'ENV:DOES_NOT_EXIST_XYZ' } |
            Should -Throw "*'DOES_NOT_EXIST_XYZ' not found*"
        }
    }
}

Describe 'Get-StringOrDefaultHC' {
    It 'returns Default for null' {
        Get-StringOrDefaultHC -Value $null -Default 'fallback' | Should -Be 'fallback'
    }

    It 'returns Default for empty string' {
        Get-StringOrDefaultHC -Value '' -Default 'fallback' | Should -Be 'fallback'
    }

    It 'returns Default for whitespace only' {
        Get-StringOrDefaultHC -Value '   ' -Default 'fallback' | Should -Be 'fallback'
    }

    It 'returns the value when non-blank' {
        Get-StringOrDefaultHC -Value 'real' -Default 'fallback' | Should -Be 'real'
    }

    It 'passes through 0 as non-blank' {
        Get-StringOrDefaultHC -Value 0 -Default 'fallback' | Should -Be 0
    }

    It 'allows an empty string as the Default' {
        Get-StringOrDefaultHC -Value $null -Default '' | Should -Be ''
    }

    It 'accepts both arguments positionally (Value then Default)' {
        Get-StringOrDefaultHC 'real' 'fallback' | Should -Be 'real'
    }

    It 'returns the positional Default when the positional Value is blank' {
        Get-StringOrDefaultHC '' 'fallback' | Should -Be 'fallback'
    }

    It 'defaults to an empty string when Default is omitted' {
        Get-StringOrDefaultHC -Value $null | Should -Be ''
    }
}

Describe 'Get-DatedLogFolderPathHC' {
    BeforeAll {
        $script:startTime = [datetime]'2024-03-07 09:05:08'
    }

    It 'creates a dated folder and returns its full path' {
        $result = Get-DatedLogFolderPathHC `
            -LogFolder 'TestDrive:\Logs' `
            -ScriptStartTime $script:startTime `
            -JsonFileName 'MyScript'

        $result | Should -Not -BeNullOrEmpty
        Test-Path -Path $result | Should -BeTrue
    }

    It 'formats the folder name as yyyy_MM_dd_HHmmss (JsonFileName)' {
        $result = Get-DatedLogFolderPathHC `
            -LogFolder 'TestDrive:\Logs' `
            -ScriptStartTime $script:startTime `
            -JsonFileName 'MyScript'

        Split-Path -Path $result -Leaf | Should -Be '2024_03_07_090508 (MyScript)'
    }

    It 'returns the original LogFolder when creation fails' {
        Mock New-Item { throw 'denied' }

        $result = Get-DatedLogFolderPathHC `
            -LogFolder 'C:\Original' `
            -ScriptStartTime $script:startTime `
            -JsonFileName 'X'

        $result | Should -Be 'C:\Original'
    }
}

Describe 'Plural' {
    It 'returns the singular word when count is 1' {
        Plural -Count 1 -Word 'error' | Should -Be 'error'
    }

    It 'pluralizes when count is 0' {
        Plural -Count 0 -Word 'error' | Should -Be 'errors'
    }

    It 'pluralizes when count is greater than 1' {
        Plural -Count 5 -Word 'warning' | Should -Be 'warnings'
    }
}

Describe 'Remove-BlankValueHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Utils.ps1"
    }

    It 'removes a key whose value is an empty string' {
        $result = Remove-BlankValueHC -Hashtable @{ Keep = 'x'; Drop = '' }

        $result.ContainsKey('Drop') | Should -BeFalse
        $result.Keep | Should -Be 'x'
    }

    It 'removes a key whose value is whitespace only' {
        $result = Remove-BlankValueHC -Hashtable @{ Drop = '   ' }

        $result.ContainsKey('Drop') | Should -BeFalse
    }

    It 'removes a key whose value is $null' {
        $result = Remove-BlankValueHC -Hashtable @{ Drop = $null }

        $result.ContainsKey('Drop') | Should -BeFalse
    }

    It 'keeps non-string values: 0, $false and an empty array' {
        $result = Remove-BlankValueHC -Hashtable @{ Zero = 0; Flag = $false; Empty = @() }

        $result.ContainsKey('Zero') | Should -BeTrue
        $result.ContainsKey('Flag') | Should -BeTrue
        $result.ContainsKey('Empty') | Should -BeTrue
    }

    It 'keeps populated arrays' {
        $result = Remove-BlankValueHC -Hashtable @{ To = @('a@example.com', 'b@example.com') }

        $result.To | Should -HaveCount 2
    }

    It 'does not modify the original hashtable' {
        $original = @{ Keep = 'x'; Drop = '' }

        $null = Remove-BlankValueHC -Hashtable $original

        $original.ContainsKey('Drop') | Should -BeTrue
    }

    It 'mirrors the failing mail case: a blank SmtpConnectionType is dropped so the default applies' {
        $mailParams = @{
            To                 = @('test@example.com')
            From               = 'noreply@example.com'
            SmtpConnectionType = $null   # unset config -> Get-StringValueHC returns $null
            Attachments        = @()
        }

        $result = Remove-BlankValueHC -Hashtable $mailParams

        $result.ContainsKey('SmtpConnectionType') | Should -BeFalse
        $result.ContainsKey('To') | Should -BeTrue
        $result.ContainsKey('Attachments') | Should -BeTrue
    }
}

Describe 'Test-ItemHasFatalErrorHC' {
    It 'returns false for null CheckList' {
        Test-ItemHasFatalErrorHC -CheckList $null | Should -BeFalse
    }

    It 'returns false for empty CheckList' {
        Test-ItemHasFatalErrorHC -CheckList @() | Should -BeFalse
    }

    It 'returns false when no FatalError is present' {
        $list = @(
            [PSCustomObject]@{ Type = 'Warning' }
            [PSCustomObject]@{ Type = 'Info' }
        )
        Test-ItemHasFatalErrorHC -CheckList $list | Should -BeFalse
    }

    It 'returns true when a FatalError is present' {
        $list = @(
            [PSCustomObject]@{ Type = 'Warning' }
            [PSCustomObject]@{ Type = 'FatalError' }
        )
        Test-ItemHasFatalErrorHC -CheckList $list | Should -BeTrue
    }
}

Describe 'New-CounterObjectHC' {
    BeforeAll {
        $script:counter = New-CounterObjectHC
    }

    It 'initializes top-level totals to 0' {
        $script:counter.TotalErrors | Should -Be 0
        $script:counter.TotalWarnings | Should -Be 0
    }

    It 'creates all four buckets with zeroed Errors and Warnings' {
        foreach ($bucket in 'FormData', 'Permissions', 'Settings', 'File') {
            $script:counter.$bucket.Errors | Should -Be 0
            $script:counter.$bucket.Warnings | Should -Be 0
        }
    }
}

Describe 'Update-MatrixCounterHC' {
    BeforeEach {
        $script:errors = [System.Collections.Generic.List[PSObject]]::new()
    }

    It 'returns an all-zero counter for empty input' {
        $context = [PSCustomObject]@{ FileResults = @(); Counter = $null }

        $result = Update-MatrixCounterHC `
            -Context $context `
            -SystemErrors ([ref]$script:errors)

        $result.TotalErrors | Should -Be 0
        $result.TotalWarnings | Should -Be 0
    }

    It 'counts file-, sheet-, and matrix-level checks into the right buckets' {
        $context = [PSCustomObject]@{
            Counter     = $null
            FileResults = @(
                [PSCustomObject]@{
                    Check    = @([PSCustomObject]@{ Type = 'FatalError' })
                    Sheets   = [PSCustomObject]@{
                        FormData    = [PSCustomObject]@{ Check = @([PSCustomObject]@{ Type = 'Warning' }) }
                        Permissions = [PSCustomObject]@{ Check = @([PSCustomObject]@{ Type = 'FatalError' }) }
                    }
                    Matrices = @(
                        [PSCustomObject]@{ Check = @(
                                [PSCustomObject]@{ Type = 'Warning' }
                                [PSCustomObject]@{ Type = 'Warning' }
                            ) 
                        }
                    )
                }
            )
        }

        $result = Update-MatrixCounterHC -Context $context -SystemErrors ([ref]$script:errors)

        $result.File.Errors | Should -Be 1
        $result.FormData.Warnings | Should -Be 1
        $result.Permissions.Errors | Should -Be 1
        $result.Settings.Warnings | Should -Be 2
    }

    It 'includes system-level errors and warnings in totals' {
        $script:errors.Add([PSCustomObject]@{ Type = 'FatalError' })
        $script:errors.Add([PSCustomObject]@{ Type = 'Warning' })
        $script:errors.Add([PSCustomObject]@{ Type = 'Warning' })

        $context = [PSCustomObject]@{ FileResults = @(); Counter = $null }

        $result = Update-MatrixCounterHC -Context $context -SystemErrors ([ref]$script:errors)

        $result.TotalErrors | Should -Be 1
        $result.TotalWarnings | Should -Be 2
    }

    It 'sums every bucket plus system errors into the grand totals' {
        $script:errors.Add([PSCustomObject]@{ Type = 'FatalError' })

        $context = [PSCustomObject]@{
            Counter     = $null
            FileResults = @(
                [PSCustomObject]@{
                    Check    = @(
                        [PSCustomObject]@{ Type = 'FatalError' }
                    )
                    Sheets   = [PSCustomObject]@{
                        FormData    = [PSCustomObject]@{ 
                            Check = @(
                                [PSCustomObject]@{ 
                                    Type = 'Warning' 
                                }
                            ) 
                        }
                        Permissions = [PSCustomObject]@{
                            Check = @() 
                        }
                    }
                    Matrices = @(
                        [PSCustomObject]@{ 
                            Check = @(
                                [PSCustomObject]@{ 
                                    Type = 'FatalError' 
                                }
                            ) 
                        }
                    )
                }
            )
        }

        $result = Update-MatrixCounterHC `
            -Context $context `
            -SystemErrors ([ref]$script:errors)

        $result.TotalErrors | Should -Be 3
        $result.TotalWarnings | Should -Be 1
    }

    It 'assigns the counter back onto the Context' {
        $context = [PSCustomObject]@{ 
            FileResults = @()
            Counter     = $null
        }

        Update-MatrixCounterHC `
            -Context $context `
            -SystemErrors ([ref]$script:errors) | Out-Null

        $context.Counter | Should -Not -BeNullOrEmpty
    }
}