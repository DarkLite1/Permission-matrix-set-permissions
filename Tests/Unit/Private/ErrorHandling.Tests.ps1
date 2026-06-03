#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
    ForEach-Object { . $_.FullName }
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
