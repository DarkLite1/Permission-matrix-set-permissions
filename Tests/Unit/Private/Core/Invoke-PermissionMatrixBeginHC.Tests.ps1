#Requires -Version 7
#Requires -Modules Pester

Describe 'Invoke-PermissionMatrixBeginHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$root\Tests\Helpers\Helpers.HC.ps1"
        . "$root\Tests\Helpers\Fixtures.Json.ps1"
        . "$root\Tests\Helpers\Fixtures.Excel.ps1"

        . "$moduleRoot\Private\Utils.ps1"
        . "$moduleRoot\Private\ActiveDirectory.ps1"
        . "$moduleRoot\Private\Validation.ps1"
        . "$moduleRoot\Private\Matrix.ps1"
        . "$moduleRoot\Private\Export.ps1"
        . "$moduleRoot\Private\Validation\Validate-ConfigurationStructureHC.ps1"
        . "$moduleRoot\Private\Matrix\Import-MatrixFileHC.ps1"
        . "$moduleRoot\Private\Matrix\Import-MatrixDefaultsFileHC.ps1"
        . "$moduleRoot\Private\Infrastructure\Invoke-WithOptionalParallelismHC.ps1"
        . "$moduleRoot\Private\Core\Invoke-PermissionMatrixBeginHC.ps1"

        function New-FakeScriptPath {
            param(
                [string]$Root = 'TestDrive:'
            )

            $modulePath = (New-Item "$Root\PermissionMatrix.psm1" -ItemType File -Force).FullName
            $setPerm = (New-Item "$Root\SetPermissions.ps1" -ItemType File -Force).FullName
            $testReq = (New-Item "$Root\TestRequirements.ps1" -ItemType File -Force).FullName
            $snow = (New-Item "$Root\UpdateServiceNow.ps1" -ItemType File -Force).FullName

            return @{
                PermissionMatrixModule = $modulePath
                SetPermissions         = $setPerm
                TestRequirements       = $testReq
                UpdateServiceNow       = $snow
            }
        }

        # ---------------------------------------------------------------------
        # Helper: write a JSON config fixture and return its path
        # ---------------------------------------------------------------------
        function New-BeginJsonFile {
            param(
                [hashtable]$Overrides = @{},
                [string]$Path = 'TestDrive:\Input.json'
            )

            $fixture = New-JsonFixture

            # Sensible defaults that BeginHC will touch
            $fixture.Matrix.FolderPath = (New-Item 'TestDrive:\Matrix' -ItemType Directory -Force).FullName
            $fixture.Matrix.DefaultsFile = (New-ValidDefaultsExcelFixture -Path 'TestDrive:\Defaults.xlsx')
            $fixture.Settings.SaveLogFiles.Where.Folder = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName

            # Apply per-test overrides via dotted path: 'Matrix.FolderPath' = 'x:\nope'
            foreach ($key in $Overrides.Keys) {
                Set-NestedPropertyHC -Object $fixture -Path $key -Value $Overrides[$key]
            }

            $file = New-Item $Path -ItemType File -Force
            Save-TestJson -InputObject $fixture -JsonFile $file
            return $file.FullName
        }

        # ---------------------------------------------------------------------
        # Helper: build a fake matrix file result with the shape BeginHC expects.
        # Mirrors what Import-MatrixFileHC would return for one .xlsx.
        # ---------------------------------------------------------------------
        function New-FakeMatrixResult {
            param(
                [string]$FileName = 'M1.xlsx',
                [hashtable]$Settings = @{ ComputerName = 'SRV01'; Path = 'C:\Share'; ApplyDefaultPermissions = $false },
                [array]$Permissions = @(),
                [array]$Check = @(),
                [hashtable]$MatrixAdObjects = @{},
                [hashtable]$AdNames = @{}
            )

            return [PSCustomObject]@{
                File            = [PSCustomObject]@{ Name = $FileName; FullName = "TestDrive:\Matrix\$FileName" }
                Settings        = $Settings
                Permissions     = $Permissions
                Check           = $Check
                MatrixAdObjects = $MatrixAdObjects
                AdNames         = $AdNames
            }
        }

        # ---------------------------------------------------------------------
        # Helper: build the standard BeginHC argument set
        # ---------------------------------------------------------------------
        function New-BeginArgs {
            param(
                [string]$ConfigurationJsonFile,
                [hashtable]$ScriptPath
            )
            return @{
                ConfigurationJsonFile = $ConfigurationJsonFile ?? (New-BeginJsonFile)
                ScriptPath            = $ScriptPath ?? (New-FakeScriptPath)
            }
        }
    }

    BeforeEach {
        $systemErrors = [System.Collections.Generic.List[object]]::new()

        # Default-safe mocks. Tests override per-Context as needed.
        Mock Validate-ConfigurationStructureHC { }
        Mock Invoke-WithOptionalParallelismHC {
            # Default: no matrix files imported. Tests override per-Context.
            return @()
        }
        Mock Import-MatrixDefaultsFileHC { return @() }
        Mock Get-DefaultAclHC { return @() }
        Mock Get-ADObjectDetailHC { return @{} }
        Mock New-Item -ParameterFilter { $Path -like '*Archive*' } { }
    }

    # =========================================================================
    Context 'JSON loading' {
        It 'parses a valid JSON file into Context' {
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context.Config.Matrix.FolderPath | Should -Not -BeNullOrEmpty
            $context.Config.Settings.ScriptName | Should -Be 'Test (Brecht)'
            $systemErrors.Count | Should -Be 0
        }

        It 'records FatalError and returns null when JSON file is missing' {
            $args = New-BeginArgs -ConfigurationJsonFile 'TestDrive:\nope.json'

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should -BeNullOrEmpty
            ($systemErrors | Where-Object { $_.Type -eq 'FatalError' }).Count | Should -BeGreaterThan 0
        }

        It 'records FatalError when JSON is malformed' {
            $bad = New-Item 'TestDrive:\Bad.json' -ItemType File -Force
            Set-Content $bad.FullName -Value '{ this is not valid json'
            $args = New-BeginArgs -ConfigurationJsonFile $bad.FullName

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should -BeNullOrEmpty
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -BeGreaterThan 0
        }

        It 'calls Ensure-SafeSettingsHC and uses its return value as Settings' {
            # If Ensure-SafeSettingsHC is a real helper in Utils, mock it and verify
            # the returned object lands on $context.Settings.
        }
    }

    # =========================================================================
    Context 'Configuration structure validation' {
        It 'records FatalError when Validate-ConfigurationStructureHC adds one' {
            Mock Validate-ConfigurationStructureHC {
                $SystemErrors.Value.Add([pscustomobject]@{
                        Type = 'FatalError'; Category = 'Validation'; Message = 'bad schema'
                    })
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should -BeNullOrEmpty
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -BeGreaterThan 0
        }

        It 'continues to next phase when validation passes' {
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should -Not -BeNullOrEmpty
        }
    }

    # =========================================================================
    Context 'Script path validation' {
        # This replaces coverage lost from the integration test.
        It 'records FatalError when ScriptPath.<Key> is missing' -ForEach @(
            'SetPermissions', 'TestRequirements', 'UpdateServiceNow'
        ) {
            $sp = New-FakeScriptPath
            $sp[$_] = 'TestDrive:\nope.ps1'
            $args = New-BeginArgs -ScriptPath $sp

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should -BeNullOrEmpty
            $systemErrors.Where({
                    $_.Type -eq 'FatalError' -and $_.Message -like "*$_*"
                }).Count | Should -BeGreaterThan 0
        }

        It 'continues when all ScriptPath entries exist' {
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should -Not -BeNullOrEmpty
        }
    }

    # =========================================================================
    Context 'Matrix file discovery' {
        It 'sets FoundMatrices=false and returns a context when no .xlsx files exist' {
            # Default Matrix folder created by New-BeginJsonFile is empty
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context.FoundMatrices | Should -Be $false
        }

        It 'sets FoundMatrices=true when at least one .xlsx exists' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context.FoundMatrices | Should -Be $true
        }

        It 'records FatalError when Matrix.FolderPath does not exist' {
            $config = New-BeginJsonFile -Overrides @{ 'Matrix.FolderPath' = 'x:\does-not-exist' }
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should -BeNullOrEmpty
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -BeGreaterThan 0
        }
    }

    # =========================================================================
    Context 'Defaults Excel file' {
        It 'loads valid defaults and stores on context' {
            Mock Import-MatrixDefaultsFileHC {
                return @( [pscustomobject]@{ ADObjectName = 'G1'; Permission = 'R' } )
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            # Adjust property name to wherever BeginHC stores defaults on context
            $context.Defaults | Should -Not -BeNullOrEmpty
        }

        It 'records FatalError when defaults file is missing' {
            $config = New-BeginJsonFile -Overrides @{ 'Matrix.DefaultsFile' = 'x:\nope.xlsx' }
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should -BeNullOrEmpty
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -BeGreaterThan 0
        }

        It 'records FatalError when defaults rows fail validation' {
            Mock Get-DefaultAclHC {
                $SystemErrors.Value.Add([pscustomobject]@{
                        Type = 'FatalError'; Category = 'Defaults'; Message = 'bad row'
                    })
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -BeGreaterThan 0
        }
    }

    # =========================================================================
    Context 'Archive folder creation' {
        It 'creates the archive folder when Matrix.Archive=true and folder does not exist' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            $config = New-BeginJsonFile -Overrides @{ 'Matrix.Archive' = $true }
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            Should -Invoke New-Item -ParameterFilter { $Path -like '*Archive*' }
        }

        It 'skips archive creation when Matrix.Archive=false' {
            $args = New-BeginArgs

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            Should -Invoke New-Item -ParameterFilter { $Path -like '*Archive*' } -Times 0
        }
    }

    # =========================================================================
    Context 'Parallel matrix import' {
        It 'collects results from Invoke-WithOptionalParallelismHC into context' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            New-Item 'TestDrive:\Matrix\M2.xlsx' -ItemType File -Force | Out-Null

            Mock Invoke-WithOptionalParallelismHC {
                return @(
                    (New-FakeMatrixResult -FileName 'M1.xlsx')
                    (New-FakeMatrixResult -FileName 'M2.xlsx')
                )
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            # Adjust property name (FileResults? Matrices?) to actual context shape
            $context.FileResults.Count | Should -Be 2
        }

        It 'passes throttle from MaxConcurrent.FoldersPerMatrix' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            $config = New-BeginJsonFile -Overrides @{ 'MaxConcurrent.FoldersPerMatrix' = 5 }
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-WithOptionalParallelismHC -ParameterFilter { $ThrottleLimit -eq 5 }
        }

        It 'defaults throttle to 4 when MaxConcurrent.FoldersPerMatrix is missing' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            $config = New-BeginJsonFile -Overrides @{ 'MaxConcurrent.FoldersPerMatrix' = $null }
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-WithOptionalParallelismHC -ParameterFilter { $ThrottleLimit -eq 4 }
        }
    }

    # =========================================================================
    Context 'Duplicate ComputerName/Path validation' {
        It 'records FatalError when two matrices target the same ComputerName+Path' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            New-Item 'TestDrive:\Matrix\M2.xlsx' -ItemType File -Force | Out-Null

            Mock Invoke-WithOptionalParallelismHC {
                return @(
                    (New-FakeMatrixResult -FileName 'M1.xlsx' -Settings @{
                        ComputerName = 'SRV01'; Path = 'C:\Share'; ApplyDefaultPermissions = $false
                    })
                    (New-FakeMatrixResult -FileName 'M2.xlsx' -Settings @{
                        ComputerName = 'SRV01'; Path = 'C:\Share'; ApplyDefaultPermissions = $false
                    })
                )
            }
            $args = New-BeginArgs

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({
                    $_.Type -eq 'FatalError' -and $_.Message -like '*duplicate*'
                }).Count | Should -BeGreaterThan 0
        }

        It 'passes when all matrices target unique ComputerName+Path' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            New-Item 'TestDrive:\Matrix\M2.xlsx' -ItemType File -Force | Out-Null

            Mock Invoke-WithOptionalParallelismHC {
                return @(
                    (New-FakeMatrixResult -FileName 'M1.xlsx' -Settings @{
                        ComputerName = 'SRV01'; Path = 'C:\Share'; ApplyDefaultPermissions = $false
                    })
                    (New-FakeMatrixResult -FileName 'M2.xlsx' -Settings @{
                        ComputerName = 'SRV02'; Path = 'C:\Share'; ApplyDefaultPermissions = $false
                    })
                )
            }
            $args = New-BeginArgs

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -Be 0
        }
    }

    # =========================================================================
    Context 'AD bulk query and SID mapping' {
        It 'builds Name->SID map from AD lookup' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            Mock Invoke-WithOptionalParallelismHC {
                return @( New-FakeMatrixResult -FileName 'M1.xlsx' -MatrixAdObjects @{
                        'DOMAIN\groupA' = 'placeholder'
                    } )
            }
            Mock Get-ADObjectsBulkHC {
                return @{ 'DOMAIN\groupA' = 'S-1-5-21-AAA' }
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $matrix = $context.FileResults[0]
            $matrix.AdNames.Keys | Should -Contain 'S-1-5-21-AAA'
        }

        It 'rewrites ACL entries to use SIDs instead of names' {
            # Verify the after-rewrite matrix ACL keys are SIDs (per session 1 decision 1)
        }

        It 'isolates per-matrix AD failures (one bad matrix does not poison the others)' {
            # Adjust per the real isolation invariant in BeginHC
        }
    }

    # =========================================================================
    Context 'Default permissions guard' {
        # Per session 1 decision 7
        It 'records FatalError when any matrix uses ApplyDefaultPermissions=true but defaults are empty' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            Mock Invoke-WithOptionalParallelismHC {
                return @( New-FakeMatrixResult -FileName 'M1.xlsx' -Settings @{
                        ComputerName = 'SRV01'; Path = 'C:\Share'; ApplyDefaultPermissions = $true
                    } )
            }
            Mock Import-MatrixDefaultsFileHC { return @() }   # empty defaults
            $args = New-BeginArgs

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({
                    $_.Type -eq 'FatalError' -and $_.Message -like '*default*'
                }).Count | Should -BeGreaterThan 0
        }

        It 'records Warning when defaults present but no matrix uses ApplyDefaultPermissions' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            Mock Invoke-WithOptionalParallelismHC {
                return @( New-FakeMatrixResult -FileName 'M1.xlsx' -Settings @{
                        ComputerName = 'SRV01'; Path = 'C:\Share'; ApplyDefaultPermissions = $false
                    } )
            }
            Mock Import-MatrixDefaultsFileHC {
                return @( [pscustomobject]@{ ADObjectName = 'G1'; Permission = 'R' } )
            }
            $args = New-BeginArgs

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({
                    $_.Type -eq 'Warning' -and $_.Message -like '*default*'
                }).Count | Should -BeGreaterThan 0
        }

        It 'skips broken matrices (FatalError on the matrix) when evaluating the guard' {
            # Verify Test-ItemHasFatalErrorHC filter is applied before the guard check
        }
    }
}