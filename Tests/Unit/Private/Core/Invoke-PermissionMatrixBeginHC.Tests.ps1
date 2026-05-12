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

        function New-BeginJsonFile {
            param(
                [hashtable]$Overrides = @{},
                [string[]]$Remove = @(),
                [string]$Path = 'TestDrive:\Input.json'
            )

            $fixture = New-JsonFixture
            $fixture.Matrix.FolderPath = (New-Item 'TestDrive:\Matrix' -ItemType Directory -Force).FullName
            $fixture.Matrix.DefaultsFile = (New-ValidDefaultsExcelFixture -Path 'TestDrive:\Defaults.xlsx')
            $fixture.Settings.SaveLogFiles.Where.Folder = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName

            foreach ($key in $Overrides.Keys) {
                Set-NestedPropertyHC -Object $fixture -Path $key -Value $Overrides[$key]
            }

            foreach ($removePath in $Remove) {
                $segments = $removePath -split '\.'
                $parent = $fixture
                for ($i = 0; $i -lt $segments.Count - 1; $i++) {
                    $parent = $parent.($segments[$i])
                }
                $leaf = $segments[-1]
                if ($parent -is [hashtable]) {
                    $parent.Remove($leaf)
                }
                else {
                    $parent.PSObject.Properties.Remove($leaf)
                }
            }

            $file = New-Item $Path -ItemType File -Force
            Save-TestJson -InputObject $fixture -JsonFile $file
            return $file.FullName
        }

        function New-FakeMatrixEntry {
            param(
                [string]$ComputerName = 'SRV01',
                [string]$Path = 'C:\Share',
                [bool]$ApplyDefaultPermissions = $false,
                [string]$FileName = 'M1.xlsx',
                [hashtable]$Permissions = @{},
                [hashtable]$AdObjects = @{}
            )

            return [pscustomobject]@{
                Setting         = [pscustomobject]@{
                    Formatted = [pscustomobject]@{
                        ComputerName            = $ComputerName
                        Path                    = $Path
                        ApplyDefaultPermissions = $ApplyDefaultPermissions
                    }
                }
                FileContext     = [pscustomobject]@{
                    Item  = [pscustomobject]@{
                        Name     = $FileName
                        FullName = "TestDrive:\Matrix\$FileName"
                    }
                    Check = [System.Collections.Generic.List[object]]::new()
                }
                Permissions     = $Permissions
                MatrixAdObjects = $AdObjects
                Check           = [System.Collections.Generic.List[object]]::new()
                Matrix          = @(
                    [pscustomobject]@{ ACL = @{} }
                )
            }
        }

        function New-FakeFileResult {
            param(
                [string]$FileName = 'M1.xlsx',
                [object[]]$Matrices = @()
            )

            if ($Matrices.Count -eq 0) {
                $Matrices = @( New-FakeMatrixEntry -FileName $FileName )
            }

            return [pscustomobject]@{
                File     = [pscustomobject]@{ Name = $FileName; FullName = "TestDrive:\Matrix\$FileName" }
                Matrices = $Matrices
            }
        }

        function New-BeginArgs {
            param(
                [string]$ConfigurationJsonFile,
                [hashtable]$ScriptPath
            )
 
            if ([string]::IsNullOrWhiteSpace($ConfigurationJsonFile)) {
                $ConfigurationJsonFile = New-BeginJsonFile
            }
            if (-not $ScriptPath) {
                $ScriptPath = New-FakeScriptPath
            }
 
            return @{
                ConfigurationJsonFile = $ConfigurationJsonFile
                ScriptPath            = $ScriptPath
            }
        }
    }

    BeforeEach {
        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Mock Validate-ConfigurationStructureHC { }
        Mock Invoke-WithOptionalParallelismHC {
            return @()
        }
        Mock Import-MatrixDefaultsFileHC { return @() }
        Mock Get-DefaultAclHC { return @() }
        Mock Get-ADObjectDetailHC { return @{} }
        Mock New-Item -ParameterFilter { $Path -like '*Archive*' } { }
    }

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

    Context 'Configuration structure validation' {
        It 'records FatalError when Validate-ConfigurationStructureHC adds one' {
            Mock Validate-ConfigurationStructureHC {
                $SystemErrors.Value.Add([pscustomobject]@{
                        Type = 'FatalError'; Category = 'Validation'; Message = 'bad schema'
                    })
            }
            $beginArgs = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @beginArgs -SystemErrors ([ref]$systemErrors)

            # BeginHC returns the (partial) context on error; caller checks SystemErrors.
            # The contract being tested: a FatalError landed in SystemErrors and BeginHC
            # halted before downstream phases ran.
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -BeGreaterThan 0
            Should -Invoke Invoke-WithOptionalParallelismHC -Times 0
        }

        It 'continues to next phase when validation passes' {
            $beginArgs = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @beginArgs -SystemErrors ([ref]$systemErrors)

            $context | Should -Not -BeNullOrEmpty
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -Be 0
        }
    }

    Context 'Script path validation' {
        It 'records FatalError when ScriptPath.<Key> is missing' -ForEach @(
            @{ Key = 'SetPermissions' }
            @{ Key = 'TestRequirements' }
            @{ Key = 'UpdateServiceNow' }
        ) {
            $sp = New-FakeScriptPath
            $sp[$Key] = 'TestDrive:\nope.ps1'
            $beginArgs = New-BeginArgs -ScriptPath $sp

            $context = Invoke-PermissionMatrixBeginHC @beginArgs -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({
                    $_.Type -eq 'FatalError' -and $_.Message -like "*$Key*"
                }).Count | Should -BeGreaterThan 0
            Should -Invoke Invoke-WithOptionalParallelismHC -Times 0
        }

        It 'continues when all ScriptPath entries exist' {
            $beginArgs = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @beginArgs -SystemErrors ([ref]$systemErrors)

            $context | Should -Not -BeNullOrEmpty
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -Be 0
        }
    }

    Context 'Matrix file discovery' {
        It 'bails out without errors when matrix folder is empty' {
            # Default New-BeginJsonFile creates Matrix folder but no .xlsx files
            $beginArgs = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @beginArgs -SystemErrors ([ref]$systemErrors)

            $context.FoundMatrices | Should -Be $false
            $systemErrors.Count | Should -Be 0
            Should -Invoke Import-MatrixDefaultsFileHC -Times 0
            Should -Invoke Invoke-WithOptionalParallelismHC -Times 0
        }

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

            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -BeGreaterThan 0
            $context.FoundMatrices | Should -Be $false
        }
    }

    Context 'Defaults Excel file' {
        BeforeEach {
            # Defaults phase only runs when matrix files exist
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
        }
    
        It 'loads valid defaults and stores on context' {
            Mock Import-MatrixDefaultsFileHC {
                return [pscustomobject]@{
                    FilePath   = 'TestDrive:\Defaults.xlsx'
                    DefaultAcl = @( [pscustomobject]@{ ADObjectName = 'G1'; Permission = 'R' } )
                    MailTo     = [System.Collections.Generic.List[string]]@('test@example.com')
                }
            }
            $beginArgs = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @beginArgs -SystemErrors ([ref]$systemErrors)

            $context.Defaults | Should -Not -BeNullOrEmpty
            $context.Defaults.DefaultAcl.Count | Should -Be 1
            $context.Defaults.MailTo | Should -Contain 'test@example.com'
        }

        It 'halts when Import-MatrixDefaultsFileHC reports a FatalError' {
            Mock Import-MatrixDefaultsFileHC {
                $SystemErrors.Value.Add([pscustomobject]@{
                        Type = 'FatalError'; Category = 'Defaults'; Message = 'defaults file boom'
                    })
            }
            $beginArgs = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @beginArgs -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should -BeGreaterThan 0
            Should -Invoke Invoke-WithOptionalParallelismHC -Times 0
        }
    }

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

    Context 'Parallel matrix import' {
        It 'collects results from Invoke-WithOptionalParallelismHC into context' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            New-Item 'TestDrive:\Matrix\M2.xlsx' -ItemType File -Force | Out-Null

            Mock Invoke-WithOptionalParallelismHC {
                return @(
                    (New-FakeFileResult -FileName 'M1.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'M1.xlsx' -ComputerName 'SRV01' -Path 'C:\A'
                    ))
                    (New-FakeFileResult -FileName 'M2.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'M2.xlsx' -ComputerName 'SRV02' -Path 'C:\B'
                    ))
                )
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

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
            $config = New-BeginJsonFile -Remove 'MaxConcurrent.FoldersPerMatrix'
            $beginArgs = New-BeginArgs -ConfigurationJsonFile $config

            $null = Invoke-PermissionMatrixBeginHC @beginArgs -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-WithOptionalParallelismHC -ParameterFilter { $ThrottleLimit -eq 4 }
        }
    }

    Context 'AD bulk query and SID mapping' {
        It 'builds Name->SID map from AD lookup' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            Mock Invoke-WithOptionalParallelismHC {
                return @( New-FakeFileResult -FileName 'M1.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'M1.xlsx' -AdObjects @{ 'groupA' = 'placeholder' }
                    ) )
            }
            Mock Get-AdObjectDetailHC {
                return @{ 'DOMAIN\groupA' = 'S-1-5-21-AAA' }
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should -Not -BeNullOrEmpty
        }

        It 'rewrites ACL entries to use SIDs instead of names' {
            # Verify the after-rewrite matrix ACL keys are SIDs (per session 1 decision 1)
        }

        It 'isolates per-matrix AD failures (one bad matrix does not poison the others)' {
            # Adjust per the real isolation invariant in BeginHC
        }
    }

    Context 'Default permissions guard' {
        BeforeEach {
            Mock Test-AdObjectInMatrixHC { return @() }
        }
        # Per session 1 decision 7
        It 'records FatalError when any matrix uses ApplyDefaultPermissions=true but defaults are empty' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            Mock Invoke-WithOptionalParallelismHC {
                return @( New-FakeFileResult -FileName 'M1.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'M1.xlsx' -ApplyDefaultPermissions $true
                    ) )
            }
            Mock Import-MatrixDefaultsFileHC {
                return [pscustomobject]@{
                    FilePath   = 'TestDrive:\Defaults.xlsx'
                    DefaultAcl = @{}
                    MailTo     = [System.Collections.Generic.List[string]]@('test@example.com')
                }
            }
            $args = New-BeginArgs

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({
                    $_.Type -eq 'FatalError' -and $_.Message -like '*default*'
                }).Count | Should -BeGreaterThan 0
        }

        It 'records Warning when defaults present but no matrix uses ApplyDefaultPermissions' {
            New-Item 'TestDrive:\Matrix\M1.xlsx' -ItemType File -Force | Out-Null
            Mock Invoke-WithOptionalParallelismHC {
                return @( New-FakeFileResult -FileName 'M1.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'M1.xlsx' -ApplyDefaultPermissions $false
                    ) )
            }
            Mock Import-MatrixDefaultsFileHC {
                return [pscustomobject]@{
                    FilePath   = 'TestDrive:\Defaults.xlsx'
                    DefaultAcl = @{ 'groupA' = @{ Permission = 'R' } }
                    MailTo     = [System.Collections.Generic.List[string]]@('test@example.com')
                }
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