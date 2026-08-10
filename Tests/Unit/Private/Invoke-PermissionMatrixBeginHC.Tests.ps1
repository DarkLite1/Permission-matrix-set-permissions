#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

Describe 'Invoke-PermissionMatrixBeginHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$root\Tests\Helpers\Helpers.HC.ps1"
        . "$root\Tests\Helpers\Fixtures.Json.ps1"
        . "$root\Tests\Helpers\Fixtures.Excel.ps1"

        Get-ChildItem -Path "$moduleRoot\Private" -Filter '*.ps1' -File |
        ForEach-Object { . $_.FullName }

        function New-FakeScriptPath {
            param([string]$Root = $TestDrive)

            return @{
                PermissionMatrixModule = (New-Item (Join-Path $Root 'PermissionMatrix.psm1') -ItemType File -Force).FullName
                SetPermissions         = (New-Item (Join-Path $Root 'SetPermissions.ps1') -ItemType File -Force).FullName
                TestRequirements       = (New-Item (Join-Path $Root 'TestRequirements.ps1') -ItemType File -Force).FullName
                UpdateServiceNow       = (New-Item (Join-Path $Root 'UpdateServiceNow.ps1') -ItemType File -Force).FullName
            }
        }

        function New-BeginJsonFile {
            param(
                [hashtable]$Overrides = @{},
                [string[]]$Remove = @(),
                [string]$Path = (Join-Path $TestDrive 'Input.json')
            )

            $fixture = New-JsonFixture
            $fixture.Matrix.FolderPath = (New-Item (Join-Path $TestDrive 'Matrix') -ItemType Directory -Force).FullName
            $fixture.Matrix.DefaultsFile = (New-ValidDefaultsExcelFixture -Path (Join-Path $TestDrive 'Defaults.xlsx'))
            $fixture.Settings.SaveLogFiles.Where.Folder = (New-Item (Join-Path $TestDrive 'Logs') -ItemType Directory -Force).FullName

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
                if ($parent -is [hashtable]) { $parent.Remove($leaf) }
                else { $parent.PSObject.Properties.Remove($leaf) }
            }

            $file = New-Item $Path -ItemType File -Force
            Save-TestJson -InputObject $fixture -JsonFile $file
            return $file.FullName
        }

        # One matrix entry inside parallelResults[].Matrices.
        # Matrix must be non-empty: Test-AdObjectInMatrixHC declares -Matrix as
        # Mandatory and the binder rejects empty arrays before the mock can
        # intercept.
        function New-FakeMatrixEntry {
            param(
                [string]$ComputerName = 'SRV01',
                [string]$Path = 'C:\Share',
                [bool]$ApplyDefaultPermissions = $false,
                [string]$FileName = 'M1.xlsx',
                [hashtable]$Permissions = @{},
                [hashtable]$AdObjects = @{},
                [hashtable]$Acl = @{},
                [object[]]$Check = @(),
                [object[]]$FileContextCheck = @()
            )

            $checkList = [System.Collections.Generic.List[object]]::new()
            foreach ($c in $Check) { $checkList.Add($c) }

            $fileContextCheckList = [System.Collections.Generic.List[object]]::new()
            foreach ($c in $FileContextCheck) { $fileContextCheckList.Add($c) }

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
                    Check = $fileContextCheckList
                }
                Permissions     = $Permissions
                MatrixAdObjects = $AdObjects
                Check           = $checkList
                Matrix          = @(
                    [pscustomobject]@{ ACL = $Acl }
                )
            }
        }

        function New-FakeFileResult {
            param(
                [string]$FileName = 'M1.xlsx',
                [object[]]$Matrices = @(),
                [object[]]$Check = @(),
                [object[]]$PermissionsCheck = @()
            )

            if ($Matrices.Count -eq 0) {
                $Matrices = @( New-FakeMatrixEntry -FileName $FileName )
            }

            $checkList = [System.Collections.Generic.List[object]]::new()
            foreach ($c in $Check) { $checkList.Add($c) }

            $permissionsCheckList = [System.Collections.Generic.List[object]]::new()
            foreach ($c in $PermissionsCheck) { $permissionsCheckList.Add($c) }

            return [pscustomobject]@{
                File     = [pscustomobject]@{ Name = $FileName; FullName = "TestDrive:\Matrix\$FileName" }
                Item     = [pscustomobject]@{ Name = $FileName; FullName = "TestDrive:\Matrix\$FileName" }
                Check    = $checkList
                Sheets   = @{
                    Permissions = @{ Check = $permissionsCheckList }
                }
                Matrices = $Matrices
            }
        }

        # Shorthand for Import-MatrixDefaultsFileHC's structured return.
        # DefaultAcl is a hashtable (BeginHC reads .Keys and .Count on it).
        function New-FakeDefaults {
            param(
                [hashtable]$DefaultAcl = @{},
                [string[]]$MailTo = @('test@example.com')
            )

            return [pscustomobject]@{
                FilePath   = (Join-Path $TestDrive 'Defaults.xlsx')
                DefaultAcl = $DefaultAcl
                MailTo     = [System.Collections.Generic.List[string]]@($MailTo)
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

        # Default-safe mocks. Each Context overrides as needed.
        Mock Test-ConfigurationStructureHC { }
        Mock Invoke-WithOptionalParallelismHC { return @() }
        Mock Import-MatrixDefaultsFileHC { return @() }
        Mock Get-DefaultAclHC { return @() }
        Mock Get-ADObjectDetailHC { return @{} }
    }

    Context 'JSON loading' {
        It 'parses a valid JSON file into Context' {
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context.Config.Matrix.FolderPath | Should-BeTruthy
            $context.Config.Settings.ScriptName | Should-Be 'Test (Brecht)'
            $systemErrors.Count | Should-Be 0
        }

        It 'records FatalError and returns null when JSON file is missing' {
            $args = New-BeginArgs -ConfigurationJsonFile (Join-Path $TestDrive 'nope.json')

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should-BeFalsy
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should-BeGreaterThan 0
        }

        It 'records FatalError when JSON is malformed' {
            $bad = New-Item (Join-Path $TestDrive 'Bad.json') -ItemType File -Force
            Set-Content $bad.FullName -Value '{ this is not valid json'
            $args = New-BeginArgs -ConfigurationJsonFile $bad.FullName

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should-BeFalsy
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should-BeGreaterThan 0
        }
    }

    Context 'Configuration structure validation' {
        It 'records FatalError when Test-ConfigurationStructureHC adds one' {
            Mock Test-ConfigurationStructureHC {
                $SystemErrors.Value.Add([pscustomobject]@{
                        Type = 'FatalError'; Category = 'Validation'; Message = 'bad schema'
                    })
            }
            $args = New-BeginArgs

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should-BeGreaterThan 0
            Should-Invoke Invoke-WithOptionalParallelismHC -Times 0
        }

        It 'continues to next phase when validation passes' {
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should-BeTruthy
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should-Be 0
        }
    }

    Context 'Script path validation' {
        It 'records FatalError when ScriptPath.<Key> is missing' -ForEach @(
            @{ Key = 'SetPermissions' }
            @{ Key = 'TestRequirements' }
            @{ Key = 'UpdateServiceNow' }
        ) {
            $sp = New-FakeScriptPath
            $sp[$Key] = (Join-Path $TestDrive 'nope.ps1')
            $args = New-BeginArgs -ScriptPath $sp

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({
                    $_.Type -eq 'FatalError' -and $_.Message -like "*$Key*"
                }).Count | Should-BeGreaterThan 0
            Should-Invoke Invoke-WithOptionalParallelismHC -Times 0
        }

        It 'continues when all ScriptPath entries exist' {
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should-BeTruthy
            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should-Be 0
        }
    }

    Context 'Matrix file discovery' {
        It 'bails out cleanly when matrix folder is empty' {
            # New-BeginJsonFile creates the Matrix folder but no .xlsx files.
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context.FoundMatrices | Should-BeFalse
            $systemErrors.Count | Should-Be 0
            Should-Invoke Import-MatrixDefaultsFileHC -Times 0
            Should-Invoke Invoke-WithOptionalParallelismHC -Times 0
        }

        It 'excludes the defaults workbook when its path is not canonical' {
            <# Regression: the exclusion compared the raw configured string
            against FileInfo.FullName. A DefaultsFile written with '.\'
            segments or forward slashes is the same file on disk but a
            different string, so the defaults workbook was imported as if it
            were an ordinary matrix file. #>
            $matrixFolder = Join-Path $TestDrive 'Matrix'

            New-Item (Join-Path $matrixFolder 'M1.xlsx') -ItemType File -Force | Out-Null
            New-ValidDefaultsExcelFixture -Path (
                Join-Path $matrixFolder 'Defaults.xlsx'
            ) | Out-Null

            # Same file, deliberately denormalized.
            $denormalized = Join-Path $matrixFolder '.\Defaults.xlsx'

            $config = New-BeginJsonFile -Overrides @{
                'Matrix.DefaultsFile' = $denormalized
            }
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            Should-Invoke Invoke-WithOptionalParallelismHC -ParameterFilter {
                @($InputObject).Count -eq 1 -and
                @($InputObject)[0].Name -eq 'M1.xlsx'
            }
        }

        It 'sets FoundMatrices=true when at least one .xlsx exists' {
            New-Item (Join-Path $TestDrive 'Matrix\M1.xlsx') -ItemType File -Force | Out-Null
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context.FoundMatrices | Should-BeTrue
        }

        It 'records FatalError when Matrix.FolderPath does not exist' {
            $config = New-BeginJsonFile -Overrides @{ 'Matrix.FolderPath' = 'x:\does-not-exist' }
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should-BeGreaterThan 0
            $context.FoundMatrices | Should-BeFalse
        }
    }

    Context 'Defaults Excel file' {
        BeforeEach {
            # Defaults phase only runs when matrix files exist.
            New-Item (Join-Path $TestDrive 'Matrix\M1.xlsx') -ItemType File -Force | Out-Null
        }

        It 'loads valid defaults and stores on context' {
            Mock Import-MatrixDefaultsFileHC {
                New-FakeDefaults -DefaultAcl @{ 'groupA' = @{ Permission = 'R' } }
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context.Defaults | Should-BeTruthy
            $context.Defaults.DefaultAcl.Count | Should-Be 1
            $context.Defaults.MailTo | Should-ContainCollection 'test@example.com'
        }

        It 'halts when Import-MatrixDefaultsFileHC reports a FatalError' {
            Mock Import-MatrixDefaultsFileHC {
                $SystemErrors.Value.Add([pscustomobject]@{
                        Type = 'FatalError'; Category = 'Defaults'; Message = 'defaults file boom'
                    })
            }
            $args = New-BeginArgs

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should-BeGreaterThan 0
            Should-Invoke Invoke-WithOptionalParallelismHC -Times 0
        }
    }

    Context 'Archive folder creation' {
        BeforeEach {
            New-Item (Join-Path $TestDrive 'Matrix\M1.xlsx') -ItemType File -Force | Out-Null
        }

        It 'creates the archive folder when Matrix.Archive=true' {
            $config = New-BeginJsonFile -Overrides @{ 'Matrix.Archive' = $true }
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)
 
            $archivePath = Join-Path $context.Config.Matrix.FolderPath 'Archive'
            Test-Path -LiteralPath $archivePath -PathType Container | Should-BeTrue
        }

        It 'skips archive creation when Matrix.Archive=false' {
            $matrixFolder = (New-Item (Join-Path $TestDrive 'NoArchiveMatrix') -ItemType Directory -Force).FullName
            $config = New-BeginJsonFile `
                -Overrides @{ 'Matrix.FolderPath' = $matrixFolder } `
                -Path (Join-Path $TestDrive 'NoArchiveInput.json')
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)
 
            $archivePath = Join-Path $context.Config.Matrix.FolderPath 'Archive'
            Test-Path -LiteralPath $archivePath -PathType Container | Should-BeFalse
        }
    }

    Context 'Parallel matrix import' {
        BeforeEach {
            New-Item (Join-Path $TestDrive 'Matrix\M1.xlsx') -ItemType File -Force | Out-Null
        }

        It 'collects results from Invoke-WithOptionalParallelismHC into context' {
            New-Item (Join-Path $TestDrive 'Matrix\M2.xlsx') -ItemType File -Force | Out-Null

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

            $context.FileResults.Count | Should-Be 2
        }

        It 'passes throttle from MaxConcurrent.FoldersPerMatrix' {
            $config = New-BeginJsonFile -Overrides @{ 'MaxConcurrent.FoldersPerMatrix' = 5 }
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            Should-Invoke Invoke-WithOptionalParallelismHC -ParameterFilter { $ThrottleLimit -eq 5 }
        }

        It 'defaults throttle to 4 when MaxConcurrent.FoldersPerMatrix is missing' {
            $config = New-BeginJsonFile -Remove 'MaxConcurrent.FoldersPerMatrix'
            $args = New-BeginArgs -ConfigurationJsonFile $config

            $null = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            Should-Invoke Invoke-WithOptionalParallelismHC -ParameterFilter { $ThrottleLimit -eq 4 }
        }
    }

    Context 'AD bulk query and SID mapping' {
        BeforeEach {
            New-Item (Join-Path $TestDrive 'Matrix\M1.xlsx') -ItemType File -Force | Out-Null
            Mock Test-AdObjectInMatrixHC { return @() }
        }

        It 'builds Name->SID map from AD lookup' {
            Mock Invoke-WithOptionalParallelismHC {
                return @( New-FakeFileResult -FileName 'M1.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'M1.xlsx' -AdObjects @{ 'groupA' = 'placeholder' }
                    ) )
            }
            Mock Get-ADObjectDetailHC {
                return @{ 'DOMAIN\groupA' = 'S-1-5-21-AAA' }
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $context | Should-BeTruthy
        }

        It 'rewrites ACL entries to use SIDs instead of names' {
            $fileResult = New-FakeFileResult -FileName 'M1.xlsx' -Matrices @(
                New-FakeMatrixEntry -FileName 'M1.xlsx' `
                    -AdObjects @{ 'groupA' = 'placeholder' } `
                    -Acl @{ 'groupA' = @{ Permission = 'R' } }
            )
            Mock Invoke-WithOptionalParallelismHC { return @( $fileResult ) }
            Mock Get-ADObjectDetailHC {
                return @(
                    @{ SamAccountName = 'groupA'; adObject = @{ ObjectSid = 'S-1-5-21-AAA' } }
                )
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $folder = $context.AllMatrices[0].Matrix[0]
            $folder.ACL.Keys | Should-ContainCollection 'S-1-5-21-AAA'
            $folder.ACL.Keys | Should-NotContainCollection 'groupA'
            $folder.AdNames['S-1-5-21-AAA'] | Should-Be 'groupA'
        }

        It 'isolates per-matrix AD failures (broken matrix does not poison the others)' {
            # Two matrices: one broken (FatalError in .Check), one clean.
            # The broken one should be skipped during ACL rewrite; the clean one
            # should still have its ACL rewritten to SIDs.
            $brokenMatrix = New-FakeMatrixEntry -FileName 'Broken.xlsx' `
                -ComputerName 'SRV01' -Path 'C:\Broken' `
                -AdObjects @{ 'groupA' = 'placeholder' } `
                -Acl @{ 'groupA' = @{ Permission = 'R' } } `
                -Check @( [pscustomobject]@{ Type = 'FatalError'; Name = 'Pre-existing'; Message = 'broken' } )

            $cleanMatrix = New-FakeMatrixEntry -FileName 'Clean.xlsx' `
                -ComputerName 'SRV02' -Path 'C:\Clean' `
                -AdObjects @{ 'groupB' = 'placeholder' } `
                -Acl @{ 'groupB' = @{ Permission = 'W' } }

            Mock Invoke-WithOptionalParallelismHC {
                return @(
                    (New-FakeFileResult -FileName 'Broken.xlsx' -Matrices @($brokenMatrix))
                    (New-FakeFileResult -FileName 'Clean.xlsx' -Matrices @($cleanMatrix))
                )
            }
            Mock Get-ADObjectDetailHC {
                return @(
                    @{ SamAccountName = 'groupA'; adObject = @{ ObjectSid = 'S-1-5-21-AAA' } }
                    @{ SamAccountName = 'groupB'; adObject = @{ ObjectSid = 'S-1-5-21-BBB' } }
                )
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            # Clean matrix got its ACL rewritten to SIDs
            $clean = $context.AllMatrices | Where-Object { $_.Setting.Formatted.Path -eq 'C:\Clean' }
            $clean.Matrix[0].ACL.Keys | Should-ContainCollection 'S-1-5-21-BBB'

            # Broken matrix was skipped — ACL keys remain unchanged (still name, not SID)
            $broken = $context.AllMatrices | Where-Object { $_.Setting.Formatted.Path -eq 'C:\Broken' }
            $broken.Matrix[0].ACL.Keys | Should-ContainCollection 'groupA'
            $broken.Matrix[0].ACL.Keys | Should-NotContainCollection 'S-1-5-21-AAA'
        }
    }

    Context 'Default permissions guard' {
        # Per session 1 decision 7: ApplyDefaultPermissions=true requires defaults;
        # defaults without any consumer logs an information record. The guard is
        # evaluated per matrix file (ApplyDefaultPermissions can differ per file),
        # so the resulting check lands on that file's own Check list.
        BeforeEach {
            New-Item (Join-Path $TestDrive 'Matrix\M1.xlsx') -ItemType File -Force | Out-Null
            Mock Test-AdObjectInMatrixHC { return @() }
        }

        It 'records FatalError on the file when any of its rows use ApplyDefaultPermissions=true but defaults are empty' {
            Mock Invoke-WithOptionalParallelismHC {
                return @( New-FakeFileResult -FileName 'M1.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'M1.xlsx' -ApplyDefaultPermissions $true
                    ) )
            }
            Mock Import-MatrixDefaultsFileHC { New-FakeDefaults -DefaultAcl @{} }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            @($context.FileResults.Check).Where({
                    $_.Type -eq 'FatalError' -and $_.Name -eq 'Empty default ACL'
                }).Count | Should-BeGreaterThan 0
        }

        It 'records Information on the file when defaults present but none of its rows use ApplyDefaultPermissions' {
            Mock Invoke-WithOptionalParallelismHC {
                return @( New-FakeFileResult -FileName 'M1.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'M1.xlsx' -ApplyDefaultPermissions $false
                    ) )
            }
            Mock Import-MatrixDefaultsFileHC {
                New-FakeDefaults -DefaultAcl @{ 'groupA' = @{ Permission = 'R' } }
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            @($context.FileResults.Check).Where({
                    $_.Type -eq 'Information' -and $_.Name -eq 'Unused defaults'
                }).Count | Should-BeGreaterThan 0
        }

        It 'evaluates each matrix file independently (one uses defaults, the other does not)' {
            # File A has a row with ApplyDefaultPermissions=TRUE -> consumes the
            # defaults, no 'Unused defaults'. File B has no such row -> defaults
            # are unused for File B, so it gets the Information record. Proves the
            # guard is per file, not global.
            Mock Invoke-WithOptionalParallelismHC {
                return @(
                    (New-FakeFileResult -FileName 'A.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'A.xlsx' -ComputerName 'SRV01' -Path 'C:\A' -ApplyDefaultPermissions $true
                    ))
                    (New-FakeFileResult -FileName 'B.xlsx' -Matrices @(
                        New-FakeMatrixEntry -FileName 'B.xlsx' -ComputerName 'SRV02' -Path 'C:\B' -ApplyDefaultPermissions $false
                    ))
                )
            }
            Mock Import-MatrixDefaultsFileHC {
                New-FakeDefaults -DefaultAcl @{ 'groupA' = @{ Permission = 'R' } }
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            $fileA = $context.FileResults | Where-Object { $_.Item.Name -eq 'A.xlsx' }
            $fileB = $context.FileResults | Where-Object { $_.Item.Name -eq 'B.xlsx' }

            @($fileA.Check).Where({ $_.Name -eq 'Unused defaults' }).Count | Should-Be 0
            @($fileB.Check).Where({
                    $_.Type -eq 'Information' -and $_.Name -eq 'Unused defaults'
                }).Count | Should-BeGreaterThan 0
        }

        It 'skips broken rows (FatalError on the row) when evaluating the guard' {
            # A row flagged with a FatalError is excluded from the guard, so its
            # ApplyDefaultPermissions=true does not count. With only the clean
            # row left ($fileUsesDefaults is null) and defaults present, the
            # Information record fires on the file.
            $brokenRow = New-FakeMatrixEntry -FileName 'M1.xlsx' `
                -ComputerName 'SRV01' -Path 'C:\Broken' `
                -ApplyDefaultPermissions $true `
                -Check @( [pscustomobject]@{ Type = 'FatalError'; Name = 'Pre-existing'; Message = 'broken' } )

            $cleanRow = New-FakeMatrixEntry -FileName 'M1.xlsx' `
                -ComputerName 'SRV02' -Path 'C:\Clean' `
                -ApplyDefaultPermissions $false

            Mock Invoke-WithOptionalParallelismHC {
                return @(
                    (New-FakeFileResult -FileName 'M1.xlsx' -Matrices @($brokenRow, $cleanRow))
                )
            }
            Mock Import-MatrixDefaultsFileHC {
                New-FakeDefaults -DefaultAcl @{ 'groupA' = @{ Permission = 'R' } }
            }
            $args = New-BeginArgs

            $context = Invoke-PermissionMatrixBeginHC @args -SystemErrors ([ref]$systemErrors)

            @($context.FileResults.Check).Where(
                {
                    $_.Type -eq 'Information' -and $_.Name -eq 'Unused defaults'
                }
            ).Count | Should-BeGreaterThan 0
        }
    }
}

Describe 'Invoke-PermissionMatrixBeginHC defaults merge (real merge loop)' {
    # The defaults-merge loop runs INSIDE the per-file scriptblock that
    # Invoke-WithOptionalParallelismHC executes, so the mocked unit tests above
    # (which stub Invoke-WithOptionalParallelismHC) never reach it. This
    # Describe drives the real merge with real Excel fixtures (sequential,
    # FoldersPerMatrix = 1) and inspects the matrix the orchestrator produced,
    # proving an inherit-only folder (no permission cells) never receives the
    # default ACL even when ApplyDefaultPermissions = TRUE.
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$root\Tests\Helpers\Helpers.HC.ps1"
        . "$root\Tests\Helpers\Fixtures.Json.ps1"
        . "$root\Tests\Helpers\Fixtures.Excel.ps1"

        Get-ChildItem -Path "$moduleRoot\Private" -Filter '*.ps1' -File |
        ForEach-Object { . $_.FullName }

        # Fake SIDs so the post-merge SID rewrite succeeds without touching AD.
        $script:sidBob = 'S-1-5-21-1111111111-2222222222-3333333333-1001'
        $script:sidMike = 'S-1-5-21-1111111111-2222222222-3333333333-1002'
        $script:sidDefault = 'S-1-5-21-1111111111-2222222222-3333333333-1003'
    }

    It 'does not merge the default ACL into a folder without permissions' {
        $matrixDir = (New-Item (Join-Path $TestDrive 'MergeMatrix') -ItemType Directory -Force).FullName
        $logsDir = (New-Item (Join-Path $TestDrive 'MergeLogs') -ItemType Directory -Force).FullName

        #region Real defaults + matrix Excel files
        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'WithEmptyPermissionFolder') `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'HQ'
                SiteCode                = 'HQ'
                ComputerName            = $env:COMPUTERNAME
                Path                    = (Join-Path $TestDrive 'Target')
                GroupName               = 'Team-A'
                Action                  = 'Check'
                ApplyDefaultPermissions = $true
            }
        ) | Out-Null
        #endregion

        #region Config JSON pointing at the real fixtures (sequential)
        $configFixture = New-JsonFixture
        $configFixture.Matrix.FolderPath = $matrixDir
        $configFixture.Matrix.DefaultsFile = $defaultsPath
        $configFixture.Settings.SaveLogFiles.Where.Folder = $logsDir
        $configFixture.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        Save-TestJson -InputObject $configFixture -JsonFile $configPath
        #endregion

        $scriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
            SetPermissions         = "$root\Scripts\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Scripts\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Scripts\Operations\UpdateServiceNow.ps1"
        }

        # Resolve every AD name the fixtures use so the SID rewrite runs; the
        # merge itself happens before this and is what we are validating.
        Mock Get-ADObjectDetailHC {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = $sidBob } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = $sidMike } }
                @{ SamAccountName = 'DefaultGroup'; adObject = @{ ObjectSid = $sidDefault } }
            )
        }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        $context = Invoke-PermissionMatrixBeginHC `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrors)

        #region No fatal errors and the matrix was built
        $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should-Be 0 -Because 'the fixture is valid and defaults do not conflict'

        $context.AllMatrices.Count | Should-Be 1
        $matrix = $context.AllMatrices[0].Matrix
        #endregion

        #region The inherit-only folder carries no ACL and no default
        $inheritOnly = $matrix | Where-Object { $_.Path -eq 'InheritOnly' }

        $inheritOnly | Should-BeTruthy `
            -Because 'the empty-permission folder must still be present in the matrix'
        $inheritOnly.Ignore | Should-BeFalse `
            -Because 'a blank row is inherit-only, not ignored'
        $inheritOnly.ACL.Count | Should-Be 0 `
            -Because 'the orchestrator must not merge the default ACL into an inherit-only folder'

        # AdNames is only added when a folder had ACL entries to rewrite.
        $inheritOnly.PSObject.Properties.Match('AdNames').Count | Should-Be 0 -Because 'no AD objects were resolved for an empty folder'
        #endregion

        #region The permissioned folder DID receive the default (contrast)
        $finance = $matrix | Where-Object { $_.Path -eq 'Finance' }

        $finance | Should-BeTruthy
        $finance.ACL.Keys | Should-ContainCollection $sidDefault `
            -Because 'ApplyDefaultPermissions=TRUE merges the default into folders that have permissions'
        $finance.AdNames[$sidDefault] | Should-Be 'DefaultGroup' -Because 'the merged default is tracked in the AdNames map'
        #endregion
    }

    It 'builds a Parent (root) matrix entry from the Path row so top-level folders are checked' {
        # Regression: the orchestrator must turn the Permissions 'Path' row
        # (row index 3, the root folder's own permissions) into a matrix entry
        # flagged Parent=$true. SetPermissions.ps1 seeds its inheritance walk
        # from folders that HAVE an ACL, so without this root entry no walk
        # starts at the share root and every TOP-LEVEL inherit-only folder
        # (e.g. 'Common') is never visited or corrected. This test would fail
        # before the fix because no Parent entry was ever produced.
        $matrixDir = (New-Item (Join-Path $TestDrive 'ParentMatrix') -ItemType Directory -Force).FullName
        $logsDir = (New-Item (Join-Path $TestDrive 'ParentLogs') -ItemType Directory -Force).FullName

        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'WithEmptyPermissionFolder') `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'HQ'
                SiteCode                = 'HQ'
                ComputerName            = $env:COMPUTERNAME
                Path                    = (Join-Path $TestDrive 'Target')
                GroupName               = 'Team-A'
                Action                  = 'Fix'
                ApplyDefaultPermissions = $true
            }
        ) | Out-Null

        $configFixture = New-JsonFixture
        $configFixture.Matrix.FolderPath = $matrixDir
        $configFixture.Matrix.DefaultsFile = $defaultsPath
        $configFixture.Settings.SaveLogFiles.Where.Folder = $logsDir
        $configFixture.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        Save-TestJson -InputObject $configFixture -JsonFile $configPath

        $scriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
            SetPermissions         = "$root\Scripts\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Scripts\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Scripts\Operations\UpdateServiceNow.ps1"
        }

        Mock Get-ADObjectDetailHC {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = $sidBob } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = $sidMike } }
                @{ SamAccountName = 'DefaultGroup'; adObject = @{ ObjectSid = $sidDefault } }
            )
        }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        $context = Invoke-PermissionMatrixBeginHC `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrors)

        $systemErrors.Where({ $_.Type -eq 'FatalError' }).Count | Should-Be 0
        $matrix = $context.AllMatrices[0].Matrix

        #region A Parent=$true root entry exists, carrying the Path-row ACL
        $parent = $matrix | Where-Object {
            $_.PSObject.Properties.Match('Parent').Count -and $_.Parent
        }

        $parent | Should-BeTruthy `
            -Because 'the Path row must become a Parent=$true entry that seeds the root inheritance walk'
        @($parent).Count | Should-Be 1 -Because 'there is exactly one root folder'
        @($parent.AdNames.Values) | Should-ContainCollection 'Bob' `
            -Because 'the Path row grants List to the first header group'
        @($parent.AdNames.Values) | Should-ContainCollection 'Mike' `
            -Because 'the Path row grants List to the second header group'
        @($parent.AdNames.Values) | Should-ContainCollection 'DefaultGroup' `
            -Because 'the root has permissions, so ApplyDefaultPermissions merges the default into it too'
        #endregion

        #region The top-level inherit-only folder is still empty (unchanged)
        $inheritOnly = $matrix | Where-Object { $_.Path -eq 'InheritOnly' }
        $inheritOnly.ACL.Count | Should-Be 0 `
            -Because 'the root entry seeds the walk but must not add an ACL to the empty folder'
        #endregion
    }
}