#Requires -Module @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
#Requires -Module ImportExcel

<#
    Round-trip tests for Test-MatrixPermissionsHC.

    Strategy
    --------
    The repo fixtures describe a Permissions sheet as a spec hashtable
    (Row1..Row4 + Data) and write it to a real .xlsx via
    New-MatrixPermissionsExcelFixture. Production reads that sheet back with
    Import-Excel -NoHeader -DataOnly, which yields PSCustomObjects whose
    columns are named P1, P2, P3 ... (NOT the writer's internal Column1..N).

    To test the function against the exact object shape it sees in production,
    each test:
      1. builds the spec from New-MatrixPermissionsFixtureRows -Scenario '...'
      2. writes it to TestDrive: with New-MatrixPermissionsExcelFixture
      3. reads it back with Import-Excel -NoHeader -DataOnly
      4. passes the result to Test-MatrixPermissionsHC and asserts the check

    This requires the ImportExcel module and a writable TestDrive, so these are
    integration-flavoured rather than pure unit tests. They will NOT run in an
    environment without PowerShell + ImportExcel.

    NOTE: paths below assume this file sits under Tests\Unit\Private\. Adjust
    $repoRoot / the dot-source paths if relocated.
#>

BeforeDiscovery {
    . "$PSScriptRoot\Helpers\Fixtures.Matrix.ps1"

    $script:PermissionFixtures = Get-MatrixPermissionsFixtures
}

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    . "$moduleRoot\Private\Utils.ps1"

    . "$root/Tests/Helpers/Fixtures.Excel.ps1"
    . "$root/Tests/Helpers/Fixtures.Matrix.ps1"

    if (-not (Get-Command Test-MatrixPermissionsHC -ErrorAction Ignore)) {
        . "$moduleRoot/Private/Validation.ps1"
    }
    if (-not (Get-Command Test-MatrixPermissionsHC -ErrorAction Ignore)) {
        throw 'Test-MatrixPermissionsHC not loaded. Fix the dot-source path in BeforeAll.'
    }

    # Build a Permissions sheet from a scenario, round-trip it through Excel,
    # and return the imported objects exactly as production sees them.
    function Get-RoundTripPermissions {
        param([Parameter(Mandatory)][string]$Scenario)

        $spec = New-MatrixPermissionsFixtureRows -Scenario $Scenario

        $dir = Join-Path 'TestDrive:' 'Matrix'
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        $path = Join-Path $dir "Permissions_$Scenario.xlsx"

        New-MatrixPermissionsExcelFixture -Path $path -Spec $spec | Out-Null

        return @(Import-Excel -Path $path -WorksheetName 'Permissions' -NoHeader -DataOnly -ErrorAction Stop)
    }
}

Describe 'Test-MatrixPermissionsHC' {

    Context 'Happy path' {
        It 'returns nothing when the Valid fixture is supplied' {
            $perms = Get-RoundTripPermissions -Scenario 'Valid'

            $result = Test-MatrixPermissionsHC -Permissions $perms

            # Function only returns $checks when Count -gt 0, so success => $null.
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'Data-driven checks from Get-MatrixPermissionsFixtures' {
        # Each fixture row: @{ Issue; Mutation; Expected }
        # 'Expected' is the check Name the function should emit.
        It 'flags <Issue> with check name <Expected>' -ForEach $PermissionFixtures {

            # The fixture 'Mutation' strings map 1:1 to a scenario name in
            # New-MatrixPermissionsFixtureRows; derive it from the Issue so we
            # can round-trip in-process rather than Invoke-Expression a string
            # that writes its own file.
            $scenario = switch ($Issue) {
                'MissingADObjectName' { 'MissingADObjectName' }
                'InvalidPermissionChar' { 'InvalidPermissionChar' }
                'MissingRows' { 'MissingRows' }
                'MissingColumns' { 'MissingColumns' }
                'MissingFolderName' { 'MissingFolderName' }
                'DuplicateFolderName' { 'DuplicateFolderName' }
                'InaccessibleFolders' { 'InaccessibleFolders' }
                default { throw "No scenario mapping for Issue '$Issue'" }
            }

            $perms = Get-RoundTripPermissions -Scenario $scenario

            $result = Test-MatrixPermissionsHC -Permissions $perms

            $result | Should -Not -BeNullOrEmpty
            ($result.Name) | Should -Contain $Expected
        }
    }

    Context 'fatal errors exit immediately' {
        It 'returns ONLY "Missing rows" for the MissingRows fixture' {
            $perms = Get-RoundTripPermissions -Scenario 'MissingRows'

            $result = Test-MatrixPermissionsHC -Permissions $perms

            @($result).Count | Should -Be 1
            $result[0].Type | Should -Be 'FatalError'
            $result[0].Name | Should -Be 'Missing rows'
        }

        It 'returns ONLY "Missing columns" for the MissingColumns fixture' {
            $perms = Get-RoundTripPermissions -Scenario 'MissingColumns'

            $result = Test-MatrixPermissionsHC -Permissions $perms

            @($result).Count | Should -Be 1
            $result[0].Type | Should -Be 'FatalError'
            $result[0].Name | Should -Be 'Missing columns'
        }
    }

    Context 'Check types are correct' {
        It 'classifies InaccessibleFolders as a Warning, not a FatalError' {
            $perms = Get-RoundTripPermissions -Scenario 'InaccessibleFolders'
            $result = Test-MatrixPermissionsHC -Permissions $perms

            $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
            $warn | Should -Not -BeNullOrEmpty
            $warn.Type | Should -Be 'Warning'
        }

        It 'classifies InvalidPermissionChar as a FatalError' {
            $perms = Get-RoundTripPermissions -Scenario 'InvalidPermissionChar'
            $result = Test-MatrixPermissionsHC -Permissions $perms

            $err = $result | Where-Object Name -EQ 'Invalid permission character'
            $err | Should -Not -BeNullOrEmpty
            $err.Type | Should -Be 'FatalError'
        }
    }

    Context 'Error handling' {
        It 'throws a descriptive error when handed an empty array' {
            { Test-MatrixPermissionsHC -Permissions @() } |
            Should -Throw -ExpectedMessage "*Failed testing the Excel sheet 'Permissions'*"
        }
    }
}
