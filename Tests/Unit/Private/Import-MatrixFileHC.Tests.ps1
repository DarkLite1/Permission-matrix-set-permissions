#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }, ImportExcel

<#
    Tests for Modules\PermissionMatrix\Private\Import-MatrixFileHC.ps1

    Approach:
        - Real ImportExcel: matrix workbooks are built with the project's own
          Excel fixtures (Tests\Helpers\Fixtures.Excel.ps1) and read back by the
          function's Get-ExcelWorkbookInfo / Import-Excel.
        - Real collaborators: the whole Private folder is dot-sourced so
          Format-SettingStringsHC, Format-PermissionsStringsHC,
          Format-FormDataStringsHC and Test-FormDataHC are the real ones.
        - The function takes a [System.IO.FileInfo] and uses .FullName, which
          does not resolve the TestDrive: PSDrive, so fixtures are written to the
          real $TestDrive path and passed via Get-Item.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    # Dot-source every Private function so Import-MatrixFileHC and its real
    # collaborators are all defined alongside each other.
    Get-ChildItem "$moduleRoot\Private\*.ps1" | ForEach-Object { . $_.FullName }

    Import-Module ImportExcel -ErrorAction Stop

    # Project Excel fixtures (New-MatrixExcelFixture, New-MatrixSettingsFixtureRows,
    # New-MatrixPermissionsExcelFixture, New-MatrixPermissionsFixtureRows, ...).
    . "$root\Tests\Helpers\Fixtures.Excel.ps1"

    function New-TestContext {
        param(
            [string]$ServiceNowFormDataExcelFile,
            [string]$OverviewHtmlFile
        )
        [pscustomobject]@{
            Config = [pscustomobject]@{
                Export = [pscustomobject]@{
                    ServiceNowFormDataExcelFile = $ServiceNowFormDataExcelFile
                    OverviewHtmlFile            = $OverviewHtmlFile
                }
            }
        }
    }

    function New-EnabledSettingsRows {
        param([int]$Count = 1)
        1..$Count | ForEach-Object {
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'HQ South'
                SiteCode                = 'CS&L'
                ComputerName            = "BEL`$FFRAN000$_"
                Path                    = "E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L$_"
                GroupName               = 'BEL ROL-AGS-SAGREV'
                Action                  = 'Fix'
                ApplyDefaultPermissions = $false
            }
        }
    }
}

Describe 'New-MatrixFileResultHC' {
    BeforeAll {
        $matrixFile = Get-Item (New-Item "$TestDrive\shape.xlsx" -ItemType File -Force).FullName
    }

    It 'exposes the matrix file on Item' {
        $result = New-MatrixFileResultHC -MatrixFile $matrixFile

        $result.Item.FullName | Should-Be $matrixFile.FullName
        $result.Item.BaseName | Should-Be 'shape'
    }

    It 'initializes every check collection so entries can be added' {
        # A null collection makes .Add() throw, which fails the test.
        $result = New-MatrixFileResultHC -MatrixFile $matrixFile

        $result.Check.Add([pscustomobject]@{ Type = 'Warning' })
        $result.Matrices.Add([pscustomobject]@{ ID = 1 })
        $result.Sheets.Permissions.Check.Add([pscustomobject]@{ Type = 'Warning' })
        $result.Sheets.FormData.Check.Add([pscustomobject]@{ Type = 'Warning' })

        $result.Check.Count | Should-Be 1
        $result.Matrices.Count | Should-Be 1
        $result.Sheets.Permissions.Check.Count | Should-Be 1
        $result.Sheets.FormData.Check.Count | Should-Be 1
    }

    It 'allows the reporting stage to assign LogFolder and ReportFilePath' {
        <# Regression: assigning a property that does not exist on a
        [PSCustomObject] throws, which fails the test. Invoke-PermissionMatrixEndHC
        writes to both of these, and its per-file loop is a single try block, so
        an incomplete result aborts logging for every remaining file in the run. #>
        $result = New-MatrixFileResultHC -MatrixFile $matrixFile

        $result.ReportFileName | Should-Be '00 - Execution Report.html'

        $result.LogFolder = 'C:\logs\shape'
        $result.ReportFilePath = 'C:\logs\shape\report.html'

        $result.LogFolder | Should-Be 'C:\logs\shape'
        $result.ReportFilePath | Should-Be 'C:\logs\shape\report.html'
    }
}

Describe 'Import-MatrixFileHC' {
    BeforeEach {
        Remove-Item (Join-Path $TestDrive '*') -Recurse -Force -ErrorAction Ignore
        $matrixPath = Join-Path $TestDrive 'matrix.xlsx'
    }

    Context 'a valid workbook with enabled settings' {
        It 'returns a populated file result with one matrix for one enabled setting' {
            New-MatrixExcelFixture -Path $matrixPath

            $result = Import-MatrixFileHC `
                -MatrixFile (Get-Item -LiteralPath $matrixPath) `
                -Context (New-TestContext)

            $result.Check | Should-BeFalsy
            $result.ExcelInfo | Should-BeTruthy
            $result.Sheets.Settings.Raw | Should-BeTruthy
            $result.Sheets.Settings.Formatted | Should-BeTruthy
            $result.Sheets.Permissions.Raw | Should-BeTruthy
            $result.Sheets.Permissions.Formatted | Should-BeTruthy
            $result.Matrices | Should-BeCollection -Count 1
            $result.Matrices[0].Setting.Raw.Status | Should-Be 'Enabled'
            $result.Matrices[0].Setting.Formatted | Should-BeTruthy
            $result.Matrices[0].ID | Should-BeTruthy
        }

        It 'sets each matrix FileContext back to the returned file result' {
            New-MatrixExcelFixture -Path $matrixPath

            $result = Import-MatrixFileHC `
                -MatrixFile (Get-Item -LiteralPath $matrixPath) `
                -Context (New-TestContext)

            $result.Matrices[0].FileContext | Should-Be $result
        }

        It 'creates one matrix per enabled setting when several are enabled' {
            New-MatrixExcelFixture -Path $matrixPath -SettingsRows (New-EnabledSettingsRows -Count 3)

            $result = Import-MatrixFileHC `
                -MatrixFile (Get-Item -LiteralPath $matrixPath) `
                -Context (New-TestContext)

            $result.Matrices | Should-BeCollection -Count 3
            $result.Check | Should-BeFalsy
        }
    }

    Context 'Status with surrounding whitespace' {
        It 'treats a Status of " Enabled " as enabled' {
            <# Regression: the enabled rows were selected from the raw sheet,
            before Format-SettingStringsHC trimmed anything, so a stray space
            in the Excel cell silently disabled the entire matrix file and
            reported only 'No enabled matrix settings'. #>
            $rows = New-MatrixSettingsFixtureRows -Scenario 'Valid'

            foreach ($row in $rows) {
                $row.Status = ' Enabled '
            }

            New-MatrixExcelFixture -Path $matrixPath -SettingsRows $rows

            $result = Import-MatrixFileHC `
                -MatrixFile (Get-Item -LiteralPath $matrixPath) `
                -Context (New-TestContext)

            $result.Check | Should-BeFalsy
            $result.Matrices | Should-BeTruthy
        }
    }

    Context 'no enabled settings' {
        It 'records a FatalError, reads no further, and creates no matrices' {
            New-MatrixExcelFixture -Path $matrixPath -Disabled

            $result = Import-MatrixFileHC `
                -MatrixFile (Get-Item -LiteralPath $matrixPath) `
                -Context (New-TestContext)

            $result.Matrices | Should-BeFalsy
            $result.Check | Should-BeCollection -Count 1
            $result.Check[0].Type | Should-Be 'FatalError'
            $result.Check[0].Name | Should-Be 'No enabled matrix settings'

            # Settings were read; the early return happens before Permissions.
            $result.Sheets.Settings.Raw | Should-BeTruthy
            $result.Sheets.Permissions.Raw | Should-BeFalsy
        }
    }

    Context 'optional FormData' {
        It 'skips FormData when no export is configured' {
            New-MatrixExcelFixture -Path $matrixPath

            $result = Import-MatrixFileHC `
                -MatrixFile (Get-Item -LiteralPath $matrixPath) `
                -Context (New-TestContext)   # both export paths unset

            $result.Sheets.FormData.Raw | Should-BeFalsy
            $result.Sheets.FormData.Formatted | Should-BeFalsy
        }

        It 'imports and formats FormData when an export is configured' {
            New-MatrixExcelFixture -Path $matrixPath

            $result = Import-MatrixFileHC `
                -MatrixFile (Get-Item -LiteralPath $matrixPath) `
                -Context (New-TestContext -OverviewHtmlFile 'C:\reports\Overview.html')

            $result.Sheets.FormData.Raw | Should-BeTruthy
            # No "missing sheet" FatalError, because the sheet exists.
            ($result.Sheets.FormData.Check | Where-Object Name -EQ "Worksheet 'FormData' not found") |
                Should-BeFalsy
            # Assumes the fixture's default FormData passes the real Test-FormDataHC.
            $result.Sheets.FormData.Formatted | Should-BeTruthy
            $result.Matrices | Should-BeCollection -Count 1
        }

        It 'records a FatalError when the FormData sheet is missing but an export is configured' {
            # Settings + Permissions only — no FormData sheet.
            New-EnabledSettingsRows -Count 1 |
                Export-Excel -Path $matrixPath -WorksheetName 'Settings' -TableName 'Settings' `
                    -ClearSheet -AutoSize -FreezeTopRow
            New-MatrixPermissionsExcelFixture `
                -Path $matrixPath `
                -Spec (New-MatrixPermissionsFixtureRows -Scenario 'Valid') | Out-Null

            $result = Import-MatrixFileHC `
                -MatrixFile (Get-Item -LiteralPath $matrixPath) `
                -Context (New-TestContext -ServiceNowFormDataExcelFile 'C:\snow\FormData.xlsx')

            $formDataCheck = $result.Sheets.FormData.Check | Where-Object Name -EQ "Worksheet 'FormData' not found"
            $formDataCheck | Should-BeTruthy
            $formDataCheck.Type | Should-Be 'FatalError'

            # The missing FormData sheet is handled in its own catch, so the rest
            # of the import continues and the matrix is still created.
            $result.Matrices | Should-BeCollection -Count 1
        }

        It 'records a FatalError when Test-FormDataHC flags the FormData' -Skip {
            # Reliably exercising this branch with the real Test-FormDataHC needs
            # its validation rules, so the fixture FormData can be made invalid in
            # a way it actually rejects. Share Test-FormDataHC (or the shape of
            # FormData it rejects) and this can be turned into a real assertion.
        }
    }

    Context 'an unreadable workbook' {
        It 'records a catch-all FatalError when a mandatory sheet cannot be read' {
            # Workbook with only a Permissions sheet — importing 'Settings' throws.
            New-MatrixPermissionsExcelFixture `
                -Path $matrixPath `
                -Spec (New-MatrixPermissionsFixtureRows -Scenario 'Valid') | Out-Null

            $result = Import-MatrixFileHC `
                -MatrixFile (Get-Item -LiteralPath $matrixPath) `
                -Context (New-TestContext)

            $result.Check | Should-BeCollection -Count 1
            $result.Check[0].Type | Should-Be 'FatalError'
            $result.Check[0].Name | Should-Be 'Excel file incorrect'
            $result.Matrices | Should-BeFalsy
        }
    }
}