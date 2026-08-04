#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
    ForEach-Object { . $_.FullName }
}

Describe 'Set-DefaultSheetFilterHC' {
    BeforeAll {
        function New-TestSheet {
            <#
                One enabled member, one disabled member and one group
                without members, so every branch of the filter is covered.
                'MemberEnabled' is column 6.
            #>
            param(
                [string]$Path,
                [string]$WorksheetName = 'AccessList'
            )

            @(
                [pscustomobject]@{
                    SamAccountName = 'grp'; Name = 'grp'; Type = 'group'
                    MemberName = 'John'; MemberSamAccountName = 'jdoe'
                    MemberEnabled = $true
                }
                [pscustomobject]@{
                    SamAccountName = 'grp'; Name = 'grp'; Type = 'group'
                    MemberName = 'Bob'; MemberSamAccountName = 'bsmith'
                    MemberEnabled = $false
                }
                [pscustomobject]@{
                    SamAccountName = 'empty'; Name = 'empty'; Type = 'group'
                    MemberName = $null; MemberSamAccountName = $null
                    MemberEnabled = $null
                }
            ) | Export-Excel -Path $Path -WorksheetName $WorksheetName `
                -TableName $WorksheetName -FreezeTopRow

            return $Path
        }

        function Get-HiddenState {
            # Hidden state of the data rows, top to bottom
            param([string]$Path, [string]$WorksheetName)

            $package = Open-ExcelPackage -Path $Path
            try {
                $sheet = $package.Workbook.Worksheets[$WorksheetName]
                return @(
                    for ($r = 2; $r -le $sheet.Dimension.End.Row; $r++) {
                        $sheet.Row($r).Hidden
                    }
                )
            }
            finally {
                Close-ExcelPackage -ExcelPackage $package -NoSave
            }
        }

        function Get-TableXml {
            param([string]$Path, [string]$WorksheetName)

            $package = Open-ExcelPackage -Path $Path
            try {
                $sheet = $package.Workbook.Worksheets[$WorksheetName]
                if (-not $sheet.Tables.Count) { return $null }
                return $sheet.Tables[0].TableXml.OuterXml
            }
            finally {
                Close-ExcelPackage -ExcelPackage $package -NoSave
            }
        }
    }

    Context 'hiding rows' {
        It 'hides the rows that do not match the visible value' {
            $path = New-TestSheet -Path (Join-Path $TestDrive 'hide.xlsx')

            Set-DefaultSheetFilterHC -Path $path `
                -WorksheetName 'AccessList' `
                -ColumnName 'MemberEnabled' -VisibleValue 'TRUE'

            Get-HiddenState -Path $path -WorksheetName 'AccessList' | Should-BeCollection @($false, $true, $true)
        }

        It 'keeps blank cells visible with -IncludeBlank' {
            $path = New-TestSheet -Path (Join-Path $TestDrive 'blank.xlsx')

            Set-DefaultSheetFilterHC -Path $path `
                -WorksheetName 'AccessList' `
                -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank

            Get-HiddenState -Path $path -WorksheetName 'AccessList' | Should-BeCollection @($false, $true, $false)
        }

        It 'matches a boolean cell against the string value' {
            # AD returns a real boolean, which EPPlus stores as a boolean
            # cell rather than the text 'TRUE'
            $path = New-TestSheet -Path (Join-Path $TestDrive 'bool.xlsx')

            $package = Open-ExcelPackage -Path $path
            try {
                $sheet = $package.Workbook.Worksheets['AccessList']
                $sheet.Cells[2, 6].Value | Should-HaveType ([bool])
            }
            finally {
                Close-ExcelPackage -ExcelPackage $package -NoSave
            }

            Set-DefaultSheetFilterHC -Path $path `
                -WorksheetName 'AccessList' `
                -ColumnName 'MemberEnabled' -VisibleValue 'TRUE'

            (Get-HiddenState -Path $path -WorksheetName 'AccessList')[0] | Should-BeFalse
        }

        It 'leaves the row data intact' {
            $path = New-TestSheet -Path (Join-Path $TestDrive 'intact.xlsx')

            Set-DefaultSheetFilterHC -Path $path `
                -WorksheetName 'AccessList' `
                -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank

            @(Import-Excel -Path $path -WorksheetName 'AccessList').Count | Should-Be 3
        }
    }

    Context 'filter definition' {
        It 'writes the criterion into the table autoFilter' {
            $path = New-TestSheet -Path (Join-Path $TestDrive 'criterion.xlsx')

            Set-DefaultSheetFilterHC -Path $path `
                -WorksheetName 'AccessList' `
                -ColumnName 'MemberEnabled' -VisibleValue 'TRUE'

            $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

            $xml | Should-MatchString '<filterColumn colId="5"'
            $xml | Should-MatchString '<filter val="TRUE"'
        }

        It 'sets the blank attribute only with -IncludeBlank' {
            $with = New-TestSheet -Path (Join-Path $TestDrive 'attr-with.xlsx')
            $without = New-TestSheet -Path (Join-Path $TestDrive 'attr-without.xlsx')

            Set-DefaultSheetFilterHC -Path $with `
                -WorksheetName 'AccessList' `
                -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank

            Set-DefaultSheetFilterHC -Path $without `
                -WorksheetName 'AccessList' `
                -ColumnName 'MemberEnabled' -VisibleValue 'TRUE'

            Get-TableXml -Path $with -WorksheetName 'AccessList' | Should-MatchString 'blank="1"'

            Get-TableXml -Path $without -WorksheetName 'AccessList' | Should-NotMatchString 'blank='
        }

        It 'does not stack criteria when applied twice' {
            # A second filterColumn for the same column makes Excel
            # report the workbook as corrupt
            $path = New-TestSheet -Path (Join-Path $TestDrive 'twice.xlsx')

            1..2 | ForEach-Object {
                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank
            }

            $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

            ([regex]::Matches($xml, '<filterColumn')).Count | Should-Be 1
        }
    }

    Context 'sheets that cannot be filtered' {
        It 'skips a worksheet that does not exist' {
            $path = New-TestSheet -Path (Join-Path $TestDrive 'missing-sheet.xlsx')

            {
                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList', 'GroupManagers' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE'
            } | Should -Not -Throw

            (Get-HiddenState -Path $path -WorksheetName 'AccessList')[1] | Should-BeTrue
        }

        It 'skips a worksheet without the requested column' {
            $path = New-TestSheet -Path (Join-Path $TestDrive 'missing-col.xlsx')

            {
                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'DoesNotExist' -VisibleValue 'TRUE'
            } | Should -Not -Throw

            Get-HiddenState -Path $path -WorksheetName 'AccessList' | Should-BeCollection @($false, $false, $false)
        }

        It 'skips a header-only worksheet' {
            # Export-*HC writes bare headers when there are no rows
            $path = Join-Path $TestDrive 'header-only.xlsx'

            $package = Open-ExcelPackage -Path $path -Create
            try {
                $sheet = Add-Worksheet -ExcelPackage $package `
                    -WorksheetName 'AccessList'
                $sheet.Cells[1, 1].Value = 'SamAccountName'
                $sheet.Cells[1, 6].Value = 'MemberEnabled'
            }
            finally {
                Close-ExcelPackage -ExcelPackage $package
            }

            {
                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE'
            } | Should -Not -Throw
        }
    }

    Context 'error handling' {
        It 'throws a descriptive error when the file cannot be opened' {
            $path = Join-Path $TestDrive 'does-not-exist.xlsx'

            {
                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' -ColumnName 'MemberEnabled'
            } | Should-Throw -ExceptionMessage "*column 'MemberEnabled'*"
        }
    }
}
