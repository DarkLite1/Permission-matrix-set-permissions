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

    Context 'excluding placeholder accounts' {
        BeforeAll {
            function New-PlaceHolderSheet {
                <#
                    An enabled member, the placeholder account, a disabled
                    member and a group without members, so the exclude
                    criterion, the include criterion and the blank branch are
                    all exercised on one sheet.
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
                        MemberName = 'Chuck Norris'
                        MemberSamAccountName = 'cnorris'
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

            function New-GroupManagersSheet {
                <#
                    'GroupManagers' has no SamAccountName column: the
                    placeholder can only be matched on its display name in
                    'ManagerMemberName'.
                #>
                param([string]$Path)

                @(
                    [pscustomobject]@{
                        GroupName = 'grp'; ManagerName = 'MgrGrp'
                        ManagerType = 'group'; ManagerMemberName = 'John'
                        MemberEnabled = $true
                    }
                    [pscustomobject]@{
                        GroupName = 'grp'; ManagerName = 'MgrGrp'
                        ManagerType = 'group'
                        ManagerMemberName = 'Chuck Norris'
                        MemberEnabled = $true
                    }
                    [pscustomobject]@{
                        GroupName = 'grp2'; ManagerName = $null
                        ManagerType = $null; ManagerMemberName = $null
                        MemberEnabled = $null
                    }
                ) | Export-Excel -Path $Path -WorksheetName 'GroupManagers' `
                    -TableName 'GroupManagers' -FreezeTopRow

                return $Path
            }
        }

        Context 'hiding rows' {
            It 'hides the placeholder row on top of the disabled rows' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-hide.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris'

                # John visible, Chuck excluded, Bob disabled, empty group blank
                Get-HiddenState -Path $path -WorksheetName 'AccessList' |
                Should-BeCollection @($false, $true, $true, $false)
            }

            It 'hides an enabled placeholder, so the criteria combine with AND' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-and.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris'

                # Row 3 passes MemberEnabled but must still be hidden
                (Get-HiddenState -Path $path -WorksheetName 'AccessList')[1] |
                Should-BeTrue
            }

            It 'matches the placeholder case insensitively' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-case.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'CNORRIS'

                (Get-HiddenState -Path $path -WorksheetName 'AccessList')[1] |
                Should-BeTrue
            }

            It 'keeps a group without members visible with -IncludeBlank' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-blank.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris'

                (Get-HiddenState -Path $path -WorksheetName 'AccessList')[3] |
                Should-BeFalse
            }

            It 'honours several placeholder accounts at once' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-many.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris', 'jdoe'

                Get-HiddenState -Path $path -WorksheetName 'AccessList' |
                Should-BeCollection @($true, $true, $true, $false)
            }

            It 'leaves the row data intact' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-intact.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris'

                @(Import-Excel -Path $path -WorksheetName 'AccessList').Count |
                Should-Be 4
            }
        }

        Context 'filter definition' {
            It 'writes an inclusion list of every other value in the column' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-xml.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris'

                $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

                $xml | Should-MatchString '<filterColumn colId="4"'
                $xml | Should-MatchString '<filter val="jdoe"'
                $xml | Should-MatchString '<filter val="bsmith"'
                $xml | Should-NotMatchString '<filter val="cnorris"'
            }

            It 'keeps the MemberEnabled criterion alongside it' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-both.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris'

                $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

                ([regex]::Matches($xml, '<filterColumn')).Count | Should-Be 2
                $xml | Should-MatchString '<filterColumn colId="5"'
                $xml | Should-MatchString '<filter val="TRUE"'
            }

            It 'writes the criteria in ascending colId order' {
                # CT_AutoFilter expects the filterColumn children in sequence
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-order.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris'

                $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

                ($xml.IndexOf('colId="4"') -lt $xml.IndexOf('colId="5"')) |
                Should-BeTrue
            }

            It 'does not stack criteria when applied twice' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-twice.xlsx')

                1..2 | ForEach-Object {
                    Set-DefaultSheetFilterHC -Path $path `
                        -WorksheetName 'AccessList' `
                        -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                        -ExcludeColumnName 'MemberSamAccountName' `
                        -ExcludeValue 'cnorris'
                }

                $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

                ([regex]::Matches($xml, '<filterColumn')).Count | Should-Be 2
            }

            It 'still orders the criteria after a second run' {
                # Re-applying removes a criterion and appends the new one at
                # the end, which is what the reordering pass exists for
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-reorder.xlsx')

                1..2 | ForEach-Object {
                    Set-DefaultSheetFilterHC -Path $path `
                        -WorksheetName 'AccessList' `
                        -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                        -ExcludeColumnName 'MemberSamAccountName' `
                        -ExcludeValue 'cnorris'
                }

                $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

                ($xml.IndexOf('colId="4"') -lt $xml.IndexOf('colId="5"')) |
                Should-BeTrue
            }
        }

        Context 'nothing to exclude' {
            It 'behaves exactly as before when no ExcludeValue is passed' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-none.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName'

                $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

                ([regex]::Matches($xml, '<filterColumn')).Count | Should-Be 1
                Get-HiddenState -Path $path -WorksheetName 'AccessList' |
                Should-BeCollection @($false, $false, $true, $false)
            }

            It 'ignores an empty ExcludeValue array' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-empty.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue @()

                ([regex]::Matches(
                    (Get-TableXml -Path $path -WorksheetName 'AccessList'),
                    '<filterColumn'
                )).Count | Should-Be 1
            }

            It 'skips an exclude column the worksheet does not have' {
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-nocol.xlsx')

                {
                    Set-DefaultSheetFilterHC -Path $path `
                        -WorksheetName 'AccessList' `
                        -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                        -ExcludeColumnName 'ManagerMemberName' `
                        -ExcludeValue 'Chuck Norris'
                } | Should -Not -Throw

                ([regex]::Matches(
                    (Get-TableXml -Path $path -WorksheetName 'AccessList'),
                    '<filterColumn'
                )).Count | Should-Be 1
            }
        }

        Context 'every value excluded' {
            BeforeAll {
                function New-AllPlaceHolderSheet {
                    param([string]$Path)

                    @(
                        [pscustomobject]@{
                            SamAccountName = 'grp'; Name = 'grp'; Type = 'group'
                            MemberName = 'Chuck Norris'
                            MemberSamAccountName = 'cnorris'
                            MemberEnabled = $true
                        }
                        [pscustomobject]@{
                            SamAccountName = 'empty'; Name = 'empty'
                            Type = 'group'; MemberName = $null
                            MemberSamAccountName = $null; MemberEnabled = $null
                        }
                    ) | Export-Excel -Path $Path -WorksheetName 'AccessList' `
                        -TableName 'AccessList' -FreezeTopRow

                    return $Path
                }
            }

            It 'still writes the criterion when blanks stay visible' {
                $path = New-AllPlaceHolderSheet -Path (Join-Path $TestDrive 'ph-all-blank.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris'

                $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

                $xml | Should-MatchString '<filterColumn colId="4"'
                $xml | Should-MatchString 'blank="1"'

                Get-HiddenState -Path $path -WorksheetName 'AccessList' |
                Should-BeCollection @($true, $false)
            }

            It 'omits the criterion when blanks are hidden too' {
                # An empty 'filters' element makes Excel report the workbook
                # as corrupt
                $path = New-AllPlaceHolderSheet -Path (Join-Path $TestDrive 'ph-all-noblank.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris'

                $xml = Get-TableXml -Path $path -WorksheetName 'AccessList'

                $xml | Should-NotMatchString '<filterColumn colId="4"'

                # The rows are hidden regardless
                Get-HiddenState -Path $path -WorksheetName 'AccessList' |
                Should-BeCollection @($true, $true)
            }
        }

        Context 'worksheets with different layouts' {
            It 'matches the display name on GroupManagers' {
                $path = New-GroupManagersSheet -Path (Join-Path $TestDrive 'gm-name.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'GroupManagers' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName', 'ManagerMemberName' `
                    -ExcludeValue 'cnorris', 'Chuck Norris'

                Get-HiddenState -Path $path -WorksheetName 'GroupManagers' |
                Should-BeCollection @($false, $true, $false)
            }

            It 'covers both worksheets in a single call' {
                $path = Join-Path $TestDrive 'both-sheets.xlsx'

                $null = New-PlaceHolderSheet -Path $path
                $null = New-GroupManagersSheet -Path $path

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList', 'GroupManagers' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName', 'ManagerMemberName' `
                    -ExcludeValue 'cnorris', 'Chuck Norris'

                Get-HiddenState -Path $path -WorksheetName 'AccessList' |
                Should-BeCollection @($false, $true, $true, $false)

                Get-HiddenState -Path $path -WorksheetName 'GroupManagers' |
                Should-BeCollection @($false, $true, $false)
            }

            It 'ignores a value that occurs in neither column' {
                # The SamAccountName and the display name are passed together,
                # so each sheet always receives values it cannot match
                $path = New-PlaceHolderSheet -Path (Join-Path $TestDrive 'ph-unmatched.xlsx')

                Set-DefaultSheetFilterHC -Path $path `
                    -WorksheetName 'AccessList' `
                    -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
                    -ExcludeColumnName 'MemberSamAccountName' `
                    -ExcludeValue 'cnorris', 'Chuck Norris', 'nobody'

                Get-HiddenState -Path $path -WorksheetName 'AccessList' |
                Should-BeCollection @($false, $true, $true, $false)
            }
        }
    }
}