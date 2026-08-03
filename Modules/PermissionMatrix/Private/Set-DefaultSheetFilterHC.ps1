function Set-DefaultSheetFilterHC {
    <#
    .SYNOPSIS
        Pre-applies a filter on one or more worksheets of an existing workbook.

    .DESCRIPTION
        Excel stores the filter *definition* and the *hidden row state* as two
        separate things and does not re-evaluate a filter when a file is
        opened. Writing only the filter definition therefore shows the funnel
        icon on the header but leaves every row visible.

        This function does both:
        - writes an 'autoFilter/filterColumn' criterion into the Excel Table
          XML, so the funnel icon is shown and 'Clear Filter' works normally
        - marks the non-matching rows as hidden, so the sheet actually opens
          filtered

        The worksheets are expected to be written by Export-Excel with
        '-TableName', which is why the criterion goes into the table part
        (/xl/tables/tableN.xml) and not into the worksheet's own autoFilter.

    .PARAMETER Path
        Path to the .xlsx file to update.

    .PARAMETER WorksheetName
        One or more worksheets to filter. Worksheets that do not exist, are
        empty, or hold no table are silently skipped.

    .PARAMETER ColumnName
        Header text of the column to filter on, matched on row 1.

    .PARAMETER VisibleValue
        The cell values that stay visible. Comparison is case insensitive.
        A boolean $true cell reads back as 'TRUE'.

    .PARAMETER IncludeBlank
        Keep rows with an empty cell visible as well. For 'MemberEnabled'
        these are the groups without members and the manager rows that hold
        no account status, which are usually worth keeping in view.

    .EXAMPLE
        Set-DefaultSheetFilterHC -Path $Path `
            -WorksheetName 'AccessList', 'GroupManagers' `
            -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank

        Opens the workbook with the disabled accounts hidden, while empty
        'MemberEnabled' cells stay visible.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$WorksheetName,
        [Parameter(Mandatory)][string]$ColumnName,
        [string[]]$VisibleValue = @('TRUE'),
        [switch]$IncludeBlank
    )

    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    try {
        $excelPackage = Open-ExcelPackage -Path $Path -ErrorAction Stop

        try {
            foreach ($sheetName in $WorksheetName) {
                $sheet = $excelPackage.Workbook.Worksheets[$sheetName]

                if ((-not $sheet) -or (-not $sheet.Dimension)) { continue }
                if ($sheet.Dimension.End.Row -le 1) { continue }

                #region Locate the column by its header text
                $columnIndex = $null

                for (
                    $c = $sheet.Dimension.Start.Column
                    $c -le $sheet.Dimension.End.Column
                    $c++
                ) {
                    if ($sheet.Cells[1, $c].Text -eq $ColumnName) {
                        $columnIndex = $c
                        break
                    }
                }

                if (-not $columnIndex) { continue }
                #endregion

                #region Hide the rows that do not match
                for ($r = 2; $r -le $sheet.Dimension.End.Row; $r++) {
                    $cellValue = $sheet.Cells[$r, $columnIndex].Value

                    $isVisible = if (
                        ($null -eq $cellValue) -or
                        ([string]$cellValue).Trim() -eq ''
                    ) {
                        $IncludeBlank.IsPresent
                    }
                    else {
                        $VisibleValue -contains ([string]$cellValue).Trim()
                    }

                    if (-not $isVisible) {
                        $sheet.Row($r).Hidden = $true
                    }
                }
                #endregion

                #region Write the filter criterion into the table XML
                $table = $sheet.Tables | Where-Object {
                    ($_.Address.Start.Column -le $columnIndex) -and
                    ($_.Address.End.Column -ge $columnIndex)
                } | Select-Object -First 1

                if (-not $table) { continue }

                $tableXml = $table.TableXml
                $tableRoot = $tableXml.DocumentElement

                $autoFilter = $tableRoot.SelectSingleNode(
                    "*[local-name()='autoFilter']"
                )

                if (-not $autoFilter) {
                    $autoFilter = $tableXml.CreateElement('autoFilter', $ns)
                    $autoFilter.SetAttribute('ref', $table.Address.Address)

                    # CT_Table requires autoFilter to be the first child
                    $null = $tableRoot.PrependChild($autoFilter)
                }

                $colId = (
                    $columnIndex - $table.Address.Start.Column
                ).ToString()

                # Stay idempotent: a second criterion for the same column
                # makes Excel report the workbook as corrupt
                $existing = $autoFilter.SelectSingleNode(
                    "*[local-name()='filterColumn'][@colId='$colId']"
                )

                if ($existing) {
                    $null = $autoFilter.RemoveChild($existing)
                }

                $filterColumn = $tableXml.CreateElement('filterColumn', $ns)
                $filterColumn.SetAttribute('colId', $colId)

                $filters = $tableXml.CreateElement('filters', $ns)

                if ($IncludeBlank) { $filters.SetAttribute('blank', '1') }

                foreach ($value in $VisibleValue) {
                    $filter = $tableXml.CreateElement('filter', $ns)
                    $filter.SetAttribute('val', $value)
                    $null = $filters.AppendChild($filter)
                }

                $null = $filterColumn.AppendChild($filters)
                $null = $autoFilter.AppendChild($filterColumn)
                #endregion
            }
        }
        finally {
            Close-ExcelPackage -ExcelPackage $excelPackage
        }
    }
    catch {
        throw "Failed applying the default filter on column '$ColumnName' in file '$Path': $_"
    }
}