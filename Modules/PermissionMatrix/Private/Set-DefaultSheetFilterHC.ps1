function Set-DefaultSheetFilterHC {
    <#
    .SYNOPSIS
        Pre-applies one or more filters on the worksheets of an existing
        workbook.

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

        Two kinds of criteria are supported and they combine with AND, exactly
        as Excel treats filters on several columns:

        - An include criterion (ColumnName / VisibleValue): only the listed
          values stay visible. Used for 'MemberEnabled'.
        - An exclude criterion (ExcludeColumnName / ExcludeValue): the listed
          values are hidden and everything else stays visible. Used for the
          placeholder accounts of 'Matrix.AdGroupPlaceHolders'.

        Excel has no native "everything except this list" filter. The exclude
        criterion therefore reproduces what Excel itself writes when boxes are
        unticked by hand: an inclusion list of every *other* distinct value
        found in that column. On a column with many distinct values this
        produces a correspondingly large 'filters' element, which is normal
        but worth knowing.

        Columns that a worksheet does not have are skipped silently, so the
        same call can cover sheets with different layouts.

        The worksheets are expected to be written by Export-Excel with
        '-TableName', which is why the criteria go into the table part
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
        no account status, which are usually worth keeping in view. Applies to
        the exclude criteria too: a group without members has no member name
        to match a placeholder against and should stay in view.

    .PARAMETER ExcludeColumnName
        Header text of one or more columns whose ExcludeValue entries must be
        hidden. Columns absent from a worksheet are skipped, so passing both
        'MemberSamAccountName' and 'ManagerMemberName' covers 'AccessList' and
        'GroupManagers' in a single call.

    .PARAMETER ExcludeValue
        The cell values to hide in every ExcludeColumnName. Comparison is case
        insensitive. Values that occur in none of the columns are harmless, so
        the SamAccountNames and their display names can be passed together.
        When empty, no exclude criterion is written at all.

    .EXAMPLE
        Set-DefaultSheetFilterHC -Path $Path `
            -WorksheetName 'AccessList', 'GroupManagers' `
            -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank

        Opens the workbook with the disabled accounts hidden, while empty
        'MemberEnabled' cells stay visible.

    .EXAMPLE
        Set-DefaultSheetFilterHC -Path $Path `
            -WorksheetName 'AccessList', 'GroupManagers' `
            -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
            -ExcludeColumnName 'MemberSamAccountName', 'ManagerMemberName' `
            -ExcludeValue 'cnorris', 'Chuck Norris'

        The same, with the placeholder account hidden as well: on 'AccessList'
        it is matched on its SamAccountName, on 'GroupManagers' on its display
        name.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$WorksheetName,
        [Parameter(Mandatory)][string]$ColumnName,
        [string[]]$VisibleValue = @('TRUE'),
        [switch]$IncludeBlank,
        [string[]]$ExcludeColumnName = @(),
        [string[]]$ExcludeValue = @()
    )

    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    #region Values to hide, matched case insensitively
    $excluded = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($ExcludeValue | Where-Object { $_ }),
        [System.StringComparer]::OrdinalIgnoreCase
    )
    #endregion

    # Reads a cell the same way for the hiding pass and for the value list, so
    # a value written into the filter XML always matches the row it came from
    $getCellText = {
        param($Cell)

        $raw = $Cell.Value

        if ($null -eq $raw) { return '' }

        return ([string]$raw).Trim()
    }

    try {
        $excelPackage = Open-ExcelPackage -Path $Path -ErrorAction Stop

        try {
            foreach ($sheetName in $WorksheetName) {
                $sheet = $excelPackage.Workbook.Worksheets[$sheetName]

                if ((-not $sheet) -or (-not $sheet.Dimension)) { continue }
                if ($sheet.Dimension.End.Row -le 1) { continue }

                $lastRow = $sheet.Dimension.End.Row

                #region Map every header on row 1 to its column index
                $headerIndex = @{}

                for (
                    $c = $sheet.Dimension.Start.Column
                    $c -le $sheet.Dimension.End.Column
                    $c++
                ) {
                    $header = $sheet.Cells[1, $c].Text

                    if ($header -and (-not $headerIndex.ContainsKey($header))) {
                        $headerIndex[$header] = $c
                    }
                }
                #endregion

                #region Build the criteria that apply to this worksheet
                $rules = [System.Collections.Generic.List[hashtable]]::new()

                if ($headerIndex.ContainsKey($ColumnName)) {
                    $rules.Add(
                        @{
                            Column       = $headerIndex[$ColumnName]
                            VisibleValue = [string[]]@($VisibleValue)
                        }
                    )
                }

                if ($excluded.Count -gt 0) {
                    foreach (
                        $name in @($ExcludeColumnName | Where-Object { $_ })
                    ) {
                        if (-not $headerIndex.ContainsKey($name)) { continue }

                        $index = $headerIndex[$name]

                        # Turn "hide these" into the inclusion list Excel
                        # understands: every other distinct value in the column
                        $keep = [System.Collections.Generic.List[string]]::new()
                        $seen = [System.Collections.Generic.HashSet[string]]::new(
                            [System.StringComparer]::OrdinalIgnoreCase
                        )

                        for ($r = 2; $r -le $lastRow; $r++) {
                            $text = & $getCellText $sheet.Cells[$r, $index]

                            if ($text -eq '') { continue }
                            if ($excluded.Contains($text)) { continue }
                            if ($seen.Add($text)) { $keep.Add($text) }
                        }

                        $rules.Add(
                            @{
                                Column       = $index
                                VisibleValue = [string[]]$keep.ToArray()
                            }
                        )
                    }
                }

                if ($rules.Count -eq 0) { continue }

                # CT_AutoFilter expects filterColumn children in ascending
                # colId order
                $rules = @(
                    $rules | Sort-Object -Property { [int]$_.Column }
                )
                #endregion

                #region Hide the rows that fail any criterion
                for ($r = 2; $r -le $lastRow; $r++) {
                    foreach ($rule in $rules) {
                        $text = & $getCellText $sheet.Cells[$r, $rule.Column]

                        $isVisible = if ($text -eq '') {
                            $IncludeBlank.IsPresent
                        }
                        else {
                            $rule.VisibleValue -contains $text
                        }

                        if (-not $isVisible) {
                            $sheet.Row($r).Hidden = $true
                            break
                        }
                    }
                }
                #endregion

                #region Write the criteria into the table XML
                $touchedAutoFilters = [System.Collections.Generic.List[object]]::new()

                foreach ($rule in $rules) {
                    $table = $sheet.Tables | Where-Object {
                        ($_.Address.Start.Column -le $rule.Column) -and
                        ($_.Address.End.Column -ge $rule.Column)
                    } | Select-Object -First 1

                    if (-not $table) { continue }

                    # Every value in the column is excluded and blanks are
                    # hidden too: an empty 'filters' element would make Excel
                    # report the workbook as corrupt
                    if (
                        ($rule.VisibleValue.Count -eq 0) -and
                        (-not $IncludeBlank)
                    ) {
                        continue
                    }

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

                    if (-not $touchedAutoFilters.Contains($autoFilter)) {
                        $null = $touchedAutoFilters.Add($autoFilter)
                    }

                    $colId = (
                        $rule.Column - $table.Address.Start.Column
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

                    foreach ($value in $rule.VisibleValue) {
                        $filter = $tableXml.CreateElement('filter', $ns)
                        $filter.SetAttribute('val', $value)
                        $null = $filters.AppendChild($filter)
                    }

                    $null = $filterColumn.AppendChild($filters)
                    $null = $autoFilter.AppendChild($filterColumn)
                }
                #endregion

                #region Restore ascending colId order
                # Re-applying a filter removes a criterion and appends the new
                # one at the end, which can leave the children out of sequence
                foreach ($autoFilter in $touchedAutoFilters) {
                    $ordered = @(
                        $autoFilter.SelectNodes(
                            "*[local-name()='filterColumn']"
                        ) | Sort-Object -Property {
                            [int]$_.GetAttribute('colId')
                        }
                    )

                    foreach ($node in $ordered) {
                        $null = $autoFilter.RemoveChild($node)
                        $null = $autoFilter.AppendChild($node)
                    }
                }
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