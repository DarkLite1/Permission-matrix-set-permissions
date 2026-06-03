function Format-FormDataStringsHC {
    <#
    .SYNOPSIS
        Return a copy of a FormData row with all string values trimmed.

    .DESCRIPTION
        Produces a new object from the input row in which every string-valued
        property has its leading and trailing whitespace removed. Values that
        are not strings (numbers, dates, booleans, $null, arrays, nested
        objects, and so on) are copied across unchanged.

        The original property order is preserved by building the result through
        an ordered dictionary, so the column layout from the source (for example
        an Excel or CSV import) is kept intact. The input object is not
        modified; a new PSCustomObject is returned.

        The function accepts pipeline input and processes one row at a time,
        emitting one cleaned object per input row. This lets it sit directly in
        a pipeline after a command such as Import-Csv or Import-Excel.

    .PARAMETER Row
        The row to normalize: any object whose properties should have their
        string values trimmed. Accepts pipeline input, so a stream of rows can
        be piped in and each is processed and emitted individually. Every
        property exposed by the object (via PSObject.Properties) is examined.

    .EXAMPLE
        [pscustomobject]@{ Name = '  Bob  '; Age = 30 } | Format-FormDataStringsHC

        Returns an object where Name is 'Bob' (trimmed) and Age is still the
        integer 30 (left unchanged, because it is not a string).

    .EXAMPLE
        Import-Csv 'C:\data\forms.csv' | Format-FormDataStringsHC

        Trims every string field of every row in the CSV, preserving the
        original column order, and emits one cleaned object per row.

    .EXAMPLE
        $row = [pscustomobject]@{ Id = 5; Label = ' active ' }
        $clean = $row | Format-FormDataStringsHC
        $row.Label    # still ' active '
        $clean.Label  # 'active'

        Shows that a new object is returned and the original row is left
        unmodified.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
        One object per input row, exposing the same properties in the same
        order, with string values trimmed.

    .NOTES
        - Only scalar [string] values are trimmed. Everything else is copied
          unchanged, including $null, numbers, dates, booleans and arrays. A
          string element inside an array property is therefore not trimmed.
        - The input object is not mutated; a new object is returned.
        - Property/column order is preserved via [ordered].
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)]
        $Row
    )

    process {
        # Use [ordered] to preserve the exact column layout from Excel
        $new = [ordered]@{} 
        
        foreach ($prop in $Row.PSObject.Properties) {
            $val = $prop.Value
            if ($val -is [string]) {
                $val = $val.Trim()
            }
            $new[$prop.Name] = $val
        }

        return [pscustomobject]$new
    }
}

function Format-PermissionsStringsHC {
    <#
    .SYNOPSIS
        Return a copy of a Permissions row with all string values trimmed and
        uppercased.

    .DESCRIPTION
        Produces a new object from a row returned by Import-Excel for the
        Permissions sheet, in which every string-valued property is trimmed of
        leading and trailing whitespace and converted to upper case. Values
        that are not strings (numbers, dates, booleans, $null, arrays, nested
        objects, and so on) are copied across unchanged.

        The original property order is preserved by building the result through
        an ordered dictionary, so the Excel column layout is kept intact. The
        input object is not modified; a new PSCustomObject is returned.

        The function accepts pipeline input and processes one row at a time,
        emitting one normalized object per input row, so it can sit directly in
        a pipeline after Import-Excel.

    .PARAMETER Row
        The row to normalize: any object whose string properties should be
        trimmed and uppercased. Accepts pipeline input, so a stream of rows can
        be piped in and each is processed and emitted individually. Every
        property exposed by the object (via PSObject.Properties) is examined.

    .EXAMPLE
        [pscustomobject]@{ Account = '  domain\bob  '; Level = 30 } | Format-PermissionsStringsHC

        Returns an object where Account is 'DOMAIN\BOB' (trimmed and uppercased)
        and Level is still the integer 30 (left unchanged, because it is not a
        string).

    .EXAMPLE
        Import-Excel 'C:\data\matrix.xlsx' -WorksheetName 'Permissions' | Format-PermissionsStringsHC

        Trims and uppercases every string field of every row on the Permissions
        sheet, preserving the original column order, and emits one normalized
        object per row.

    .EXAMPLE
        $row = [pscustomobject]@{ Id = 5; Right = ' read ' }
        $clean = $row | Format-PermissionsStringsHC
        $row.Right    # still ' read '
        $clean.Right  # 'READ'

        Shows that a new object is returned and the original row is left
        unmodified.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
        One object per input row, exposing the same properties in the same
        order, with string values trimmed and uppercased.

    .NOTES
        - Only scalar [string] values are trimmed and uppercased. Everything
          else is copied unchanged, including $null, numbers, dates, booleans
          and arrays. A string element inside an array property is therefore
          left as-is.
        - Trimming removes leading and trailing whitespace only; whitespace
          inside the value (for example between words) is preserved.
        - The input object is not mutated; a new object is returned.
        - Property/column order is preserved via [ordered].
    #>

    [CmdletBinding()]
    param(
        # Allow the function to accept rows directly from the pipeline
        [Parameter(Mandatory, ValueFromPipeline = $true)]
        $Row
    )

    process {
        # MUST use [ordered] to preserve the exact Excel column layout!
        $new = [ordered]@{} 
        
        foreach ($prop in $Row.PSObject.Properties) {
            $val = $prop.Value
            if ($val -is [string]) {
                $val = $val.Trim().ToUpper()
            }
            $new[$prop.Name] = $val
        }

        return [pscustomobject]$new
    }
}

function Format-SettingStringsHC {
    <#
    .SYNOPSIS
        Return a normalized copy of a Settings row: trimmed strings, cleaned
        Path, uppercased ComputerName, title-cased Action, and a boolean
        ApplyDefaultPermissions.

    .DESCRIPTION
        Takes a Settings object (typically a row from Import-Excel) and returns
        a cleaned shallow copy. The original object is not reassigned in place;
        a copy is made via PSObject.Copy() and all changes are applied to the
        copy.

        The following normalizations are applied, in order:

        - Every string-valued property has its leading and trailing whitespace
          trimmed. Non-string properties are left unchanged.
        - Path has any trailing '\' and '/' characters removed.
        - ComputerName is converted to upper case.
        - Action is converted to title case (for example 'fIx' becomes 'Fix'
          and 'REPORT' becomes 'Report'), for consistent UI reporting.
        - ApplyDefaultPermissions, when present and non-empty, is parsed into a
          real [bool].

        Each of the named transforms is applied only when its property is
        present and not null/empty/whitespace, so a Settings object missing any
        of these properties is handled without error.

        The function accepts pipeline input and processes one row at a time,
        emitting one normalized object per input row, so it can sit directly in
        a pipeline after Import-Excel.

    .PARAMETER Settings
        The Settings row to normalize. Accepts pipeline input, so a stream of
        rows can be piped in and each is processed and emitted individually.
        All string properties are trimmed; the properties Path, ComputerName,
        Action and ApplyDefaultPermissions receive additional, property-specific
        treatment when present.

    .EXAMPLE
        $s = [pscustomobject]@{
            Path                    = '  C:\Data\Share\  '
            ComputerName            = ' server01 '
            Action                  = 'fIx'
            ApplyDefaultPermissions = 'true'
        }
        $s | Format-SettingStringsHC

        Returns an object with Path 'C:\Data\Share', ComputerName 'SERVER01',
        Action 'Fix', and ApplyDefaultPermissions as the boolean $true.

    .EXAMPLE
        Import-Excel 'C:\data\matrix.xlsx' -WorksheetName 'Settings' | Format-SettingStringsHC

        Normalizes every row on the Settings sheet, emitting one cleaned object
        per row.

    .EXAMPLE
        $s = [pscustomobject]@{ Path = 'C:\Logs'; Note = '  keep me  ' }
        $s | Format-SettingStringsHC

        Returns Path 'C:\Logs' and Note 'keep me'. ComputerName, Action and
        ApplyDefaultPermissions are absent, so only the universal string
        trimming applies.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
        One normalized object per input row, of the same type and shape as the
        input, with the transforms above applied.

    .NOTES
        - The copy is shallow (PSObject.Copy()). Reassigning scalar string and
          boolean properties on the copy does not affect the input, but any
          reference-type property (array, nested object) is shared with the
          original; mutating its contents would affect both.
        - ApplyDefaultPermissions is parsed with [bool]::TryParse, which only
          recognizes the text 'true'/'false' (case-insensitive). Any other
          value, including '1', '0', 'yes' and 'no', fails to parse and results
          in $false.
        - Path trimming removes every trailing slash/backslash, not just one.
        - ComputerName uppercasing and Action title-casing use the current
          culture.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)]
        [object]$Settings
    )

    process {
        # Create a shallow copy so we don't mutate the raw object
        $S = $Settings.PSObject.Copy()

        # Universally trim ALL string properties
        foreach ($p in $S.PSObject.Properties) {
            if ($p.Value -is [string]) {
                $p.Value = $p.Value.Trim()
            }
        }

        # Clean Path: Strip trailing slashes
        if (-not [string]::IsNullOrWhiteSpace($S.Path)) {
            $S.Path = $S.Path.TrimEnd([char[]]@('\', '/'))
        }

        # ComputerName to uppercase
        if (-not [string]::IsNullOrWhiteSpace($S.ComputerName)) {
            $S.ComputerName = $S.ComputerName.ToUpper()
        }

        # Clean Action: TitleCase for clean UI reporting
        # (e.g., 'fIx' -> 'Fix', 'REPORT' -> 'Report')
        if (-not [string]::IsNullOrWhiteSpace($S.Action)) {
            $S.Action = (Get-Culture).TextInfo.ToTitleCase($S.Action.ToLower())
        }

        # Convert ApplyDefaultPermissions to boolean
        if (
            $S.PSObject.Properties.Match('ApplyDefaultPermissions').Count -gt 0 -and
            -not [string]::IsNullOrWhiteSpace($S.ApplyDefaultPermissions)
        ) {
            $parsed = $false
            $null = [bool]::TryParse($S.ApplyDefaultPermissions.ToString(), [ref]$parsed)
            $S.ApplyDefaultPermissions = $parsed
        }

        return $S
    }
}

function ConvertTo-MatrixADNamesHC {
    <#
        Converts matrix column headers + GroupName + SiteCode into AD lookup objects.
        This is used to build per-settings AD identifiers.

        Params:
            -Begin         = GroupName
            -Middle        = SiteCode
            -ColumnHeaders = First 3 header rows from Permissions sheet
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Begin,
        [Parameter(Mandatory)][string]$Middle,
        [Parameter(Mandatory)][array]$ColumnHeaders
    )

    $list = @()

    # Add begin/middle values directly (group & site)
    if ($Begin) { $list += $Begin }
    if ($Middle) { $list += $Middle }

    # Extract SamAccountNames from header rows
    foreach ($h in $ColumnHeaders) {
        if ($h.P2) { $list += $h.P2 }
    }

    return $list | Where-Object { $_ } | Sort-Object -Unique
}

function Get-DefaultAclHC {
    <#
    .SYNOPSIS
        Builds the default ACL hash table from the Defaults.xlsx Settings sheet.

    .DESCRIPTION
        Validates each row that has either ADObjectName or Permission populated:
        - both must be present (incomplete pairs are flagged)
        - Permission must be a valid character (L, R, W, F)
        - duplicate ADObjectNames are flagged

        Rows where both ADObjectName and Permission are empty are ignored
        (these are typically MailTo-only rows or trailing blank rows).

        Permission 'I' (inherit) is intentionally rejected here — defaults
        are explicit grants by definition; "inherit by default" is meaningless.

        Returns a hashtable:
            Key:   ADObjectName
            Value: Permission Char
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Sheet,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    # Mirrors Test-MatrixPermissionsHC's accepted set, minus 'I' (inherit).
    $validPermissions = @('L', 'R', 'W', 'F')

    $acl = @{}

    foreach ($row in $Sheet) {
        $rawName = if ($row.ADObjectName) { 
            $row.ADObjectName.ToString().Trim() 
        }
        else { '' }
        $rawPerm = if ($row.Permission) { 
            $row.Permission.ToString().Trim().ToUpper() 
        }
        else { '' }

        # Both empty: not an ACL row (likely MailTo-only). Skip silently.
        if (-not $rawName -and -not $rawPerm) { continue }

        # ADObjectName missing but Permission set
        if (-not $rawName) {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Incomplete default ACL entry' `
                -Message "Defaults row has Permission '$rawPerm' but no ADObjectName." `
                -Category 'Matrix' `
                -SystemErrors $SystemErrors
            continue
        }

        # ADObjectName set but Permission missing
        if (-not $rawPerm) {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Incomplete default ACL entry' `
                -Message "Defaults entry '$rawName' has no permission assigned." `
                -Category 'Matrix' `
                -SystemErrors $SystemErrors
            continue
        }

        # Permission character invalid
        if ($rawPerm -notin $validPermissions) {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Invalid default ACL permission' `
                -Message "Defaults entry '$rawName' has invalid permission '$rawPerm'. Valid values: $($validPermissions -join ', ')." `
                -Category 'Matrix' `
                -SystemErrors $SystemErrors
            continue
        }

        # Duplicate ADObjectName in defaults
        if ($acl.ContainsKey($rawName)) {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Duplicate default ACL entry' `
                -Message "Defaults defines '$rawName' more than once." `
                -Category 'Matrix' `
                -SystemErrors $SystemErrors
            continue
        }

        $acl[$rawName] = $rawPerm
    }

    return $acl
}

function Get-MatrixADObjectsMapHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$PermissionsSheet,
        [Parameter(Mandatory)][object]$SettingRow
    )

    $headerRows = $PermissionsSheet | Select-Object -First 3
    $adObjectsMap = [ordered]@{}

    $colIndex = 2
    while ($true) {
        $colName = "P$colIndex"

        # Stop if the column doesn't exist
        if (-not $headerRows[0].PSObject.Properties.Match($colName).Count) {
            break
        }

        # Walk header rows bottom-to-top, resolving placeholders.
        # Each non-empty row contributes one part; the parts are joined with
        # a single space. Empty rows are skipped so we don't get double spaces.
        $parts = for ($i = $headerRows.Count - 1; $i -ge 0; $i--) {
            $cellValue = $headerRows[$i].$colName
            if ([string]::IsNullOrWhiteSpace($cellValue)) { continue }

            switch ($cellValue) {
                'GroupName' { $SettingRow.GroupName }
                'SiteCode' { $SettingRow.SiteCode }
                default { $cellValue }
            }
        }

        $adName = ($parts -join ' ').Trim()

        if (-not [string]::IsNullOrWhiteSpace($adName)) {
            $adObjectsMap[$colName] = $adName
        }

        $colIndex++
    }

    return $adObjectsMap
}

function ConvertTo-MatrixAclHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$DataRows,
        [Parameter(Mandatory)][hashtable]$AdObjectsMap
    )

    $matrix = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($row in $DataRows) {
        if (-not $row.P1) { continue }

        $acl = @{}
        foreach ($colName in $AdObjectsMap.Keys) {
            $perm = $row.$colName
            if ($perm -and $perm -ne 'I') {
                # Map the permission to the resolved AD Object name
                $acl[$AdObjectsMap[$colName]] = $perm
            }
        }

        $matrix.Add(
            [pscustomobject]@{
                Path = $row.P1
                ACL  = $acl
            }
        )
    }

    return $matrix.ToArray()
}

function Merge-DefaultPermissionsHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Defaults,
        [Parameter(Mandatory)][hashtable]$MatrixAcl,
        [Parameter(Mandatory)][bool]$ApplyDefaultPermissions
    )

    # Note: .Clone() ensures we don't accidentally link objects in memory
    if (-not $ApplyDefaultPermissions) {
        return $MatrixAcl.Clone()
    }

    # Check for conflicts where the same key exists in both hashtables
    $conflicts = $Defaults.Keys | Where-Object { $MatrixAcl.ContainsKey($_) }
    if ($conflicts) {
        throw "Defaults conflict detected. The following AD Objects are defined in both the Matrix and Defaults: $($conflicts -join ', ')"
    }

    # No conflicts, safely merge defaults into the Matrix ACL
    $mergedAcl = $MatrixAcl.Clone()
    foreach ($key in $Defaults.Keys) {
        $mergedAcl[$key] = $Defaults[$key]
    }

    return $mergedAcl
}