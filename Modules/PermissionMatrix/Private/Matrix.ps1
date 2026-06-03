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
    .SYNOPSIS
        Build a sorted, de-duplicated list of AD names from a group name, a
        site code and the SamAccountNames in the Permissions header rows.

    .DESCRIPTION
        Collects a flat list of name strings to be used later for AD lookups,
        drawn from three sources:

        - Begin (the group name), added as-is.
        - Middle (the site code), added as-is.
        - The P2 property of each object in ColumnHeaders (a SamAccountName).

        The combined values are then filtered to drop empty entries, sorted and
        de-duplicated. The result is a list of plain strings, not objects; these
        names are the inputs to a subsequent AD lookup rather than AD objects
        themselves.

    .PARAMETER Begin
        The group name (GroupName). Added directly to the result list.
        Mandatory.

    .PARAMETER Middle
        The site code (SiteCode). Added directly to the result list. Mandatory.

    .PARAMETER ColumnHeaders
        An array of header-row objects from the Permissions sheet (the caller
        typically passes the first three header rows). The P2 property of each
        object — a SamAccountName — is collected; all other properties are
        ignored. Objects without a P2 value contribute nothing.

    .EXAMPLE
        $headers = @(
            [pscustomobject]@{ P2 = 'svc-app1' },
            [pscustomobject]@{ P2 = 'svc-app2' },
            [pscustomobject]@{ P2 = $null }
        )
        ConvertTo-MatrixADNamesHC -Begin 'Finance-RW' -Middle 'BRU' -ColumnHeaders $headers

        Returns 'BRU', 'Finance-RW', 'svc-app1' and 'svc-app2', sorted and
        unique. The header whose P2 is $null contributes nothing.

    .EXAMPLE
        $headers = @([pscustomobject]@{ P2 = 'BRU' })
        ConvertTo-MatrixADNamesHC -Begin 'Finance-RW' -Middle 'BRU' -ColumnHeaders $headers

        Returns 'BRU' and 'Finance-RW'. The site code 'BRU' and the header's P2
        'BRU' are the same value, so the duplicate is collapsed.

    .OUTPUTS
        System.String
        Zero or more unique name strings, sorted alphabetically.

    .NOTES
        - The result is a flat list of strings, not objects; the names are
          intended as input for later AD lookups.
        - Only the P2 property of each ColumnHeaders object is read. Other
          properties (P1, P3, and so on) are ignored.
        - The "first 3 header rows" is a caller convention; the function
          processes every element of ColumnHeaders, however many are passed.
        - Values are not trimmed. A value consisting only of whitespace is
          truthy, so it passes both the per-item check and the final filter and
          survives into the result.
        - De-duplication via Sort-Object -Unique is case-insensitive, so names
          differing only in casing are collapsed into one entry.
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
        Build the default ACL hashtable from the Defaults.xlsx Settings sheet,
        validating each entry and recording problems in SystemErrors.

    .DESCRIPTION
        Walks the rows of the Defaults Settings sheet and builds a hashtable
        mapping each ADObjectName to its permission character.

        Each row is classified by whether ADObjectName and/or Permission are
        populated (both are trimmed first; Permission is also upper-cased):

        - Both empty: the row is not an ACL entry (typically a MailTo-only row
          or a trailing blank row) and is skipped silently.
        - Only one of the two populated: the pair is incomplete and a
          FatalError is recorded; the row is skipped.
        - Permission populated but not one of the valid characters: a
          FatalError is recorded; the row is skipped.
        - ADObjectName already seen: the duplicate is reported as a FatalError
          and skipped, so the first occurrence wins.
        - Otherwise: the ADObjectName/Permission pair is added to the result.

        Valid permission characters are L, R, W and F. The permission 'I'
        (inherit) is intentionally rejected: defaults are explicit grants by
        definition, so "inherit by default" is meaningless.

        Validation problems are not thrown; they are appended to the
        SystemErrors accumulator via Add-ErrorHC, so the function always
        returns a hashtable (containing only the rows that passed validation),
        and the caller inspects SystemErrors to decide how to proceed.

    .PARAMETER Sheet
        The rows of the Defaults Settings sheet, as an array of objects (for
        example from Import-Excel). Each row is expected to expose ADObjectName
        and Permission properties. An empty collection is allowed and yields an
        empty hashtable. Mandatory.

    .PARAMETER SystemErrors
        A [ref] to the caller's system-error accumulator. Validation failures
        (incomplete pairs, invalid permissions, duplicates) are added to it via
        Add-ErrorHC rather than thrown. This is an in/out parameter: the
        function appends to whatever it references. Mandatory.

    .EXAMPLE
        $errors = @()
        $sheet = @(
            [pscustomobject]@{ ADObjectName = 'GRP-RW'; Permission = 'w' },
            [pscustomobject]@{ ADObjectName = 'GRP-RO'; Permission = 'R' }
        )
        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

        Returns @{ 'GRP-RW' = 'W'; 'GRP-RO' = 'R' }. The lower-case 'w' is
        upper-cased. $errors stays empty because both rows are valid.

    .EXAMPLE
        $errors = @()
        $sheet = @(
            [pscustomobject]@{ ADObjectName = 'GRP-RW'; Permission = 'F' },
            [pscustomobject]@{ ADObjectName = 'GRP-RW'; Permission = 'R' },
            [pscustomobject]@{ ADObjectName = 'GRP-X';  Permission = 'I' },
            [pscustomobject]@{ ADObjectName = '';       Permission = 'R' }
        )
        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

        Returns @{ 'GRP-RW' = 'F' }. The second 'GRP-RW' is a duplicate, 'I' is
        not a valid default permission, and the last row has a permission but no
        ADObjectName — each adds a FatalError to $errors and is skipped.

    .OUTPUTS
        System.Collections.Hashtable
        A hashtable whose keys are ADObjectNames and whose values are the
        validated, upper-cased permission characters. Only rows that passed all
        validation are included.

    .NOTES
        - The function does not throw on bad data; it accumulates FatalErrors in
          SystemErrors and returns whatever passed validation. Callers must
          check SystemErrors, not just the returned hashtable.
        - On a duplicate ADObjectName the first occurrence is kept and later
          ones are rejected.
        - ADObjectName and Permission are trimmed; Permission is upper-cased, so
          permission matching is case-insensitive. Key matching on ADObjectName
          is whatever the hashtable uses by default — case-insensitive for
          string keys — so names differing only in case collide and the second
          is treated as a duplicate.
        - The valid set (L, R, W, F) mirrors Test-MatrixPermissionsHC minus 'I'.
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
    <#
    .SYNOPSIS
        Build an ordered map of permission column name to assembled AD object
        name, resolving GroupName/SiteCode placeholders from a setting row.

    .DESCRIPTION
        Reads the header rows of the Permissions sheet and, for each permission
        column, assembles the AD object name that column refers to. The result
        is an ordered dictionary keyed by column name (P2, P3, ...) whose values
        are the assembled names.

        The first three rows of PermissionsSheet are treated as header rows.
        Columns are scanned starting at P2 and increasing (P2, P3, P4, ...);
        scanning stops at the first column name that is not present on the first
        header row.

        For each column, the three header cells are walked from the bottom row
        up to the top row. Each cell is resolved as follows:

        - An empty or whitespace cell is skipped (so it cannot introduce a blank
          part).
        - A cell equal to 'GroupName' is replaced with SettingRow.GroupName.
        - A cell equal to 'SiteCode' is replaced with SettingRow.SiteCode.
        - Any other cell is used literally.

        The resolved parts are joined with a single space and trimmed to form
        the AD object name. Columns whose assembled name is empty are omitted
        from the map.

    .PARAMETER PermissionsSheet
        The Permissions sheet as an array of row objects (for example from
        Import-Excel). Only the first three rows are used, as the header rows.
        Each row is expected to expose the permission columns as properties
        named P2, P3, and so on. Mandatory.

    .PARAMETER SettingRow
        The single setting row that supplies the placeholder values. Its
        GroupName and SiteCode properties are substituted wherever the header
        cells contain the literals 'GroupName' or 'SiteCode'. Mandatory.

    .EXAMPLE
        $setting = [pscustomobject]@{ GroupName = 'GRP'; SiteCode = 'BRU' }
        $sheet = @(
            [pscustomobject]@{ P2 = 'GroupName'; P3 = 'GroupName' }
            [pscustomobject]@{ P2 = 'SiteCode';  P3 = '' }
            [pscustomobject]@{ P2 = 'Mgrs';      P3 = 'Users' }
        )
        Get-MatrixADObjectsMapHC -PermissionsSheet $sheet -SettingRow $setting

        Returns an ordered map @{ P2 = 'Mgrs BRU GRP'; P3 = 'Users GRP' }.
        For P2 the rows are walked bottom-to-top: 'Mgrs' (literal), then
        'SiteCode' -> 'BRU', then 'GroupName' -> 'GRP'. For P3 the empty middle
        cell is skipped, leaving 'Users' and 'GRP'.

    .EXAMPLE
        $setting = [pscustomobject]@{ GroupName = 'GRP'; SiteCode = 'BRU' }
        $sheet = @(
            [pscustomobject]@{ P2 = 'GroupName'; P3 = '' }
            [pscustomobject]@{ P2 = 'SiteCode';  P3 = '' }
            [pscustomobject]@{ P2 = 'Admins';    P3 = '' }
        )
        Get-MatrixADObjectsMapHC -PermissionsSheet $sheet -SettingRow $setting

        Returns @{ P2 = 'Admins BRU GRP' }. P3 has only empty header cells, so
        its assembled name is empty and the column is left out of the map.

    .OUTPUTS
        System.Collections.Specialized.OrderedDictionary
        An ordered map whose keys are permission column names (P2, P3, ...) in
        ascending order and whose values are the assembled AD object names.
        Columns that assemble to an empty string are not included.

    .NOTES
        - Only the first three rows of PermissionsSheet are used as header rows.
        - Columns are scanned from P2 upward and scanning stops at the first
          column name absent from the first header row, so a gap in the column
          numbering (e.g. P2, P3, then P5) ends the scan early.
        - Column presence is tested against the first header row only; columns
          that exist on a later header row but not the first are never reached.
        - Header rows are walked bottom-to-top, so the lowest header row's value
          appears first in the assembled name.
        - Placeholder matching uses a switch, which is case-insensitive by
          default, so 'groupname'/'sitecode' in any casing are also resolved.
        - Empty header cells are skipped, but an empty *resolved* placeholder is
          not: if a cell says 'GroupName' and SettingRow.GroupName is empty, an
          empty part is still produced.
    #>
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
     <#
    .SYNOPSIS
        Convert Permissions data rows into per-path ACL objects, using a
        column-to-AD-name map to resolve each permission column.

    .DESCRIPTION
        Walks the data rows of the Permissions sheet and produces one object per
        path, pairing the path with an ACL hashtable that maps resolved AD
        object names to permission characters.

        For each data row:

        - P1 holds the path. Rows with no P1 are skipped entirely.
        - For every column in AdObjectsMap (P2, P3, ...), the row's value in
          that column is the permission for the corresponding AD object. The
          permission is added to the row's ACL only when it is non-empty and
          not 'I' (inherit); the ACL key is the resolved AD name taken from
          AdObjectsMap, and the value is the permission character.

        Each surviving row becomes a PSCustomObject with Path and ACL
        properties. The collected objects are returned as an array.

    .PARAMETER DataRows
        The data rows of the Permissions sheet (the rows below the header
        rows), as an array of objects. Each row is expected to expose P1 (the
        path) and the permission columns P2, P3, and so on. Rows whose P1 is
        empty are ignored. Mandatory.

    .PARAMETER AdObjectsMap
        A hashtable mapping permission column names (P2, P3, ...) to the
        resolved AD object names for those columns — typically the output of
        Get-MatrixADObjectsMapHC. Only the columns present as keys here are
        read from each data row; any other columns on the row are ignored.
        Mandatory.

    .EXAMPLE
        $map = @{ P2 = 'Mgrs BRU GRP'; P3 = 'Users GRP' }
        $rows = @(
            [pscustomobject]@{ P1 = '\\srv\Finance'; P2 = 'F'; P3 = 'R' },
            [pscustomobject]@{ P1 = '\\srv\HR';      P2 = 'I'; P3 = 'W' }
        )
        ConvertTo-MatrixAclHC -DataRows $rows -AdObjectsMap $map

        Returns two objects. '\\srv\Finance' gets ACL
        @{ 'Mgrs BRU GRP' = 'F'; 'Users GRP' = 'R' }. '\\srv\HR' gets ACL
        @{ 'Users GRP' = 'W' } — its P2 value 'I' is an inherit marker and is
        excluded.

    .EXAMPLE
        $map = @{ P2 = 'Mgrs BRU GRP' }
        $rows = @(
            [pscustomobject]@{ P1 = '';           P2 = 'F' },
            [pscustomobject]@{ P1 = '\\srv\Logs'; P2 = 'I' }
        )
        ConvertTo-MatrixAclHC -DataRows $rows -AdObjectsMap $map

        Returns one object: Path '\\srv\Logs' with an empty ACL hashtable. The
        first row has no P1 and is skipped; the second has a path but its only
        permission is 'I', so nothing is added to its ACL.

    .OUTPUTS
        System.Management.Automation.PSCustomObject[]
        An array of objects, one per data row with a non-empty P1, each having:
        - Path: the value of the row's P1.
        - ACL:  a hashtable mapping resolved AD object names to permission
                characters. May be empty if the row has no granted permissions.
        Returns an empty array when no rows qualify.

    .NOTES
        - Rows with an empty P1 are dropped; a path is required to produce an
          entry.
        - Permission values are taken from the row as-is: they are not trimmed,
          upper-cased or validated against a permitted set, unlike
          Get-DefaultAclHC. Only two values are special-cased — empty (skipped)
          and 'I' (skipped). The 'I' comparison is case-insensitive, so 'i' is
          also excluded, but a value with surrounding whitespace such as ' I '
          does not match and would be kept.
        - A path whose every permission is empty or 'I' still produces an entry,
          with an empty ACL hashtable.
        - The ACL is a default @{} hashtable, so its AD-name keys are matched
          case-insensitively. If two columns in AdObjectsMap resolve to the same
          AD name, the later column's permission overwrites the earlier one's
          for that path, silently.
    #>
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