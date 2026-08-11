function Format-FormDataStringsHC {
    <#
    .SYNOPSIS
        Return a copy of a FormData row with all string values trimmed.

    .DESCRIPTION
        Non-string values (numbers, dates, booleans, $null, arrays, nested
        objects) are copied across unchanged. The input row is not modified.

    .NOTES
        Only scalar [string] values are trimmed, so a string element inside an
        array property is left as-is.

    .EXAMPLE
        Import-Csv 'C:\data\forms.csv' | Format-FormDataStringsHC
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
        Return a copy of a Permissions row with all string values trimmed, and
        every column except the P1 path column upper-cased.

    .DESCRIPTION
        P1 holds the folder path (Column A) and keeps its original
        capitalization; every other column holds a permission character matched
        case-insensitively and is upper-cased. Non-string values are copied
        across unchanged and the input row is not modified.

    .NOTES
        Only scalar [string] values are handled, so a string element inside an
        array property is left as-is.

    .EXAMPLE
        Import-Excel 'C:\data\matrix.xlsx' -WorksheetName 'Permissions' | Format-PermissionsStringsHC
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
                # P1 is the folder path (Column A). It is used verbatim to
                # create missing folders on disk, so its original capitalization
                # MUST be preserved. Remove trailing separators here so every
                # downstream matrix consumer sees one canonical folder path.
                # Every other column holds a permission character (R/W/L/F/I)
                # that is matched case-insensitively, so those are trimmed AND
                # upper-cased.
                if ($prop.Name -eq 'P1') {
                    $val = $val.Trim()

                    $rootPath = [System.IO.Path]::GetPathRoot($val)
                    if ((-not $rootPath) -or ($val.Length -gt $rootPath.Length)) {
                        $val = $val.TrimEnd(
                            [System.IO.Path]::DirectorySeparatorChar,
                            [System.IO.Path]::AltDirectorySeparatorChar
                        )
                    }
                }
                else {
                    $val = $val.Trim().ToUpper()
                }
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
        Each named transform is applied only when its property is present and
        not blank, so a Settings object missing any of them is handled without
        error.

    .NOTES
        - The copy is shallow (PSObject.Copy()). Reassigning scalar properties
          does not affect the input, but any reference-type property (array,
          nested object) is shared with the original.
        - ApplyDefaultPermissions is parsed with [bool]::TryParse, which only
          recognizes the text 'true'/'false'. Any other value, INCLUDING '1',
          '0', 'yes' and 'no', fails to parse and results in $false.
          Test-MatrixSettingRowHC rejects those values before this runs.
        - ComputerName uppercasing and Action title-casing use the current
          culture.

    .EXAMPLE
        Import-Excel 'C:\data\matrix.xlsx' -WorksheetName 'Settings' | Format-SettingStringsHC
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

function Get-DefaultAclHC {
    <#
    .SYNOPSIS
        Build the default ACL hashtable from the Defaults.xlsx Settings sheet,
        validating each entry and recording problems in SystemErrors.

    .DESCRIPTION
        Maps each ADObjectName to its permission character. Rows where both
        ADObjectName and Permission are empty are not ACL rows (typically
        MailTo-only or trailing blanks) and are skipped silently. Incomplete
        pairs, invalid permissions and duplicates are rejected.

    .NOTES
        - Does NOT throw. Problems are appended to SystemErrors via Add-ErrorHC
          and the function still returns the rows that passed, so callers must
          inspect SystemErrors rather than just the returned hashtable.
        - Permission 'I' (inherit) is deliberately rejected: defaults are
          explicit grants by definition, so "inherit by default" is meaningless.
        - On a duplicate ADObjectName the first occurrence is kept. Key matching
          is case-insensitive, so names differing only in case collide.

    .PARAMETER Sheet
        The rows of the Defaults Settings sheet, each exposing ADObjectName and
        Permission. An empty collection yields an empty hashtable.

    .PARAMETER SystemErrors
        A [ref] to the caller's system-error accumulator. In/out parameter.

    .EXAMPLE
        $errors = [System.Collections.Generic.List[object]]::new()
        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)
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
        The first three rows of PermissionsSheet are header rows. Columns are
        scanned from P2 upward; for each one the three header cells are walked
        bottom-to-top, 'GroupName'/'SiteCode' are replaced with the values from
        SettingRow, and the parts are joined with a single space.

    .NOTES
        - Scanning STOPS at the first column name absent from the FIRST header
          row, so a gap in the numbering (P2, P3, then P5) ends the scan early,
          and a column present only on a later header row is never reached.
        - Empty header cells are skipped so they cannot introduce a blank part,
          but an empty *resolved* placeholder is not: a cell saying 'GroupName'
          with an empty SettingRow.GroupName still produces an empty part.
        - Columns assembling to an empty name are omitted from the map.
        - Placeholder matching uses a switch, so it is case-insensitive.

    .PARAMETER PermissionsSheet
        The Permissions sheet rows. Only the first three are used, as headers.

    .PARAMETER SettingRow
        The setting row supplying the GroupName and SiteCode values.

    .EXAMPLE
        $setting = [pscustomobject]@{ GroupName = 'GRP'; SiteCode = 'BRU' }
        $sheet = @(
            [pscustomobject]@{ P2 = 'GroupName'; P3 = 'GroupName' }
            [pscustomobject]@{ P2 = 'SiteCode';  P3 = '' }
            [pscustomobject]@{ P2 = 'Mgrs';      P3 = 'Users' }
        )
        Get-MatrixADObjectsMapHC -PermissionsSheet $sheet -SettingRow $setting

        Returns @{ P2 = 'Mgrs BRU GRP'; P3 = 'Users GRP' }. P3's empty middle
        cell is skipped rather than producing a double space.
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
        Emits one object per data row with a non-empty P1, each carrying Path,
        ACL (resolved AD name to permission character) and Ignore. Only the
        columns present in AdObjectsMap are read; anything else on the row is
        ignored.

    .NOTES
        - Permission values are NOT validated against a permitted set here,
          unlike Get-DefaultAclHC. Only two values are special-cased: empty
          (skipped) and 'I' (flags the row as ignored).
        - A single 'I' in ANY permission column flags the whole row: it gets an
          empty ACL and Ignore = $true, and any other permissions on that row
          are discarded, because an ignored folder is left untouched.
        - A row whose every permission is empty still produces an entry with an
          empty ACL and Ignore = $false. Downstream that means inherit-only.
        - If two columns in AdObjectsMap resolve to the same AD name, the later
          column's permission silently overwrites the earlier one.

    .PARAMETER DataRows
        The Permissions sheet rows below the headers. Rows with an empty P1 are
        dropped.

    .PARAMETER AdObjectsMap
        Column name to resolved AD object name, typically from
        Get-MatrixADObjectsMapHC.

    .EXAMPLE
        $map = @{ P2 = 'Mgrs BRU GRP'; P3 = 'Users GRP' }
        $rows = @(
            [pscustomobject]@{ P1 = '\\srv\Finance'; P2 = 'F'; P3 = 'R' },
            [pscustomobject]@{ P1 = '\\srv\HR';      P2 = 'I'; P3 = 'W' }
        )
        ConvertTo-MatrixAclHC -DataRows $rows -AdObjectsMap $map

        Finance gets both permissions. HR is flagged ignored and its 'W' is
        discarded.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$DataRows,
        [Parameter(Mandatory)][hashtable]$AdObjectsMap
    )

    $matrix = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($row in $DataRows) {
        if (-not $row.P1) { continue }

        $acl = @{}
        $isIgnored = $false

        foreach ($colName in $AdObjectsMap.Keys) {
            $perm = $row.$colName

            if (-not $perm) { continue }

            # 'I' (Ignore) marks the whole folder entry to be skipped: the
            # script must not touch it or apply any permissions, whether from
            # the matrix or the defaults. A single 'I' in any permission column
            # flags the row, matching the ignore detection in the validation
            # stage and the documented behaviour in the README. Values reaching
            # here are already trimmed and upper-cased by
            # Format-PermissionsStringsHC.
            if ("$perm".Trim().ToUpper() -eq 'I') {
                $isIgnored = $true
                continue
            }

            # Map the permission to the resolved AD Object name
            $acl[$AdObjectsMap[$colName]] = $perm
        }

        # An ignored entry carries no ACL. Downstream defaults merging and
        # SetPermissions.ps1 both key off the Ignore flag to leave the folder
        # (and its subtree) completely untouched.
        $matrix.Add(
            [pscustomobject]@{
                Path   = $row.P1
                ACL    = if ($isIgnored) { @{} } else { $acl }
                Ignore = $isIgnored
            }
        )
    }

    return $matrix.ToArray()
}

function Merge-DefaultPermissionsHC {
    <#
    .SYNOPSIS
        Merge the default ACL into a matrix ACL, unless defaults are disabled or
        the two define the same AD object.

    .DESCRIPTION
        With ApplyDefaultPermissions $false, Defaults is not consulted at all
        and a clone of MatrixAcl is returned. With $true, the default entries
        are added to a clone of MatrixAcl provided no AD object appears in both.

    .NOTES
        - A conflict is a hard error, not resolved by precedence: an AD object
          must be defined by the matrix or by the defaults, never both. The
          function THROWS and returns nothing, so callers must be ready for a
          terminating error rather than an empty result.
        - Conflict detection is case-insensitive: 'Admins' and 'admins' collide.
        - Neither input is mutated; the result is always a clone.
        - Permission characters are copied as-is, not validated or normalized.

    .EXAMPLE
        Merge-DefaultPermissionsHC `
            -Defaults @{ 'Admins' = 'F' } `
            -MatrixAcl @{ 'Users GRP' = 'R' } `
            -ApplyDefaultPermissions $true

        Returns @{ 'Users GRP' = 'R'; 'Admins' = 'F' }.
    #>
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