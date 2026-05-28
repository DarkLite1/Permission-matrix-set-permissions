function Format-FormDataStringsHC {
    <#
    .SYNOPSIS
        Normalizes a FormData row. Ensures all string values are cleanly trimmed.
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
        Normalizes a row returned from Import-Excel for the Permissions sheet.
        Ensures trimming, uppercasing, and removal of whitespace.
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
        Normalizes a Settings row. Ensures trimming, clean paths, and 
        normalized action casing.
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