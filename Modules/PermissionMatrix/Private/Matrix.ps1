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

# function ConvertTo-MatrixAclHC {
#     <#
#         Converts the permissions sheet (minus header rows) and AD objects
#         into a structured matrix of ACL assignments.

#         Params:
#             -NonHeaderRows = rows after first 3
#             -ADObjects     = array of AD identifiers
#     #>
#     [CmdletBinding()]
#     param(
#         [Parameter(Mandatory)][array]$NonHeaderRows,
#         [Parameter(Mandatory)][array]$ADObjects
#     )

#     $matrix = @()

#     foreach ($row in $NonHeaderRows) {

#         $path = $row.P1
#         if (-not $path) {
#             continue
#         }

#         $entry = [ordered]@{
#             Path = $path
#             ACL  = @{}
#         }

#         # For each AD object, assign the permission
#         for ($i = 0; $i -lt $ADObjects.Count; $i++) {

#             $colName = "P$($i+2)"   # Permissions columns start at P2
#             if ($row.PSObject.Properties[$colName]) {
#                 $perm = $row.$colName
#                 if ($perm -and $perm -ne 'I') {
#                     # Ignore == skip
#                     $entry.ACL[$ADObjects[$i]] = $perm
#                 }
#             }
#         }

#         $matrix += [pscustomobject]$entry
#     }

#     return $matrix
# }

function Get-DefaultAclHC {
    <#
        Builds the default ACL hash table from the Defaults.xlsx sheet.
        Sheet must contain at least: MailTo, ADObjectName, Permission

        Returns a hashtable:
            Key:   ADObjectName
            Value: Permission Char
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Sheet
    )

    $acl = @{}

    foreach ($row in $Sheet) {
        if ($row.ADObjectName -and $row.Permission) {
            $name = $row.ADObjectName.ToString().Trim()
            $perm = $row.Permission.ToString().Trim().ToUpper()

            if ($name -and $perm) {
                $acl[$name] = $perm
            }
        }
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