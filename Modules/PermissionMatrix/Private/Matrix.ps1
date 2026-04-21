
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

function ConvertTo-MatrixAclHC {
    <#
        Converts the permissions sheet (minus header rows) and AD objects
        into a structured matrix of ACL assignments.

        Params:
            -NonHeaderRows = rows after first 3
            -ADObjects     = array of AD identifiers
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$NonHeaderRows,
        [Parameter(Mandatory)][array]$ADObjects
    )

    $matrix = @()

    foreach ($row in $NonHeaderRows) {

        $path = $row.P1
        if (-not $path) {
            continue
        }

        $entry = [ordered]@{
            Path = $path
            ACL  = @{}
        }

        # For each AD object, assign the permission
        for ($i = 0; $i -lt $ADObjects.Count; $i++) {

            $colName = "P$($i+2)"   # Permissions columns start at P2
            if ($row.PSObject.Properties[$colName]) {
                $perm = $row.$colName
                if ($perm -and $perm -ne 'I') {
                    # Ignore == skip
                    $entry.ACL[$ADObjects[$i]] = $perm
                }
            }
        }

        $matrix += [pscustomobject]$entry
    }

    return $matrix
}

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