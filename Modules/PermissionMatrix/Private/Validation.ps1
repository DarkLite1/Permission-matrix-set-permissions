function New-HcError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Name,
        [Parameter()][string]$Description,
        [Parameter()][object]$Value,
        [Parameter()][string]$Category
    )

    return [pscustomobject]@{
        DateTime    = Get-Date
        Type        = $Type
        Name        = $Name
        Description = $Description
        Value       = $Value
        Category    = $Category
    }
}

function ConvertTo-StructuredObjectHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)] 
        $InputObject
    )

    process {
        foreach ($obj in $InputObject) {
            
            if ($null -eq $obj) { continue }

            if ($obj -is [string]) {
                New-HcError -Type 'Information' -Name 'Message' -Description $obj
                continue
            }

            if ($obj -is [hashtable] -or $obj -is [pscustomobject]) {
                $obj
                continue
            }

            New-HcError -Type 'Information' -Name 'UnknownObject' -Description "$obj"
        }
    }
}

function Test-MatrixFileHC {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $MatrixObject)

    $checks = @()

    if (-not $MatrixObject.Settings -or $MatrixObject.Settings.Count -eq 0) {
        $checks += New-HcError -Type 'Warning' -Name 'Matrix disabled' `
            -Description 'No Settings rows found.' -Category 'File'
    }

    if (-not $MatrixObject.Permissions -or $MatrixObject.Permissions.Count -eq 0) {
        $checks += New-HcError -Type 'FatalError' -Name 'Missing Permissions sheet' `
            -Description 'Permissions sheet missing or empty.' -Category 'File'
    }

    return $checks
}

function Test-MatrixPermissionsHC {
    <#
    .SYNOPSIS
        Verify input for the Excel sheet 'Permissions'.

    .DESCRIPTION
        Verify if all input in the Excel sheet 'Permissions' is correct. When
        incorrect input is detected an object is returned containing all the
        details about the issue. 
        This test is best run before expanding the matrix as it will save time.

    .PARAMETER Permissions
        The objects coming from the Excel sheet 'Permissions', as retrieved by
        Import-Excel.
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [parameter(Mandatory)]
        [PSCustomObject[]]$Permissions
    )

    $checks = [System.Collections.Generic.List[pscustomobject]]::new()

    try {
        $Props = $Permissions[0].PSObject.Properties.Name
        $FirstProperty = $Props[0]

        #region Structural Validation (Fatal - Exits Immediately)
        if ($Permissions.Count -lt 4) {
            $checks.Add([pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'Missing rows'
                    Description = 'At least 4 rows are required: 3 header rows and 1 row for the parent folder.'
                    Value       = "$($Permissions.Count) rows"
                })
            return $checks
        }

        if ($Props.Count -lt 2) {
            $checks.Add([pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'Missing columns'
                    Description = 'At least 2 columns are required: 1 for the folder names and 1 where the permissions are defined.'
                    Value       = "$($Props.Count) column"
                })
            return $checks
        }
        #endregion

        #region Missing header SamAccountName
        $missingSamAccountNames = [System.Collections.Generic.List[string]]::new()

        foreach ($col in $Props) {
            if ($col -eq $FirstProperty) { continue }

            if ([string]::IsNullOrWhiteSpace($Permissions[0].$col) -and 
                [string]::IsNullOrWhiteSpace($Permissions[1].$col) -and 
                [string]::IsNullOrWhiteSpace($Permissions[2].$col)) {
                $missingSamAccountNames.Add($col)
            }
        }

        if ($missingSamAccountNames.Count -gt 0) {
            $checks.Add([pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'Missing header SamAccountName'
                    Description = 'The header rows need to contain the SamAccountName of the AD object for which permissions are defined in the column.'
                    Value       = "Columns: $($missingSamAccountNames -join ', ')"
                })
        }
        #endregion

        # Separate Headers from Data
        $NonHeaderRows = $Permissions | Select-Object -Skip 3
        $FolderNames = $NonHeaderRows | Select-Object -Skip 1

        #region Permission character unknown
        $InvalidChars = [System.Collections.Generic.List[string]]::new()
        
        foreach ($Row in $NonHeaderRows) {
            $PermColumns = $Row.PSObject.Properties.Where({ $_.Name -ne $FirstProperty })
            foreach ($Col in $PermColumns) {
                $Ace = $Col.Value
                if (
                    -not [string]::IsNullOrWhiteSpace($Ace) -and 
                    $Ace -notmatch '^(L|R|W|I|F)$'
                ) {
                    $InvalidChars.Add($Ace)
                }
            }
        }

        if ($InvalidChars.Count -gt 0) {
            $checks.Add([pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'Permission character unknown'
                    Description = "Supported characters are 'F', 'W', 'R', 'L', 'I' or blank."
                    Value       = "Characters: $($InvalidChars | Select-Object -Unique) -join ', '"
                })
        }
        #endregion

        #region Folder name missing
        $MissingFolders = $FolderNames.Where({ [string]::IsNullOrWhiteSpace($_.$FirstProperty) })
        if ($MissingFolders.Count -gt 0) {
            $checks.Add([pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'Folder name missing'
                    Description = 'Missing folder name in the first column.'
                    Value       = "$($MissingFolders.Count) missing folder name(s)"
                })
        }
        #endregion

        #region Duplicate folder name
        $NotUniqueFolder = $FolderNames.$FirstProperty | Group-Object | Where-Object Count -GE 2
        if ($NotUniqueFolder) {
            $checks.Add([pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'Duplicate folder name'
                    Description = 'Every folder name in the first column needs to be unique.'
                    Value       = ($NotUniqueFolder.Name) -join ', '
                })
        }
        #endregion

        #region Deepest folder has only List permissions or none at all
        $FolderRows = $Permissions | Select-Object -Skip 4
        $Paths = @($FolderRows.$FirstProperty)

        # Faster check for deepest folders
        $DeepestFolders = foreach ($P in $Paths) {
            if (-not ($Paths.Where({ $_ -ne $P -and $_ -like "$P\*" }))) {
                $P
            }
        }

        # Parent folder permissions (Row index 3)
        $ParentFolderPermissions = $Permissions[3].PSObject.Properties.Where({ 
                $_.Name -ne $FirstProperty -and -not [string]::IsNullOrWhiteSpace($_.Value) 
            }).Value

        $ParentFolderHasPermission = [bool]($ParentFolderPermissions.Where({ $_ -ne 'L' }))
        $inAccessibleFolders = [System.Collections.Generic.List[string]]::new()

        foreach ($Row in $FolderRows.Where({ $_.$FirstProperty -in $DeepestFolders })) {
            $Perms = $Row.PSObject.Properties.Where({
                    $_.Name -ne $FirstProperty -and 
                    -not [string]::IsNullOrWhiteSpace($_.Value) -and 
                    $_.Value -ne 'L'
                }).Value

            if ((-not $Perms) -and (-not $ParentFolderHasPermission)) {
                $inAccessibleFolders.Add($Row.$FirstProperty)
            }
        }

        if ($inAccessibleFolders.Count -gt 0) {
            $checks.Add([pscustomobject]@{
                    Type        = 'Warning'
                    Name        = 'Matrix design flaw'
                    Description = 'All folders need to be accessible by the end user. Please define at least (R)ead or (W)rite on the deepest folder.'
                    Value       = $inAccessibleFolders -join ', '
                })
        }
        #endregion

        # Output all collected errors at the end
        if ($checks.Count -gt 0) {
            return $checks
        }

    }
    catch {
        throw "Failed testing the Excel sheet 'Permissions' for incorrect data: $_"
    }
}

function Test-MatrixFormDataHC {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)] $FormData)

    if (-not $FormData) {
        return New-HcError `
            -Type 'Warning' `
            -Name 'FormData missing' `
            -Description 'FormData is required for specific exports.' `
            -Category 'FormData'
    }
}

function Test-MatrixSettingRowHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$SettingRow,
        [Parameter()][bool]$RequireGroupName = $false,
        [Parameter()][bool]$RequireSiteCode = $false
    )

    $checks = [System.Collections.Generic.List[pscustomobject]]::new()
    
    $validActions = @('Fix', 'New', 'Check')   

    if ([string]::IsNullOrWhiteSpace($SettingRow.Action)) {
        $checks.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Missing Action'
                Description = "The column 'Action' cannot be empty."
                Value       = $null
            })
    }
    elseif ($SettingRow.Action -notin $validActions) {
        $checks.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Invalid Action'
                Description = "Supported Action values are '$($validActions -join "', '")'."
                Value       = "Found: '$($SettingRow.Action)'"
            })
    }

    if ([string]::IsNullOrWhiteSpace($SettingRow.Path)) {
        $checks.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Missing Path'
                Description = "The column 'Path' cannot be empty."
                Value       = $null
            })
    }

    if ([string]::IsNullOrWhiteSpace($SettingRow.ComputerName)) {
        $checks.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Missing ComputerName'
                Description = "The column 'ComputerName' cannot be empty."
                Value       = $null
            })
    }

    if (
        $RequireSiteCode -and 
        [string]::IsNullOrWhiteSpace($SettingRow.SiteCode)
    ) {
        $checks.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Missing SiteCode'
                Description = "The column 'SiteCode' cannot be empty because it is used as a placeholder in the Permissions sheet."
                Value       = $null
            })
    }

    if (
        $RequireGroupName -and
        [string]::IsNullOrWhiteSpace($SettingRow.GroupName)
    ) {
        $checks.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Missing GroupName'
                Description = "The column 'GroupName' cannot be empty because it is used as a placeholder in the Permissions sheet."
                Value       = $null
            })
    } 
    

    return $checks
}

function Test-AdObjectsHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$ADObjects,
        [Parameter(Mandatory)]       $AdInfo
    )

    $checks = @()

    foreach ($obj in $ADObjects) {
        if ($obj -notin $AdInfo) {
            $checks += New-HcError `
                -Type 'Warning' `
                -Name 'Missing AD Object' `
                -Description "AD object '$obj' not found." `
                -Category 'AD'
        }
    }

    return $checks
}

function Test-ExpandedMatrixHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Matrix,
        [Parameter(Mandatory)]       $ADObject,
        [Parameter(Mandatory)]       $DefaultAcl,
        [Parameter(Mandatory)][string[]]$ExcludedSamAccountName
    )

    $checks = @()

    foreach ($row in $Matrix) {
        foreach ($ad in $row.ACL.Keys) {

            if ($ad -in $ExcludedSamAccountName) { continue }

            if ($ad -notin $ADObject) {
                $checks += New-HcError `
                    -Type 'Warning' `
                    -Name 'Unknown AD Object in ACL' `
                    -Description "Unknown AD object '$ad' in ACL." `
                    -Category 'ExpandedMatrix'
            }
        }
    }

    return $checks
}