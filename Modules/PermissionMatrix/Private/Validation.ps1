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
    param([Parameter(Mandatory)] $Objects)

    $output = @()

    foreach ($obj in @($Objects)) {
        if ($null -eq $obj) { continue }

        if ($obj -is [string]) {
            $output += New-HcError -Type 'Information' -Name 'Message' -Description $obj
            continue
        }

        if ($obj -is [hashtable] -or $obj -is [pscustomobject]) {
            $output += $obj
            continue
        }

        $output += New-HcError -Type 'Information' -Name 'UnknownObject' -Description "$obj"
    }

    return $output
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
    [CmdletBinding()]
    param([Parameter(Mandatory)][array]$Permissions)

    if ($Permissions.Count -lt 4) {
        return New-HcError -Type 'FatalError' -Name 'Invalid Permissions Sheet' `
            -Description 'Permissions sheet must contain at least 4 rows.' `
            -Category 'Permissions'
    }
}

function Test-MatrixFormDataHC {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $FormData)

    if (-not $FormData) {
        return New-HcError -Type 'Warning' -Name 'FormData missing' `
            -Description 'FormData is required for specific exports.' `
            -Category 'FormData'
    }
}

function Test-MatrixSettingHC {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Setting)

    $checks = @()

    if (-not $Setting.ComputerName) {
        $checks += New-HcError -Type 'FatalError' -Name 'Missing ComputerName' `
            -Description 'ComputerName is mandatory.' -Category 'Settings'
    }
    if (-not $Setting.Path) {
        $checks += New-HcError -Type 'FatalError' -Name 'Missing Path' `
            -Description 'Path is mandatory.' -Category 'Settings'
    }
    if (-not $Setting.Action) {
        $checks += New-HcError -Type 'FatalError' -Name 'Missing Action' `
            -Description 'Action is mandatory.' -Category 'Settings'
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