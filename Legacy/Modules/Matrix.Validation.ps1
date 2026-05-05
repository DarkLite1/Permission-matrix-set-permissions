
function Test-DuplicateSettingsHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ImportedMatrix
    )

    $groups = $ImportedMatrix.Settings |
    Group-Object { $_.Import.ComputerName }, { $_.Import.Path } |
    Where-Object Count -GE 2

    foreach ($group in $groups) {
        foreach ($setting in $group.Group) {
            $setting.Check += New-HcError `
                -Category 'Settings' `
                -Type 'FatalError' `
                -Name 'Duplicate ComputerName/Path combination' `
                -Description 'Each ComputerName+Path must be unique across all matrices' `
                -Value @{ ComputerName = $setting.Import.ComputerName; Path = $setting.Import.Path }
        }
    }
}
function Invoke-ExpandedMatrixCheckHC {
    param(
        [Parameter(Mandatory)]
        $ImportedMatrix,
        [Parameter(Mandatory)]
        $ADObjectDetails,
        [Parameter(Mandatory)]
        $DefaultAcl,
        [Parameter(Mandatory)]
        $AdGroupPlaceHolders
    )

    foreach ($S in $ImportedMatrix.Settings) {

        if (-not $S.Matrix) { continue }

        $expandedCheck = Test-ExpandedMatrixHC `
            -Matrix $S.Matrix `
            -ADObject $ADObjectDetails `
            -DefaultAcl $DefaultAcl `
            -AdGroupPlaceHolders $AdGroupPlaceHolders

        if ($expandedCheck) {
            $S.Check += ($expandedCheck | ConvertTo-StructuredObjectHC)
        }
    }
}
class HcMatrixWorkItem {
    [string]  $MatrixFileName
    [string]  $LogFolder
    [object]  $Permissions
    [object]  $FormData
    [object[]]$Settings     # array of setting rows
    [object]  $FileMeta     # Excel info, file paths
}
function Get-AllAdIdentifiersHC {
    param([array]$ImportedMatrix)

    $names = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($I in $ImportedMatrix) {
        foreach ($S in $I.Settings) {
            if ($S.Import.GroupName) { $null = $names.Add($S.Import.GroupName) }
            if ($S.Import.SiteCode) { $null = $names.Add($S.Import.SiteCode) }
        }

        $headerSam = $I.Permissions.Import |
        Select-Object -First 3 |
        ForEach-Object { $_.P2 } |
        Where-Object { $_ }

        foreach ($n in $headerSam) { $null = $names.Add($n) }

        if ($I.FormData.Import -and $I.FormData.Import.MatrixResponsible) {
            foreach ($r in $I.FormData.Import.MatrixResponsible.Split(',')) {
                $null = $names.Add($r.Trim())
            }
        }
    }

    return $names.ToArray()
}


