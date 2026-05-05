function Merge-DefaultPermissionsHC {
    <#
    .SYNOPSIS
        Applies default permissions based on a strict toggle setting.
    .DESCRIPTION
        - If ApplyDefaultPermissions is $false: Returns only the matrix permissions.
        - If ApplyDefaultPermissions is $true: Checks for overlap. If an AD Object exists in 
          both, it throws a terminating error. If no overlap, defaults are appended.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [array]$Defaults = @(),

        [Parameter(Mandatory = $false)]
        [array]$Matrix = @(),

        [Parameter(Mandatory)]
        [bool]$ApplyDefaultPermissions
    )

    # Rule 1: If the setting is disabled, the Matrix manages everything alone.
    if (-not $ApplyDefaultPermissions) {
        return @($Matrix)
    }

    # Rule 2: If enabled, check for conflicts.
    $matrixObjects = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $Matrix) { 
        $null = $matrixObjects.Add($m.ADObject) 
    }

    $conflicts = [System.Collections.Generic.List[string]]::new()
    foreach ($d in $Defaults) {
        if ($matrixObjects.Contains($d.ADObject)) {
            $conflicts.Add($d.ADObject)
        }
    }

    if ($conflicts.Count -gt 0) {
        throw "Defaults conflict detected. The following AD Objects are defined in both the Matrix and Defaults: $($conflicts -join ', ')"
    }

    # Rule 3: No conflicts exist, safely append the defaults to the matrix!
    $mergedList = [System.Collections.Generic.List[pscustomobject]]::new()
    if ($Matrix) { $mergedList.AddRange([pscustomobject[]]$Matrix) }
    if ($Defaults) { $mergedList.AddRange([pscustomobject[]]$Defaults) }

    return $mergedList.ToArray()
}