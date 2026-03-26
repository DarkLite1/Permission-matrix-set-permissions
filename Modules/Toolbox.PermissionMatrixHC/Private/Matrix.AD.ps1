function Test-AdObjectsHC {
    <#
    .SYNOPSIS
        Verify AD Objects used in the matrix headers.

    .DESCRIPTION
        Checks for duplicate or missing AD object names in the generated header 
        row objects to prevent matrix execution failures.
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [hashtable]$ADObjects
    )

    process {
        try {
            $SamAccountNames = $ADObjects.Values.SamAccountName

            #region Duplicate AD Objects
            $NotUniqueADObjects = $SamAccountNames | Group-Object | Where-Object { $_.Count -ge 2 }

            if ($NotUniqueADObjects) {
                [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'AD Object not unique'
                    Description = "All objects defined in the matrix need to be unique. Duplicate AD Objects can also be generated from the 'Settings' worksheet combined with the header rows in the 'Permissions' worksheet."
                    Value       = $NotUniqueADObjects.Name | Sort-Object
                }
            }
            #endregion

            #region AD Object name missing
            $HasMissingName = $false
            
            # Loop through names and break instantly if we find a blank one
            foreach ($Name in $SamAccountNames) {
                if ([string]::IsNullOrWhiteSpace($Name)) {
                    $HasMissingName = $true
                    break
                }
            }

            if ($HasMissingName) {
                [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'AD Object name missing'
                    Description = "Every column in the worksheet 'Permissions' needs to have an AD object name in the header row. The AD object name can not be blank."
                    Value       = $null
                }
            }
            #endregion
        }
        catch {
            throw "Failed testing AD object names: $_"
        }
    }
}
