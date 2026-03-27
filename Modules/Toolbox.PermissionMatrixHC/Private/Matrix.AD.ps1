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
function Get-AdUserPrincipalNameHC {
    <#
    .SYNOPSIS
        Convert a list of e-mail addresses to a list of UserPrincipalNames.

    .DESCRIPTION
        The list to convert can contain user e-mail addresses or group e-mail
        addresses. For groups the user members are retrieved. The result will
        only contain UserPrincipalNames from AD user accounts that are enabled.

    .PARAMETER Name
        Can be an e-mail address or a SamAccountName of a user object or a
        group object in AD.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [String[]]$Name,
        
        [String[]]$ExcludeSamAccountName = @()
    )

    process {
        try {
            # Use Generic Lists for massive speed improvements over +=
            $NotFound = [System.Collections.Generic.List[string]]::new()
            $UpnList = [System.Collections.Generic.List[string]]::new()

            $UniqueNames = $Name | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

            foreach ($N in $UniqueNames) {
                # Find the object in AD
                $AdObject = Get-ADObject -Filter "ProxyAddresses -eq 'smtp:$N' -or SamAccountName -eq '$N'" -Property 'Mail'

                if ($AdObject.Count -ge 2) {
                    throw "Multiple results found for name '$N': $($AdObject.Name -join ', '). Skipping."
                    
                }

                if (-not $AdObject) {
                    $NotFound.Add($N)
                    continue
                }

                $AdUsers = $null

                if ($AdObject.ObjectClass -eq 'group') {
                    $AdUsers = Get-ADGroupMember -Identity $AdObject -Recursive
                }
                elseif ($AdObject.ObjectClass -eq 'user') {
                    $AdUsers = $AdObject
                }

                if ($AdUsers) {
                    $AdUsers | Get-ADUser -Properties Enabled, Mail -ErrorAction SilentlyContinue | 
                    ForEach-Object {
                        if ($_.Mail -and $_.Enabled -and $_.SamAccountName -notin $ExcludeSamAccountName) {
                            if (-not [string]::IsNullOrWhiteSpace($_.UserPrincipalName)) {
                                $UpnList.Add($_.UserPrincipalName)
                            }
                        }
                    }
                }
            }

            return @{
                notFound          = @($NotFound)
                userPrincipalName = @($UpnList | Sort-Object -Unique)
            }
        }
        catch {
            throw "Failed converting email address or SamAccountName to userPrincipalName: $_"
        }
    }
}