function Get-ADObjectDetailHC {
    <#
    .SYNOPSIS
        Retrieve details about an AD object.

    .DESCRIPTION
        Retrieve details about an AD object. If the object is not found the
        property 'adObject' is blank. If it is a group, the group members are
        retrieved and stored in the property 'adGroupMember'.

    .PARAMETER ADObjectName
        Name of the user or group objects to search for.

    .PARAMETER Type
        The type of strings passed to ADObjectName.

        Valid values are:
        - DistinguishedName
        - SamAccountName

    .PARAMETER MaxThreads
        Maximum concurrent AD requests.
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [Parameter(Mandatory)]
        [String[]]$ADObjectName,
        [Parameter(Mandatory)]
        [ValidateSet('SamAccountName', 'DistinguishedName')]
        [String]$Type,
        [Int]$MaxThreads = 7
    )

    $ADObjectName = $ADObjectName | Sort-Object -Unique

    $ADObjectName | ForEach-Object -ThrottleLimit $MaxThreads -Parallel {
        $propertyType = $using:Type
        $name = $_

        try {
            Write-Verbose "Get AD details for '$name'"

            #region Get AD object details
            $searcher = [System.DirectoryServices.DirectorySearcher]::new()

            if ($propertyType -eq 'SamAccountName') {
                $searcher.Filter = "(samaccountname=$name)"
            }
            else {
                $searcher.Filter = "(distinguishedname=$name)"
            }

            $searcher.PropertiesToLoad.AddRange(
                @(
                    'distinguishedname',
                    'samaccountname',
                    'manager',
                    'managedby',
                    'name',
                    'objectclass'
                )
            )

            $searchResult = $searcher.FindOne()
            $adObject = $null
            $adGroupMember = $null

            if ($searchResult) {
                $props = $searchResult.Properties
                $isGroup = $props['objectclass'] -contains 'group'

                $adObject = [PSCustomObject]@{
                    DistinguishedName = if ($props['distinguishedname']) { $props['distinguishedname'][0] } else { $null }
                    SamAccountName    = if ($props['samaccountname']) { $props['samaccountname'][0] } else { $null }
                    ManagedBy         = if ($props['manager']) { $props['manager'][0] } elseif ($props['managedby']) { $props['managedby'][0] } else { $null }
                    Name              = if ($props['name']) { $props['name'][0] } else { $null }
                    ObjectClass       = if ($isGroup) { 'group' } else { 'user' }
                }

                if ($isGroup) {
                    if ($adObject.Name -eq 'Domain Users') {
                        $adGroupMember = @([PSCustomObject]@{
                                objectClass       = 'user'
                                Name              = 'All users'
                                SamAccountName    = 'All users'
                                DistinguishedName = $null
                            })
                    }
                    else {
                        # Safely scope and dispose of memory-heavy classes
                        $ctx = [System.DirectoryServices.AccountManagement.PrincipalContext]::new([System.DirectoryServices.AccountManagement.ContextType]::Domain)

                        $group = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($ctx, $adObject.DistinguishedName)

                        if ($group) {
                            $adGroupMember = foreach ($m in $group.GetMembers($true)) {
                                [PSCustomObject]@{
                                    objectClass       = $m.StructuralObjectClass
                                    Name              = $m.Name
                                    SamAccountName    = $m.SamAccountName
                                    DistinguishedName = $m.DistinguishedName
                                }
                            }
                            $group.Dispose()
                        }
                        $ctx.Dispose()
                    }
                }
            }
            #endregion

            $returnObj = [PSCustomObject]@{
                adObject      = $adObject
                adGroupMember = $adGroupMember
            }
            $returnObj | Add-Member -MemberType NoteProperty -Name $propertyType -Value $name

            return $returnObj
        }
        catch {
            $M = $_
            Write-Warning "Failed retrieving AD object details for '$name': $M"
            
            # Return the blank object so the orchestrator can continue processing other accounts safely
            $returnObj = [PSCustomObject]@{
                adObject      = $null
                adGroupMember = $null
            }
            $returnObj | Add-Member -MemberType NoteProperty -Name $propertyType -Value $name

            return $returnObj
        }
    }

    Write-Verbose 'All AD object details retrieved'
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