function Get-ADObjectDetailHC {
    <#
    .SYNOPSIS
        Retrieve details about an AD object.

    .DESCRIPTION
        Retrieve details about an AD object. If the object is not found the
        property 'adObject' is blank. If it is a group, the group members are
        retrieved and stored in the property 'adGroupMember'.

        Searches the local domain first for speed; falls back to the forest
        Global Catalog only when an object isn't found locally, so cross-domain
        lookups still resolve without slowing down the common case.

    .PARAMETER ADObjectName
        Name of the user or group objects to search for.

    .PARAMETER Type
        The type of strings passed to ADObjectName.
        Valid values: SamAccountName, DistinguishedName

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

    # Resolve the GC path once on the calling thread. Used only as a fallback
    # when the local-domain search misses, so the cost is paid rarely.
    if (-not $script:CachedGcPath) {
        $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
        $script:CachedGcPath = "GC://$($forest.RootDomain.Name)"
    }
    $gcPath = $script:CachedGcPath

    $ADObjectName = $ADObjectName | Sort-Object -Unique

    $ADObjectName | ForEach-Object -ThrottleLimit $MaxThreads -Parallel {
        $propertyType = $using:Type
        $gcPath = $using:gcPath
        $name = $_

        # Local helper: build the LDAP filter for a given input name.
        # Handles bare names, DOMAIN\user, and user@domain UPNs.
        function Get-Filter {
            param($n, $t)
            if ($t -eq 'DistinguishedName') {
                return "(distinguishedName=$n)"
            }
            if ($n -match '@') {
                return "(|(samAccountName=$n)(userPrincipalName=$n))"
            }
            if ($n -match '\\') {
                $clean = $n -replace '^.*\\', ''
                return "(samAccountName=$clean)"
            }
            return "(samAccountName=$n)"
        }

        $propertiesToLoad = @(
            'distinguishedname', 'samaccountname', 'manager',
            'managedby', 'name', 'objectclass'
        )

        $searcher = $null
        $root = $null
        $searchResult = $null

        try {
            Write-Verbose "Get AD details for '$name'"

            # Fast path: local-domain searcher (no path = nearest DC via DC locator)
            $searcher = [System.DirectoryServices.DirectorySearcher]::new()
            $searcher.Filter = Get-Filter $name $propertyType
            $null = $searcher.PropertiesToLoad.AddRange($propertiesToLoad)
            $searchResult = $searcher.FindOne()
            $searcher.Dispose()
            $searcher = $null

            # Fallback: hit the forest GC only when the local search missed.
            # This is where cross-domain trust resolution actually happens.
            if (-not $searchResult) {
                Write-Verbose "Not found locally, trying GC for '$name'"
                $root = [System.DirectoryServices.DirectoryEntry]::new($gcPath)
                $searcher = [System.DirectoryServices.DirectorySearcher]::new($root)
                $searcher.Filter = Get-Filter $name $propertyType
                $null = $searcher.PropertiesToLoad.AddRange($propertiesToLoad)
                $searchResult = $searcher.FindOne()
            }

            $adObject = $null
            $adGroupMember = $null

            if ($searchResult) {
                $props = $searchResult.Properties
                $isGroup = $props['objectclass'] -contains 'group'

                $adObject = [PSCustomObject]@{
                    DistinguishedName = if ($props['distinguishedname'].Count) { $props['distinguishedname'][0] } else { $null }
                    SamAccountName    = if ($props['samaccountname'].Count) { $props['samaccountname'][0] } else { $null }
                    ManagedBy         = if ($props['manager'].Count) { $props['manager'][0] }
                    elseif ($props['managedby'].Count) { $props['managedby'][0] }
                    else { $null }
                    Name              = if ($props['name'].Count) { $props['name'][0] } else { $null }
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
                        # Build PrincipalContext against the group's own domain
                        # (parsed from its DN) so cross-domain group expansion works.
                        $groupDn = $adObject.DistinguishedName
                        $groupDomain = ($groupDn -split ',' |
                            Where-Object { $_ -like 'DC=*' } |
                            ForEach-Object { $_.Substring(3) }) -join '.'

                        $ctx = [System.DirectoryServices.AccountManagement.PrincipalContext]::new(
                            [System.DirectoryServices.AccountManagement.ContextType]::Domain,
                            $groupDomain
                        )
                        try {
                            $group = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity(
                                $ctx,
                                [System.DirectoryServices.AccountManagement.IdentityType]::DistinguishedName,
                                $groupDn
                            )
                            if ($group) {
                                try {
                                    $adGroupMember = foreach ($m in $group.GetMembers($true)) {
                                        [PSCustomObject]@{
                                            objectClass       = $m.StructuralObjectClass
                                            Name              = $m.Name
                                            SamAccountName    = $m.SamAccountName
                                            DistinguishedName = $m.DistinguishedName
                                        }
                                    }
                                }
                                finally { $group.Dispose() }
                            }
                        }
                        finally { $ctx.Dispose() }
                    }
                }
            }

            $returnObj = [PSCustomObject]@{
                adObject      = $adObject
                adGroupMember = $adGroupMember
            }
            $returnObj | Add-Member -MemberType NoteProperty -Name $propertyType -Value $name
            return $returnObj
        }
        catch {
            Write-Warning "Failed retrieving AD object details for '$name': $_"
            $returnObj = [PSCustomObject]@{
                adObject      = $null
                adGroupMember = $null
            }
            $returnObj | Add-Member -MemberType NoteProperty -Name $propertyType -Value $name
            return $returnObj
        }
        finally {
            if ($searcher) { $searcher.Dispose() }
            if ($root) { $root.Dispose() }
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

function Get-ADObjectDetailHC {
    <#
    .SYNOPSIS
        Retrieve details about an AD object.

    .DESCRIPTION
        Retrieve details about an AD object. If the object is not found the
        property 'adObject' is blank. If it is a group, the group members are
        retrieved and stored in the property 'adGroupMember'.
        This function natively supports down-level logons (DOMAIN\User) and UPNs.

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

        # Ensure required assemblies are loaded in the parallel runspace
        Add-Type -AssemblyName 'System.DirectoryServices.AccountManagement'
        Add-Type -AssemblyName 'System.DirectoryServices'

        try {
            Write-Verbose "Get AD details for '$name'"

            $adObject = $null
            $adGroupMember = $null

            $ctx = [System.DirectoryServices.AccountManagement.PrincipalContext]::new([System.DirectoryServices.AccountManagement.ContextType]::Domain)

            # Search for the principal.
            # If Type is DistinguishedName, enforce that type. 
            # Otherwise, allow FindByIdentity to flexibly parse DOMAIN\User, User@Domain.com, or SamAccountName.
            if ($propertyType -eq 'DistinguishedName') {
                $principal = [System.DirectoryServices.AccountManagement.Principal]::FindByIdentity($ctx, [System.DirectoryServices.AccountManagement.IdentityType]::DistinguishedName, $name)
            } 
            else {
                $principal = [System.DirectoryServices.AccountManagement.Principal]::FindByIdentity($ctx, $name)
            }

            if ($principal) {
                $isGroup = $principal -is [System.DirectoryServices.AccountManagement.GroupPrincipal]

                # Retrieve ManagedBy/Manager from the underlying DirectoryEntry
                $de = $principal.GetUnderlyingObject() -as [System.DirectoryServices.DirectoryEntry]
                $managedBy = $null
                
                if ($de) {
                    if ($de.Properties.Contains('manager')) {
                        $managedBy = $de.Properties['manager'][0]
                    } 
                    elseif ($de.Properties.Contains('managedby')) {
                        $managedBy = $de.Properties['managedby'][0]
                    }
                }

                $adObject = [PSCustomObject]@{
                    DistinguishedName = $principal.DistinguishedName
                    SamAccountName    = $principal.SamAccountName
                    ManagedBy         = $managedBy
                    Name              = $principal.Name
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
                        $adGroupMember = foreach ($m in $principal.GetMembers($true)) {
                            [PSCustomObject]@{
                                objectClass       = $m.StructuralObjectClass
                                Name              = $m.Name
                                SamAccountName    = $m.SamAccountName
                                DistinguishedName = $m.DistinguishedName
                            }
                        }
                    }
                }
                $principal.Dispose()
            }
            $ctx.Dispose()

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