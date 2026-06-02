function Get-ADObjectDetailHC {
    <#
    .SYNOPSIS
        High-speed, parallel retrieval of Active Directory object details and 
        recursive group memberships.

    .DESCRIPTION
        Retrieves fundamental properties (SID, DistinguishedName, Manager, 
        etc.) for an array of Active Directory objects. 
        
        If the targeted object is a Group, the script performs a recursive 
        expansion (extracting users from nested child groups) and stores the 
        results in the 'adGroupMember' property. Note: The 'Domain Users' group 
        is explicitly bypassed for performance and safety reasons.

        Performance Features:
        1. Multi-Threading: 
            Processes the input array concurrently using PowerShell 7's 
            Parallel runspaces.
        2. Tiered Lookups: 
            Searches the local domain first for maximum speed. It only falls 
            back to the Forest Global Catalog (GC) when an object isn't found 
            locally, ensuring cross-domain trust lookups succeed without 
            penalizing the speed of the common case.

    .PARAMETER ADObjectName
        An array of Active Directory object names (users or groups) to resolve.

    .PARAMETER Type
        The format of the strings provided in the ADObjectName parameter. 
        Valid values: 
        'SamAccountName' (e.g., 'jdoe' or 'DOMAIN\jdoe') or 'DistinguishedName'.

    .PARAMETER MaxThreads
        The maximum number of concurrent Active Directory LDAP queries to 
        execute. (Default: 7)

    .EXAMPLE
        $targets = @('HR_Managers', 'DOMAIN\jdoe', 'asmith@domain.com')
        $adDetails = Get-ADObjectDetailHC `
            -ADObjectName $targets `
            -Type 'SamAccountName' `
            -MaxThreads 10

        # View the recursive members of the HR_Managers group
        $adDetails | Where-Object SamAccountName -eq 'HR_Managers' | 
        Select-Object -ExpandProperty adGroupMember
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
            'managedby', 'name', 'objectclass', 'objectsid'
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
                    ObjectSid         = if ($props['objectsid'].Count) {
                        # objectSid comes back as a byte[] from the searcher; convert to S-1-5-... string form
                        [System.Security.Principal.SecurityIdentifier]::new(
                            [byte[]]$props['objectsid'][0], 0
                        ).Value
                    }
                    else { $null }
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
        Converts a list of e-mail addresses or SamAccountNames into a unique 
        array of Active Directory UserPrincipalNames.

    .DESCRIPTION
        Resolves a mixed array of user and group identifiers against Active 
        Directory. 
        
        If a group is detected, the script recursively expands its membership 
        to find all nested users. The final output is strictly filtered: only 
        AD users that are currently Enabled, have a populated 'Mail' attribute, 
        and are not explicitly excluded will be returned.

        Returns a hashtable containing two arrays:
        - 'notFound': Inputs that could not be matched to an AD object.
        - 'userPrincipalName': The deduplicated list of resolved UPNs.

    .PARAMETER Name
        An array of strings. Can be a standard e-mail address (matched against ProxyAddresses) or a SamAccountName of a user or group object in AD.

    .PARAMETER ExcludeSamAccountName
        An optional array of SamAccountNames to explicitly ignore. This is highly useful for stripping service accounts or administrator accounts out of expanded group memberships.

    .EXAMPLE
        $targets = @('HR_Team@domain.com', 'jdoe')
        $exclusions = @('svc_hr_scanner', 'admin_jdoe')
        
        $result = Get-AdUserPrincipalNameHC -Name $targets -ExcludeSamAccountName $exclusions
        
        Write-Host "Found UPNs: $($result.userPrincipalName -join ', ')"
        Write-Host "Unresolved: $($result.notFound -join ', ')"
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