#requires -Modules Pester

#
# ActiveDirectory.ps1 INTEGRATION tests
#
# These run against the LIVE Active Directory of the machine executing them.
# Per project decision they are integration tests, not unit tests:
#
#   * Get-ADObjectDetailHC uses [DirectorySearcher]::new() and the static
#     [GroupPrincipal]::FindByIdentity(...) inside ForEach-Object -Parallel —
#     none of which Pester can mock.
#   * Get-AdUserPrincipalNameHC passes/pipes objects into the real Get-ADObject
#     / Get-ADGroupMember / Get-ADUser cmdlets, whose -Identity binding and
#     validation run BEFORE any mock body, so fabricated objects can't be used.
#
# Real directory objects satisfy that binding, so the suite auto-discovers a
# qualifying user and group at run time and asserts on invariants (shape,
# class, enabled/mail filtering) rather than hard-coded identities, since the
# concrete objects differ per domain.
#
# REQUIREMENT: a reachable domain controller and the ActiveDirectory module.
# Per decision, the suite HARD-FAILS (does not skip) when AD is unavailable or
# when no qualifying objects can be found — an environment that can't run these
# is treated as a broken test environment, not a pass.
#

Describe 'ActiveDirectory.ps1 - AD Lookup Functions (integration)' {

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\ActiveDirectory.ps1"

        # ---- Hard requirement: ActiveDirectory module + reachable domain ----
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            throw 'Integration tests require the ActiveDirectory module, which is not installed on this runner.'
        }
        Import-Module ActiveDirectory -ErrorAction Stop

        try {
            $script:Domain = Get-ADDomain -ErrorAction Stop
        }
        catch {
            throw "Integration tests require a reachable Active Directory domain. Get-ADDomain failed: $_"
        }

        # ---- Auto-discover a qualifying USER ----
        # Qualifying = enabled, has a Mail value, and has a UserPrincipalName,
        # because Get-AdUserPrincipalNameHC filters on exactly those.
        $script:TestUser = Get-ADUser -ResultSetSize 25 `
            -Filter "Enabled -eq 'True' -and Mail -like '*' -and UserPrincipalName -like '*'" `
            -Properties Mail, UserPrincipalName, Enabled |
        Where-Object { $_.Mail -and $_.UserPrincipalName } |
        Select-Object -First 1

        if (-not $script:TestUser) {
            throw 'Integration tests require at least one enabled, mail-enabled user in the domain; none was found.'
        }

        # ---- Auto-discover a qualifying GROUP ----
        # Qualifying = a group with at least one enabled, mail-enabled member
        # (so the expansion assertions are not vacuous). We also capture that
        # member set's UPNs to assert against.
        $script:TestGroup = $null
        $script:TestGroupUpns = @()

        $candidateGroups = Get-ADGroup -ResultSetSize 50 -Filter "Mail -like '*'" -Properties Mail
        if (-not $candidateGroups) {
            # Fall back to any group if none are mail-enabled; the function keys
            # off the group object, not its Mail, for expansion.
            $candidateGroups = Get-ADGroup -ResultSetSize 50 -Filter *
        }

        foreach ($g in $candidateGroups) {
            $members = @(
                Get-ADGroupMember -Identity $g -Recursive -ErrorAction SilentlyContinue |
                Get-ADUser -Properties Enabled, Mail, UserPrincipalName -ErrorAction SilentlyContinue |
                Where-Object { $_.Enabled -and $_.Mail -and $_.UserPrincipalName }
            )

            if ($members.Count -gt 0) {
                $script:TestGroup = $g
                $script:TestGroupUpns = @($members.UserPrincipalName | Sort-Object -Unique)
                break
            }
        }

        if (-not $script:TestGroup) {
            throw 'Integration tests require a group with at least one enabled, mail-enabled member; none was found.'
        }

        # A name guaranteed not to resolve, for the not-found paths.
        $script:BogusName = "zzz-no-such-object-$([guid]::NewGuid().Guid)@example.invalid"
    }

    Context 'Get-ADObjectDetailHC - parameter contract' {

        It 'requires ADObjectName' {
            (Get-Command Get-ADObjectDetailHC).Parameters['ADObjectName'].Attributes.Mandatory |
            Should -Contain $true
        }

        It 'requires Type' {
            (Get-Command Get-ADObjectDetailHC).Parameters['Type'].Attributes.Mandatory |
            Should -Contain $true
        }

        It 'restricts Type to SamAccountName or DistinguishedName' {
            $validate = (Get-Command Get-ADObjectDetailHC).Parameters['Type'].Attributes |
            Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }

            $validate.ValidValues | Should -Contain 'SamAccountName'
            $validate.ValidValues | Should -Contain 'DistinguishedName'
            $validate.ValidValues.Count | Should -Be 2
        }

        It 'rejects an invalid Type value' {
            { Get-ADObjectDetailHC -ADObjectName 'x' -Type 'NotAValidType' -MaxThreads 1 } |
            Should -Throw
        }
    }

    Context 'Get-ADObjectDetailHC - user lookup' {

        It 'resolves a real user by SamAccountName' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestUser.SamAccountName -Type SamAccountName -MaxThreads 1

            $res | Should -Not -BeNullOrEmpty
            $res[0].adObject | Should -Not -BeNullOrEmpty
            $res[0].adObject.ObjectClass | Should -Be 'user'
            $res[0].adObject.SamAccountName | Should -Be $TestUser.SamAccountName
            $res[0].adObject.DistinguishedName | Should -Not -BeNullOrEmpty
        }

        It 'does not populate adGroupMember for a user' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestUser.SamAccountName -Type SamAccountName -MaxThreads 1

            $res[0].adGroupMember | Should -BeNullOrEmpty
        }

        It 'echoes the input back on the dynamic SamAccountName property' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestUser.SamAccountName -Type SamAccountName -MaxThreads 1

            $res[0].SamAccountName | Should -Be $TestUser.SamAccountName
        }

        It 'resolves the same user by DistinguishedName' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestUser.DistinguishedName -Type DistinguishedName -MaxThreads 1

            $res[0].adObject | Should -Not -BeNullOrEmpty
            $res[0].adObject.ObjectClass | Should -Be 'user'
            $res[0].adObject.DistinguishedName | Should -Be $TestUser.DistinguishedName
        }

        It 'returns a null adObject for a name that does not exist' {
            $res = Get-ADObjectDetailHC -ADObjectName 'zzz-no-such-sam-acct-xyzzy' -Type SamAccountName -MaxThreads 1

            $res[0].adObject | Should -BeNullOrEmpty
            $res[0].adGroupMember | Should -BeNullOrEmpty
        }
    }

    Context 'Get-ADObjectDetailHC - group lookup & expansion' {

        It 'resolves a real group and classifies it as a group' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestGroup.SamAccountName -Type SamAccountName -MaxThreads 1

            $res[0].adObject | Should -Not -BeNullOrEmpty
            $res[0].adObject.ObjectClass | Should -Be 'group'
        }

        It 'expands group members with the documented shape' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestGroup.SamAccountName -Type SamAccountName -MaxThreads 1

            $res[0].adGroupMember | Should -Not -BeNullOrEmpty
            foreach ($m in $res[0].adGroupMember) {
                $m.PSObject.Properties.Name | Should -Contain 'objectClass'
                $m.PSObject.Properties.Name | Should -Contain 'Name'
                $m.PSObject.Properties.Name | Should -Contain 'SamAccountName'
                $m.PSObject.Properties.Name | Should -Contain 'DistinguishedName'
            }
        }

        It 'applies the Domain Users special-case' {
            # 'Domain Users' is a well-known group present in every domain; the
            # function short-circuits its expansion to a single synthetic entry.
            $res = Get-ADObjectDetailHC -ADObjectName 'Domain Users' -Type SamAccountName -MaxThreads 1

            $res[0].adObject.ObjectClass | Should -Be 'group'
            $res[0].adObject.Name | Should -Be 'Domain Users'
            @($res[0].adGroupMember).Count | Should -Be 1
            $res[0].adGroupMember[0].Name | Should -Be 'All users'
            $res[0].adGroupMember[0].SamAccountName | Should -Be 'All users'
        }
    }

    Context 'Get-AdUserPrincipalNameHC - user resolution' {

        It 'returns the UPN for a real user mail address' {
            $res = Get-AdUserPrincipalNameHC -Name $TestUser.Mail

            $res.userPrincipalName | Should -Contain $TestUser.UserPrincipalName
            $res.notFound.Count | Should -Be 0
        }

        It 'returns the UPN for a real user SamAccountName' {
            $res = Get-AdUserPrincipalNameHC -Name $TestUser.SamAccountName

            $res.userPrincipalName | Should -Contain $TestUser.UserPrincipalName
        }

        It 'adds an unresolvable name to notFound' {
            $res = Get-AdUserPrincipalNameHC -Name $BogusName

            $res.notFound | Should -Contain $BogusName
            $res.userPrincipalName.Count | Should -Be 0
        }

        It 'excludes a user listed in -ExcludeSamAccountName' {
            $res = Get-AdUserPrincipalNameHC -Name $TestUser.Mail `
                -ExcludeSamAccountName $TestUser.SamAccountName

            $res.userPrincipalName | Should -Not -Contain $TestUser.UserPrincipalName
        }

        It 'resolves the matched name while still flagging an unmatched one' {
            $res = Get-AdUserPrincipalNameHC -Name @($TestUser.Mail, $BogusName)

            $res.userPrincipalName | Should -Contain $TestUser.UserPrincipalName
            $res.notFound | Should -Contain $BogusName
        }
    }

    Context 'Get-AdUserPrincipalNameHC - group expansion' {

        It 'expands a group to the UPNs of its enabled, mail-enabled members' {
            $res = Get-AdUserPrincipalNameHC -Name $TestGroup.SamAccountName

            $res.userPrincipalName.Count | Should -BeGreaterThan 0
            foreach ($upn in $TestGroupUpns) {
                $res.userPrincipalName | Should -Contain $upn
            }
        }

        It 'only returns UPNs of enabled, mail-bearing users' {
            $res = Get-AdUserPrincipalNameHC -Name $TestGroup.SamAccountName

            foreach ($upn in $res.userPrincipalName) {
                $u = Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -Properties Enabled, Mail
                $u | Should -Not -BeNullOrEmpty
                $u.Enabled | Should -BeTrue
                $u.Mail | Should -Not -BeNullOrEmpty
            }
        }

        It 'returns a de-duplicated UPN list' {
            $res = Get-AdUserPrincipalNameHC -Name $TestGroup.SamAccountName

            ($res.userPrincipalName | Sort-Object -Unique).Count |
            Should -Be $res.userPrincipalName.Count
        }
    }

    Context 'Get-AdUserPrincipalNameHC - contract' {

        It 'returns a hashtable exposing notFound and userPrincipalName' {
            $res = Get-AdUserPrincipalNameHC -Name $BogusName

            $res | Should -BeOfType [hashtable]
            $res.Keys | Should -Contain 'notFound'
            $res.Keys | Should -Contain 'userPrincipalName'
        }
    }
}