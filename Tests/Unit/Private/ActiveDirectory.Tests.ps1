#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

#
# ActiveDirectory.ps1 INTEGRATION tests
#
# These run against the LIVE Active Directory of the machine executing them.
# Per project decision they are integration tests, not unit tests:
#
#   * Get-ADObjectDetailHC uses [DirectorySearcher]::new() and the static
#     [GroupPrincipal]::FindByIdentity(...) inside ForEach-Object -Parallel —
#     none of which Pester can mock.
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

        Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
        ForEach-Object { . $_.FullName }

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
        # Stricter than the Get-ADObjectDetailHC tests below strictly need
        # (they use SamAccountName and DistinguishedName only), but kept so the
        # same discovered user stays usable if mail-based tests return.
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
    }

    Context 'Get-ADObjectDetailHC - parameter contract' {

        It 'requires ADObjectName' {
            (Get-Command Get-ADObjectDetailHC).Parameters['ADObjectName'].Attributes.Mandatory | Should-ContainCollection $true
        }

        It 'requires Type' {
            (Get-Command Get-ADObjectDetailHC).Parameters['Type'].Attributes.Mandatory | Should-ContainCollection $true
        }

        It 'restricts Type to SamAccountName or DistinguishedName' {
            $validate = (Get-Command Get-ADObjectDetailHC).Parameters['Type'].Attributes |
            Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }

            $validate.ValidValues | Should-ContainCollection 'SamAccountName'
            $validate.ValidValues | Should-ContainCollection 'DistinguishedName'
            $validate.ValidValues.Count | Should-Be 2
        }

        It 'rejects an invalid Type value' {
            { Get-ADObjectDetailHC -ADObjectName 'x' -Type 'NotAValidType' -MaxThreads 1 } | Should-Throw
        }
    }

    Context 'Get-ADObjectDetailHC - user lookup' {

        It 'resolves a real user by SamAccountName' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestUser.SamAccountName -Type SamAccountName -MaxThreads 1

            $res | Should-BeTruthy
            $res[0].adObject | Should-BeTruthy
            $res[0].adObject.ObjectClass | Should-Be 'user'
            $res[0].adObject.SamAccountName | Should-Be $TestUser.SamAccountName
            $res[0].adObject.DistinguishedName | Should-BeTruthy
        }

        It 'does not populate adGroupMember for a user' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestUser.SamAccountName -Type SamAccountName -MaxThreads 1

            $res[0].adGroupMember | Should-BeFalsy
        }

        It 'echoes the input back on the dynamic SamAccountName property' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestUser.SamAccountName -Type SamAccountName -MaxThreads 1

            $res[0].SamAccountName | Should-Be $TestUser.SamAccountName
        }

        It 'resolves the same user by DistinguishedName' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestUser.DistinguishedName -Type DistinguishedName -MaxThreads 1

            $res[0].adObject | Should-BeTruthy
            $res[0].adObject.ObjectClass | Should-Be 'user'
            $res[0].adObject.DistinguishedName | Should-Be $TestUser.DistinguishedName
        }

        It 'returns a null adObject for a name that does not exist' {
            $res = Get-ADObjectDetailHC -ADObjectName 'zzz-no-such-sam-acct-xyzzy' -Type SamAccountName -MaxThreads 1

            $res[0].adObject | Should-BeFalsy
            $res[0].adGroupMember | Should-BeFalsy
        }
    }

    Context 'Get-ADObjectDetailHC - group lookup & expansion' {

        It 'resolves a real group and classifies it as a group' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestGroup.SamAccountName -Type SamAccountName -MaxThreads 1

            $res[0].adObject | Should-BeTruthy
            $res[0].adObject.ObjectClass | Should-Be 'group'
        }

        It 'expands group members with the documented shape' {
            $res = Get-ADObjectDetailHC -ADObjectName $TestGroup.SamAccountName -Type SamAccountName -MaxThreads 1

            $res[0].adGroupMember | Should-BeTruthy
            foreach ($m in $res[0].adGroupMember) {
                $m.PSObject.Properties.Name | Should-ContainCollection 'objectClass'
                $m.PSObject.Properties.Name | Should-ContainCollection 'Name'
                $m.PSObject.Properties.Name | Should-ContainCollection 'SamAccountName'
                $m.PSObject.Properties.Name | Should-ContainCollection 'DistinguishedName'
            }
        }

        It 'applies the Domain Users special-case' {
            # 'Domain Users' is a well-known group present in every domain; the
            # function short-circuits its expansion to a single synthetic entry.
            $res = Get-ADObjectDetailHC -ADObjectName 'Domain Users' -Type SamAccountName -MaxThreads 1

            $res[0].adObject.ObjectClass | Should-Be 'group'
            $res[0].adObject.Name | Should-Be 'Domain Users'
            @($res[0].adGroupMember).Count | Should-Be 1
            $res[0].adGroupMember[0].Name | Should-Be 'All users'
            $res[0].adGroupMember[0].SamAccountName | Should-Be 'All users'
        }
    }
}