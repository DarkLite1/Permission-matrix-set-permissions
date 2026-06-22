#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Unit test for Resolve-ResponsibleEmailHC.
#
# The e-mail pass-through case makes no AD call and always runs. The user/group
# cases need the ActiveDirectory cmdlets to exist so Pester can mock them; where
# they are absent (e.g. plain CI) those cases are skipped.

Describe 'Resolve-ResponsibleEmailHC' {
    BeforeDiscovery {
        $script:AdAvailable = [bool](Get-Command Get-ADObject -ErrorAction SilentlyContinue)
    }

    BeforeAll {
        $dir = $PSScriptRoot
        while ($dir -and -not (Test-Path (Join-Path $dir 'Modules\PermissionMatrix\PermissionMatrix.psm1'))) {
            $dir = Split-Path $dir -Parent
        }
        if (-not $dir) {
            throw "Could not find Modules\PermissionMatrix\PermissionMatrix.psm1 above '$PSScriptRoot'."
        }
        Import-Module (Join-Path $dir 'Modules\PermissionMatrix\PermissionMatrix.psm1') -Force
    }

    It 'passes e-mail addresses through unchanged and de-duplicated' {
        $r = InModuleScope PermissionMatrix {
            Resolve-ResponsibleEmailHC -Responsible 'a@x.com, b@y.com, a@x.com'
        }
        $r.Emails     | Should -Be @('a@x.com', 'b@y.com')
        $r.Unresolved | Should -BeNullOrEmpty
    }

    It 'resolves a user to its mail attribute' -Skip:(-not $AdAvailable) {
        Mock Get-ADObject -ModuleName PermissionMatrix {
            [pscustomobject]@{ objectClass = 'user'; mail = 'user@x.com'; DistinguishedName = 'CN=U' }
        }
        $r = InModuleScope PermissionMatrix { Resolve-ResponsibleEmailHC -Responsible 'jdoe' }
        $r.Emails | Should -Be @('user@x.com')
    }

    It 'resolves a group to member e-mail, recursing nested groups' -Skip:(-not $AdAvailable) {
        Mock Get-ADObject -ModuleName PermissionMatrix {
            [pscustomobject]@{ objectClass = 'group'; mail = $null; DistinguishedName = 'CN=G' }
        }
        Mock Get-ADGroupMember -ModuleName PermissionMatrix {
            @(
                [pscustomobject]@{ objectClass = 'user'; name = 'Alice'; distinguishedName = 'CN=Alice' }
                [pscustomobject]@{ objectClass = 'user'; name = 'Bob'; distinguishedName = 'CN=Bob' }
            )
        }
        Mock Get-ADUser -ModuleName PermissionMatrix {
            if ("$Identity" -like '*Alice*') { [pscustomobject]@{ EmailAddress = 'alice@x.com' } }
            else { [pscustomobject]@{ EmailAddress = 'bob@x.com' } }
        }
        $r = InModuleScope PermissionMatrix { Resolve-ResponsibleEmailHC -Responsible 'Some Group' }
        $r.Emails     | Should -Be @('alice@x.com', 'bob@x.com')
        $r.Unresolved | Should -BeNullOrEmpty
    }

    It 'reports group members without an e-mail address' -Skip:(-not $AdAvailable) {
        Mock Get-ADObject -ModuleName PermissionMatrix {
            [pscustomobject]@{ objectClass = 'group'; mail = $null; DistinguishedName = 'CN=G' }
        }
        Mock Get-ADGroupMember -ModuleName PermissionMatrix {
            @(
                [pscustomobject]@{ objectClass = 'user'; name = 'Alice'; distinguishedName = 'CN=Alice' }
                [pscustomobject]@{ objectClass = 'user'; name = 'NoMail'; distinguishedName = 'CN=NoMail' }
            )
        }
        Mock Get-ADUser -ModuleName PermissionMatrix {
            if ("$Identity" -like '*Alice*') { [pscustomobject]@{ EmailAddress = 'alice@x.com' } }
            else { [pscustomobject]@{ EmailAddress = $null } }
        }
        $r = InModuleScope PermissionMatrix { Resolve-ResponsibleEmailHC -Responsible 'Some Group' }
        $r.Emails                    | Should -Be @('alice@x.com')
        (@($r.Unresolved) -join ';') | Should -Match 'NoMail'
    }

    It 'reports a token that cannot be found in AD' -Skip:(-not $AdAvailable) {
        Mock Get-ADObject -ModuleName PermissionMatrix { }
        $r = InModuleScope PermissionMatrix { Resolve-ResponsibleEmailHC -Responsible 'ghost' }
        $r.Emails                    | Should -BeNullOrEmpty
        (@($r.Unresolved) -join ';') | Should -Match 'ghost'
    }

    It 'excludes a placeholder listed directly as the responsible' {
        # No AD call: the placeholder is dropped before any lookup.
        $r = InModuleScope PermissionMatrix {
            Resolve-ResponsibleEmailHC -Responsible 'cnorris' -ExcludeSamAccountName @('cnorris')
        }
        $r.Emails     | Should -BeNullOrEmpty
        $r.Unresolved | Should -BeNullOrEmpty
    }

    It 'excludes placeholder group members (Matrix.AdGroupPlaceHolders)' -Skip:(-not $AdAvailable) {
        Mock Get-ADObject -ModuleName PermissionMatrix {
            [pscustomobject]@{ objectClass = 'group'; mail = $null; DistinguishedName = 'CN=G' }
        }
        Mock Get-ADGroupMember -ModuleName PermissionMatrix {
            @(
                [pscustomobject]@{ objectClass = 'user'; name = 'Alice'; SamAccountName = 'alice'; distinguishedName = 'CN=Alice' }
                [pscustomobject]@{ objectClass = 'user'; name = 'cnorris'; SamAccountName = 'cnorris'; distinguishedName = 'CN=cnorris' }
            )
        }
        Mock Get-ADUser -ModuleName PermissionMatrix {
            if ("$Identity" -like '*Alice*') { [pscustomobject]@{ EmailAddress = 'alice@x.com' } }
            else { [pscustomobject]@{ EmailAddress = 'cnorris@x.com' } }
        }
        $r = InModuleScope PermissionMatrix {
            Resolve-ResponsibleEmailHC -Responsible 'Some Group' -ExcludeSamAccountName @('cnorris')
        }
        $r.Emails     | Should -Be @('alice@x.com')
        $r.Unresolved | Should -BeNullOrEmpty
    }
}