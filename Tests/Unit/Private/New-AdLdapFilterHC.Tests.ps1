#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

<#
    Tests for New-AdLdapFilterHC in
    Modules\PermissionMatrix\Private\ActiveDirectory.ps1

    Unlike the rest of ActiveDirectory.ps1, this function is pure string
    building with no directory access, so it is a unit test and needs no
    domain controller. It was extracted out of the ForEach-Object -Parallel
    block in Get-ADObjectDetailHC precisely so it could be covered here.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    . "$root\Modules\PermissionMatrix\Private\ActiveDirectory.ps1"
}

Describe 'New-AdLdapFilterHC' {
    Context 'Input shapes' {
        It 'builds a samAccountName filter for a bare name' {
            New-AdLdapFilterHC -Name 'jdoe' -Type 'SamAccountName' |
            Should-Be '(samAccountName=jdoe)'
        }

        It 'strips the NetBIOS prefix from DOMAIN\name' {
            New-AdLdapFilterHC -Name 'DOMAIN\jdoe' -Type 'SamAccountName' |
            Should-Be '(samAccountName=jdoe)'
        }

        It 'searches both attributes for a UPN' {
            New-AdLdapFilterHC -Name 'jdoe@contoso.com' -Type 'SamAccountName' |
            Should-Be '(|(samAccountName=jdoe@contoso.com)(userPrincipalName=jdoe@contoso.com))'
        }

        It 'builds a distinguishedName filter' {
            New-AdLdapFilterHC -Name 'CN=jdoe,OU=Users,DC=contoso,DC=com' -Type 'DistinguishedName' |
            Should-Be '(distinguishedName=CN=jdoe,OU=Users,DC=contoso,DC=com)'
        }
    }

    Context 'RFC 4515 escaping' {
        It 'escapes parentheses so the filter stays well formed' {
            # Unescaped, this built a malformed filter: the search threw and
            # the group was wrongly reported as missing from AD.
            New-AdLdapFilterHC -Name 'Finance (EU)' -Type 'SamAccountName' |
            Should-Be '(samAccountName=Finance \28EU\29)'
        }

        It 'escapes an asterisk so it cannot act as a wildcard' {
            <# The important one. Unescaped, 'HR_*' matched any object whose
            name started with HR_, so permissions could be granted to a group
            the matrix never named. #>
            New-AdLdapFilterHC -Name 'HR_*' -Type 'SamAccountName' |
            Should-Be '(samAccountName=HR_\2a)'
        }

        It 'escapes a NUL character' {
            New-AdLdapFilterHC -Name "jdoe`0" -Type 'SamAccountName' |
            Should-Be '(samAccountName=jdoe\00)'
        }

        It 'escapes the backslash in a distinguished name' {
            # AD returns DNs with commas escaped as '\,', which is not a valid
            # filter escape sequence, so managers with a comma never resolved.
            New-AdLdapFilterHC -Name 'CN=Doe\, John,OU=Users,DC=contoso,DC=com' -Type 'DistinguishedName' |
            Should-Be '(distinguishedName=CN=Doe\5c, John,OU=Users,DC=contoso,DC=com)'
        }

        It 'escapes the backslash before the characters it introduces' {
            <# Ordering guard: escaping '(' first would turn the backslash of
            the resulting '\28' into '\5c28'. #>
            New-AdLdapFilterHC -Name 'OU=A\(B' -Type 'DistinguishedName' |
            Should-Be '(distinguishedName=OU=A\5c\28B)'
        }

        It 'escapes both halves of a UPN filter identically' {
            New-AdLdapFilterHC -Name 'a*b@contoso.com' -Type 'SamAccountName' |
            Should-Be '(|(samAccountName=a\2ab@contoso.com)(userPrincipalName=a\2ab@contoso.com))'
        }

        It 'strips the NetBIOS prefix before escaping' {
            # The prefix separator must be matched on the raw value, otherwise
            # it is already '\5c' by the time the strip runs.
            New-AdLdapFilterHC -Name 'DOMAIN\Finance (EU)' -Type 'SamAccountName' |
            Should-Be '(samAccountName=Finance \28EU\29)'
        }
    }
}