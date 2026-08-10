#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

<#
    Tests for ConvertTo-EscapedLdapValueHC and New-AdLdapFilterHC in
    Modules\PermissionMatrix\Private\ActiveDirectory.ps1

    Unlike the rest of ActiveDirectory.ps1, both functions are pure string
    building with no directory access, so these are unit tests and need no
    domain controller. They were extracted out of the ForEach-Object -Parallel
    block in Get-ADObjectDetailHC precisely so they could be covered here.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    . "$root\Modules\PermissionMatrix\Private\ActiveDirectory.ps1"
}

Describe 'ConvertTo-EscapedLdapValueHC' {
    It 'leaves a plain value untouched' {
        ConvertTo-EscapedLdapValueHC -Value 'jdoe' | Should-Be 'jdoe'
    }

    It 'escapes parentheses' {
        ConvertTo-EscapedLdapValueHC -Value 'Finance (EU)' |
        Should-Be 'Finance \28EU\29'
    }

    It 'escapes an asterisk so it cannot act as a wildcard' {
        ConvertTo-EscapedLdapValueHC -Value 'HR_*' | Should-Be 'HR_\2a'
    }

    It 'escapes a backslash' {
        ConvertTo-EscapedLdapValueHC -Value 'Doe\, John' | Should-Be 'Doe\5c, John'
    }

    It 'escapes a NUL character' {
        ConvertTo-EscapedLdapValueHC -Value "jdoe`0" | Should-Be 'jdoe\00'
    }

    It 'escapes the backslash before the characters it introduces' {
        # Ordering guard: escaping '(' first would turn the backslash of the
        # resulting '\28' into '\5c28'.
        ConvertTo-EscapedLdapValueHC -Value 'A\(B' | Should-Be 'A\5c\28B'
    }

    It 'accepts an empty string' {
        ConvertTo-EscapedLdapValueHC -Value '' | Should-Be ''
    }
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

Describe 'Use inside ForEach-Object -Parallel runspaces' {
    <#
        Get-ADObjectDetailHC builds its filters inside ForEach-Object -Parallel.
        Those runspaces do not inherit the caller's session state, so the two
        functions are passed in as source and redefined per runspace with
        ${function:Name} = $using:definition.

        That idiom is the fragile part of the arrangement: it breaks silently
        if a definition stops round-tripping through .ToString(), or if a
        dependency is forgotten. New-AdLdapFilterHC calling
        ConvertTo-EscapedLdapValueHC is exactly such a dependency, and leaving
        it out fails only in the parallel path, never in the tests above.

        The definitions are captured inside each It rather than in a BeforeAll,
        so $using: reads them from the immediate local scope, exactly as
        Get-ADObjectDetailHC does.
    #>

    It 'builds the same filters in a runspace as it does in process' {
        $filterFunction = ${function:New-AdLdapFilterHC}.ToString()
        $escapeFunction = ${function:ConvertTo-EscapedLdapValueHC}.ToString()

        $names = @(
            'jdoe'
            'Finance (EU)'
            'DOMAIN\jdoe'
            'HR_*'
            'a*b@contoso.com'
        )

        $expected = $names | ForEach-Object {
            New-AdLdapFilterHC -Name $_ -Type 'SamAccountName'
        } | Sort-Object

        # Parallel output order is not guaranteed, so both sides are sorted.
        $actual = $names | ForEach-Object -ThrottleLimit 3 -Parallel {
            ${function:ConvertTo-EscapedLdapValueHC} = $using:escapeFunction
            ${function:New-AdLdapFilterHC} = $using:filterFunction

            New-AdLdapFilterHC -Name $_ -Type 'SamAccountName'
        } | Sort-Object

        $actual | Should-BeCollection $expected
    }

    It 'reaches the injected escape helper from the injected filter builder' {
        # Without $escapeFunction this throws CommandNotFoundException inside
        # the runspace, which is the failure the parallel path would hide.
        $filterFunction = ${function:New-AdLdapFilterHC}.ToString()
        $escapeFunction = ${function:ConvertTo-EscapedLdapValueHC}.ToString()

        $result = @('Finance (EU)') | ForEach-Object -Parallel {
            ${function:ConvertTo-EscapedLdapValueHC} = $using:escapeFunction
            ${function:New-AdLdapFilterHC} = $using:filterFunction

            New-AdLdapFilterHC -Name $_ -Type 'SamAccountName'
        }

        $result | Should-Be '(samAccountName=Finance \28EU\29)'
    }

    It 'builds distinguishedName filters in a runspace' {
        $filterFunction = ${function:New-AdLdapFilterHC}.ToString()
        $escapeFunction = ${function:ConvertTo-EscapedLdapValueHC}.ToString()

        $result = @('OU=A\(B') | ForEach-Object -Parallel {
            ${function:ConvertTo-EscapedLdapValueHC} = $using:escapeFunction
            ${function:New-AdLdapFilterHC} = $using:filterFunction

            New-AdLdapFilterHC -Name $_ -Type 'DistinguishedName'
        }

        $result | Should-Be '(distinguishedName=OU=A\5c\28B)'
    }
}