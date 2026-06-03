#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    . "$moduleRoot\Private\Utils.ps1"
    . "$moduleRoot\Private\Matrix.ps1"
}

Describe 'Format-FormDataStringsHC' {
    It 'trims all string properties' {
        $row = [pscustomobject]@{ Name = '  Bob  '; Note = ' hi ' }

        $res = Format-FormDataStringsHC -Row $row

        $res.Name | Should -Be 'Bob'
        $res.Note | Should -Be 'hi'
    }

    It 'does not change casing' {
        $res = Format-FormDataStringsHC -Row ([pscustomobject]@{ Name = ' MixedCase ' })
        $res.Name | Should -Be 'MixedCase'
    }

    It 'leaves non-string values untouched' {
        $row = [pscustomobject]@{ Count = 5; Flag = $true; Empty = $null }

        $res = Format-FormDataStringsHC -Row $row

        $res.Count | Should -Be 5
        $res.Flag | Should -BeTrue
        $res.Empty | Should -BeNullOrEmpty
    }

    It 'preserves column order' {
        $row = [pscustomobject]@{ Z = '1'; A = '2'; M = '3' }

        $res = Format-FormDataStringsHC -Row $row

        $res.PSObject.Properties.Name | Should -Be @('Z', 'A', 'M')
    }

    It 'accepts input from the pipeline' {
        $res = [pscustomobject]@{ Name = ' x ' } | Format-FormDataStringsHC
        $res.Name | Should -Be 'x'
    }

    It 'processes multiple rows from the pipeline' {
        $rows = @(
            [pscustomobject]@{ Name = ' a ' }
            [pscustomobject]@{ Name = ' b ' }
        )

        $res = $rows | Format-FormDataStringsHC

        $res.Count | Should -Be 2
        $res[0].Name | Should -Be 'a'
        $res[1].Name | Should -Be 'b'
    }
}

Describe 'Format-PermissionsStringsHC' {
    It 'trims and uppercases all string properties' {
        $row = [pscustomobject]@{ P1 = ' path '; P2 = ' r '; P3 = ' f ' }

        $res = Format-PermissionsStringsHC -Row $row

        $res.P1 | Should -Be 'PATH'
        $res.P2 | Should -Be 'R'
        $res.P3 | Should -Be 'F'
    }

    It 'leaves non-string values untouched' {
        $row = [pscustomobject]@{ P1 = 'x'; Count = 3 }

        $res = Format-PermissionsStringsHC -Row $row

        $res.Count | Should -Be 3
    }

    It 'preserves column order' {
        $row = [pscustomobject]@{ P3 = 'c'; P1 = 'a'; P2 = 'b' }

        $res = Format-PermissionsStringsHC -Row $row

        $res.PSObject.Properties.Name | Should -Be @('P3', 'P1', 'P2')
    }

    It 'accepts input from the pipeline' {
        $res = [pscustomobject]@{ P1 = ' abc ' } | Format-PermissionsStringsHC
        $res.P1 | Should -Be 'ABC'
    }
}

Describe 'Format-SettingStringsHC' {
    It 'trims all strings and title-cases the action' {
        $settings = [pscustomobject]@{
            Path   = ' C:\Test\ '
            Action = ' fix '
            Other  = '  value '
        }

        $res = Format-SettingStringsHC -Settings $settings

        $res.Path | Should -Be 'C:\Test'
        $res.Action | Should -Be 'Fix'
        $res.Other | Should -Be 'value'
    }

    It 'strips both forward and back trailing slashes from Path' {
        $res = Format-SettingStringsHC -Settings ([pscustomobject]@{ Path = 'C:\a\b//' })
        $res.Path | Should -Be 'C:\a\b'
    }

    It 'uppercases ComputerName' {
        $res = Format-SettingStringsHC -Settings ([pscustomobject]@{ ComputerName = ' srv01 ' })
        $res.ComputerName | Should -Be 'SRV01'
    }

    It 'title-cases an all-caps action' {
        $res = Format-SettingStringsHC -Settings ([pscustomobject]@{ Action = 'REPORT' })
        $res.Action | Should -Be 'Report'
    }

    It 'parses ApplyDefaultPermissions string into a boolean' {
        $res = Format-SettingStringsHC -Settings ([pscustomobject]@{ ApplyDefaultPermissions = 'true' })
        $res.ApplyDefaultPermissions | Should -BeOfType [bool]
        $res.ApplyDefaultPermissions | Should -BeTrue
    }

    It 'parses a false ApplyDefaultPermissions value' {
        $res = Format-SettingStringsHC -Settings ([pscustomobject]@{ ApplyDefaultPermissions = 'false' })
        $res.ApplyDefaultPermissions | Should -BeFalse
    }

    It 'leaves an unparseable ApplyDefaultPermissions as false' {
        $res = Format-SettingStringsHC -Settings ([pscustomobject]@{ ApplyDefaultPermissions = 'maybe' })
        $res.ApplyDefaultPermissions | Should -BeFalse
    }

    It 'does not mutate the original input object' {
        $settings = [pscustomobject]@{ Path = ' C:\x\ ' }

        $null = Format-SettingStringsHC -Settings $settings

        $settings.Path | Should -Be ' C:\x\ '
    }

    It 'does not error when optional properties are absent' {
        { Format-SettingStringsHC -Settings ([pscustomobject]@{ Other = 'x' }) } |
        Should -Not -Throw
    }

    It 'accepts input from the pipeline' {
        $res = [pscustomobject]@{ Action = 'fix' } | Format-SettingStringsHC
        $res.Action | Should -Be 'Fix'
    }
}

Describe 'ConvertTo-MatrixADNamesHC' {
    It 'adds Begin, Middle and header SamAccountNames, de-duplicated' {
        $headers = @(
            [pscustomobject]@{ P2 = 'Header1' }
            [pscustomobject]@{ P2 = 'Header2' }
            [pscustomobject]@{ P2 = 'Header1' }   # duplicate
        )

        $res = ConvertTo-MatrixADNamesHC -Begin 'GroupA' -Middle 'SiteB' -ColumnHeaders $headers

        $res | Should -Contain 'GroupA'
        $res | Should -Contain 'SiteB'
        $res | Should -Contain 'Header1'
        $res | Should -Contain 'Header2'
        $res.Count | Should -Be 4
    }

    It 'returns a unique sorted list' {
        $headers = @([pscustomobject]@{ P2 = 'Zeta' })

        $res = ConvertTo-MatrixADNamesHC -Begin 'Alpha' -Middle 'Mid' -ColumnHeaders $headers

        $res | Should -Be @('Alpha', 'Mid', 'Zeta')
    }

    It 'skips header rows without a P2 value' {
        $headers = @(
            [pscustomobject]@{ P2 = 'Keep' }
            [pscustomobject]@{ P2 = '' }
            [pscustomobject]@{ P2 = $null }
        )

        $res = ConvertTo-MatrixADNamesHC -Begin 'B' -Middle 'M' -ColumnHeaders $headers

        $res | Should -Contain 'Keep'
        $res.Count | Should -Be 3
    }
}

Describe 'Get-MatrixADObjectsMapHC' {
    It 'resolves GroupName and SiteCode placeholders against the setting row' {
        $sheet = @(
            [pscustomobject]@{ P1 = ''; P2 = 'GroupName'; P3 = 'SiteCode' }
            [pscustomobject]@{ P1 = ''; P2 = 'Admins'; P3 = 'Users' }
            [pscustomobject]@{ P1 = ''; P2 = ''; P3 = '' }
        )
        $setting = [pscustomobject]@{ GroupName = 'ACME'; SiteCode = 'BRU' }

        $map = Get-MatrixADObjectsMapHC -PermissionsSheet $sheet -SettingRow $setting

        # Header rows are walked bottom-to-top and joined with a space.
        $map['P2'] | Should -Be 'Admins ACME'
        $map['P3'] | Should -Be 'Users BRU'
    }

    It 'skips empty header cells without producing double spaces' {
        $sheet = @(
            [pscustomobject]@{ P1 = ''; P2 = '' }
            [pscustomobject]@{ P1 = ''; P2 = 'Solo' }
            [pscustomobject]@{ P1 = ''; P2 = '' }
        )
        $setting = [pscustomobject]@{ GroupName = 'G'; SiteCode = 'S' }

        $map = Get-MatrixADObjectsMapHC -PermissionsSheet $sheet -SettingRow $setting

        $map['P2'] | Should -Be 'Solo'
    }

    It 'omits a column whose resolved name is blank' {
        $sheet = @(
            [pscustomobject]@{ P1 = ''; P2 = '' }
            [pscustomobject]@{ P1 = ''; P2 = '' }
            [pscustomobject]@{ P1 = ''; P2 = '' }
        )
        $setting = [pscustomobject]@{ GroupName = 'G'; SiteCode = 'S' }

        $map = Get-MatrixADObjectsMapHC -PermissionsSheet $sheet -SettingRow $setting

        $map.Keys | Should -Not -Contain 'P2'
    }

    It 'stops at the first non-existent column' {
        $sheet = @(
            [pscustomobject]@{ P1 = ''; P2 = 'A'; P3 = 'B' }
            [pscustomobject]@{ P1 = ''; P2 = 'A'; P3 = 'B' }
            [pscustomobject]@{ P1 = ''; P2 = 'A'; P3 = 'B' }
        )
        $setting = [pscustomobject]@{ GroupName = 'G'; SiteCode = 'S' }

        $map = Get-MatrixADObjectsMapHC -PermissionsSheet $sheet -SettingRow $setting

        $map.Keys | Should -Be @('P2', 'P3')
    }

    It 'returns an ordered map keyed by column name' {
        $sheet = @(
            [pscustomobject]@{ P1 = ''; P2 = 'X'; P3 = 'Y' }
            [pscustomobject]@{ P1 = ''; P2 = 'X'; P3 = 'Y' }
            [pscustomobject]@{ P1 = ''; P2 = 'X'; P3 = 'Y' }
        )
        $setting = [pscustomobject]@{ GroupName = 'G'; SiteCode = 'S' }

        $map = Get-MatrixADObjectsMapHC -PermissionsSheet $sheet -SettingRow $setting

        $map.Keys | Should -Be @('P2', 'P3')
    }
}

Describe 'ConvertTo-MatrixAclHC' {
    BeforeAll {
        $script:adMap = [ordered]@{
            P2 = 'Obj1'
            P3 = 'Obj2'
        }
    }

    It 'builds ACL rules from data rows' {
        $rows = @(
            [pscustomobject]@{ P1 = 'Folder1'; P2 = 'R'; P3 = 'W' }
        )

        $res = ConvertTo-MatrixAclHC -DataRows $rows -AdObjectsMap $script:adMap

        $res.Count | Should -Be 1
        $res[0].Path | Should -Be 'Folder1'
        $res[0].ACL['Obj1'] | Should -Be 'R'
        $res[0].ACL['Obj2'] | Should -Be 'W'
    }

    It "ignores cells with permission 'I' (inherit)" {
        $rows = @(
            [pscustomobject]@{ P1 = 'Folder2'; P2 = 'I'; P3 = 'F' }
        )

        $res = ConvertTo-MatrixAclHC -DataRows $rows -AdObjectsMap $script:adMap

        $res[0].ACL.ContainsKey('Obj1') | Should -BeFalse
        $res[0].ACL['Obj2'] | Should -Be 'F'
    }

    It 'skips rows with no path (P1 empty)' {
        $rows = @(
            [pscustomobject]@{ P1 = ''; P2 = 'R'; P3 = 'W' }
            [pscustomobject]@{ P1 = 'Keep'; P2 = 'F'; P3 = '' }
        )

        $res = ConvertTo-MatrixAclHC -DataRows $rows -AdObjectsMap $script:adMap

        $res.Count | Should -Be 1
        $res[0].Path | Should -Be 'Keep'
    }

    It 'omits empty permission cells from the ACL' {
        $rows = @(
            [pscustomobject]@{ P1 = 'Folder'; P2 = 'R'; P3 = '' }
        )

        $res = ConvertTo-MatrixAclHC -DataRows $rows -AdObjectsMap $script:adMap

        $res[0].ACL.ContainsKey('Obj2') | Should -BeFalse
        $res[0].ACL.Count | Should -Be 1
    }

    It 'returns an empty result when all rows lack a path' {
        $rows = @(
            [pscustomobject]@{ P1 = ''; P2 = 'R' }
        )

        $res = ConvertTo-MatrixAclHC -DataRows $rows -AdObjectsMap $script:adMap

        @($res).Count | Should -Be 0
    }
}

Describe 'Get-DefaultAclHC' {
    BeforeEach {
        $script:errors = [System.Collections.Generic.List[PSObject]]::new()
    }

    It 'extracts valid default ACL entries' {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = 'Bob'; Permission = 'L' }
            [pscustomobject]@{ ADObjectName = 'Mike'; Permission = 'R' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl['Bob'] | Should -Be 'L'
        $acl['Mike'] | Should -Be 'R'
        $acl.Keys.Count | Should -Be 2
        $script:errors.Count | Should -Be 0
    }

    It 'trims and uppercases the permission character' {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = ' Bob '; Permission = ' w ' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl['Bob'] | Should -Be 'W'
    }

    It 'silently skips rows where both name and permission are empty' {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = ''; Permission = '' }
            [pscustomobject]@{ ADObjectName = 'Bob'; Permission = 'R' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl.Keys.Count | Should -Be 1
        $script:errors.Count | Should -Be 0
    }

    It 'flags a row with Permission but no ADObjectName' {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = ''; Permission = 'F' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl.Keys.Count | Should -Be 0
        $script:errors.Count | Should -Be 1
        $script:errors[0].Type | Should -Be 'FatalError'
        $script:errors[0].Name | Should -Be 'Incomplete default ACL entry'
    }

    It 'flags a row with ADObjectName but no Permission' {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = 'Bob'; Permission = '' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl.Keys.Count | Should -Be 0
        $script:errors[0].Name | Should -Be 'Incomplete default ACL entry'
    }

    It 'flags an invalid permission character' {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = 'Bob'; Permission = 'X' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl.Keys.Count | Should -Be 0
        $script:errors[0].Name | Should -Be 'Invalid default ACL permission'
    }

    It "rejects the 'I' (inherit) permission in defaults" {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = 'Bob'; Permission = 'I' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl.Keys.Count | Should -Be 0
        $script:errors[0].Name | Should -Be 'Invalid default ACL permission'
    }

    It 'flags a duplicate ADObjectName (case insensitive)' {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = 'BOB'; Permission = 'R' }
            [pscustomobject]@{ ADObjectName = 'Bob'; Permission = 'W' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl.Keys.Count | Should -Be 1
        $acl['Bob'] | Should -Be 'R'   # first one wins
        $script:errors[0].Name | Should -Be 'Duplicate default ACL entry'
    }

    It 'returns an empty hashtable when every row is a skip-row' {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = ''; Permission = '' }
            [pscustomobject]@{ ADObjectName = $null; Permission = $null }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl | Should -BeOfType [hashtable]
        $acl.Keys.Count | Should -Be 0
        $script:errors.Count | Should -Be 0
    }

    It 'returns an empty hashtable for an empty sheet' {
        $acl = Get-DefaultAclHC -Sheet @() -SystemErrors ([ref]$script:errors)

        $acl | Should -BeOfType [hashtable]
        $acl.Keys.Count | Should -Be 0
        $script:errors.Count | Should -Be 0
    }

    It 'silently skips a MailTo-only row (name and permission empty)' {
        # The Defaults sheet carries a MailTo column; rows that only populate
        # MailTo have empty ADObjectName/Permission and must be ignored.
        $sheet = @(
            [pscustomobject]@{ ADObjectName = $null; Permission = $null; MailTo = 'mike@contoso.com' }
            [pscustomobject]@{ ADObjectName = 'Bob'; Permission = 'R' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $acl.Keys.Count | Should -Be 1
        $acl['Bob'] | Should -Be 'R'
        $script:errors.Count | Should -Be 0
    }

    It 'continues processing valid rows after a bad row' {
        $sheet = @(
            [pscustomobject]@{ ADObjectName = 'BadOne'; Permission = 'X' }
            [pscustomobject]@{ ADObjectName = 'GoodOne'; Permission = 'L' }
        )

        $acl = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$script:errors)

        $script:errors.Count | Should -Be 1
        $acl.Keys.Count | Should -Be 1
        $acl['GoodOne'] | Should -Be 'L'
    }
}

Describe 'Merge-DefaultPermissionsHC' {
    It 'returns a clone of the matrix ACL when ApplyDefaultPermissions is false' {
        $defaults = @{ Bob = 'F' }
        $matrix = @{ Alice = 'R' }

        $res = Merge-DefaultPermissionsHC -Defaults $defaults -MatrixAcl $matrix -ApplyDefaultPermissions $false

        $res.ContainsKey('Bob') | Should -BeFalse
        $res['Alice'] | Should -Be 'R'
    }

    It 'merges defaults into the matrix ACL when enabled' {
        $defaults = @{ Bob = 'F' }
        $matrix = @{ Alice = 'R' }

        $res = Merge-DefaultPermissionsHC -Defaults $defaults -MatrixAcl $matrix -ApplyDefaultPermissions $true

        $res['Alice'] | Should -Be 'R'
        $res['Bob'] | Should -Be 'F'
        $res.Keys.Count | Should -Be 2
    }

    It 'throws on a key present in both matrix and defaults' {
        $defaults = @{ Shared = 'F' }
        $matrix = @{ Shared = 'R' }

        { Merge-DefaultPermissionsHC -Defaults $defaults -MatrixAcl $matrix -ApplyDefaultPermissions $true } |
        Should -Throw '*conflict*'
    }

    It 'names the conflicting AD object in the error' {
        $defaults = @{ Conflicto = 'F' }
        $matrix = @{ Conflicto = 'R' }

        { Merge-DefaultPermissionsHC -Defaults $defaults -MatrixAcl $matrix -ApplyDefaultPermissions $true } |
        Should -Throw '*Conflicto*'
    }

    It 'does not mutate the original matrix ACL' {
        $defaults = @{ Bob = 'F' }
        $matrix = @{ Alice = 'R' }

        $null = Merge-DefaultPermissionsHC -Defaults $defaults -MatrixAcl $matrix -ApplyDefaultPermissions $true

        $matrix.ContainsKey('Bob') | Should -BeFalse
        $matrix.Keys.Count | Should -Be 1
    }

    It 'handles empty defaults' {
        $matrix = @{ Alice = 'R' }

        $res = Merge-DefaultPermissionsHC -Defaults @{} -MatrixAcl $matrix -ApplyDefaultPermissions $true

        $res['Alice'] | Should -Be 'R'
        $res.Keys.Count | Should -Be 1
    }
}