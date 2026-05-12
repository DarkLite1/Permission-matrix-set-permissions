#requires -Modules Pester

Describe 'Matrix.ps1 - Matrix Domain Functions' {

    BeforeAll {
        $root = Split-Path -Parent $MyInvocation.MyCommand.Path
        $file = Join-Path $root '../Modules/Toolbox.PermissionMatrixHC/Private/Matrix.ps1'
        . $file
    }

    Context 'Format-PermissionsStringsHC' {

        It 'Uppercases and trims all string properties' {
            $row = [pscustomobject]@{
                P1 = ' path '
                P2 = ' r '
                P3 = ' f '
            }

            $res = Format-PermissionsStringsHC -Row $row
            $res.P1 | Should -Be 'PATH'
            $res.P2 | Should -Be 'R'
            $res.P3 | Should -Be 'F'
        }
    }


    Context 'Format-SettingStringsHC' {

        It 'Trims all strings and normalizes action' {
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
    }


    Context 'ConvertTo-MatrixADNamesHC' {

        It 'Adds Begin, Middle, and header SamAccountNames' {
            $headers = @(
                [pscustomobject]@{ P2 = 'Header1' },
                [pscustomobject]@{ P2 = 'Header2' },
                [pscustomobject]@{ P2 = 'Header1' }  # Duplicate
            )

            $res = ConvertTo-MatrixADNamesHC `
                -Begin 'GroupA' `
                -Middle 'SiteB' `
                -ColumnHeaders $headers

            $res | Should -Contain 'GroupA'
            $res | Should -Contain 'SiteB'
            $res | Should -Contain 'Header1'
            $res | Should -Contain 'Header2'
            $res.Count | Should -Be 4
        }
    }


    Context 'ConvertTo-MatrixAclHC' {

        It 'Builds ACL rules from non-header rows and AD objects' {

            $rows = @(
                [pscustomobject]@{ P1 = 'Folder1'; P2 = 'R'; P3 = 'W' },
                [pscustomobject]@{ P1 = 'Folder2'; P2 = 'I'; P3 = 'F' }  # I = ignore
            )

            $ad = @('Obj1', 'Obj2')

            $res = ConvertTo-MatrixAclHC `
                -NonHeaderRows $rows `
                -ADObjects $ad

            $res.Count | Should -Be 2

            $res[0].ACL['Obj1'] | Should -Be 'R'
            $res[0].ACL['Obj2'] | Should -Be 'W'

            # Folder2: Obj1 ignored, Obj2 = F
            $res[1].ACL.ContainsKey('Obj1') | Should -BeFalse
            $res[1].ACL['Obj2'] | Should -Be 'F'
        }
    }


    Context 'Get-DefaultAclHC' {

        It 'Extracts default ACL entries' {
            $sheet = @(
                [pscustomobject]@{ ADObjectName = 'Bob'; Permission = 'L' },
                [pscustomobject]@{ ADObjectName = 'Mike'; Permission = 'R' },
                [pscustomobject]@{ ADObjectName = ''; Permission = 'F' }  # ignored
            )

            $acl = Get-DefaultAclHC -Sheet $sheet

            $acl['Bob'] | Should -Be 'L'
            $acl['Mike'] | Should -Be 'R'
            $acl.Keys.Count | Should -Be 2
        }
    }
}