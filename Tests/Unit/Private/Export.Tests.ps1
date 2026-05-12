#requires -Modules Pester

Describe 'Export.ps1 - Export Functions' {

    BeforeAll {
        $root = Split-Path -Parent $MyInvocation.MyCommand.Path
        $file = Join-Path $root '../Modules/Toolbox.PermissionMatrixHC/Private/Export.ps1'
        . $file
    }


    #
    # Build-ExportDataHC
    #
    Context 'Build-ExportDataHC' {

        It 'Aggregates permissions and formdata rows' {
            $imported = @(
                [pscustomobject]@{
                    File     = @{ Item = @{Name = 'A.xlsx' } }
                    Settings = @(
                        [pscustomobject]@{
                            Import = @{
                                ComputerName = 'PC1'
                                Path         = 'C:\'
                                Action       = 'Set'
                            }
                            Check  = @(
                                @{ Type = 'FatalError' }
                                @{ Type = 'Warning' }
                            )
                        }
                    )
                    FormData = @{
                        Import = @(
                            [pscustomobject]@{ Field = 'X'; Value = '1' }
                        )
                    }
                }
            )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            # permissions
            $res.Permissions.Count | Should -Be 1
            $res.Permissions[0].Errors | Should -Be 1
            $res.Permissions[0].Warnings | Should -Be 1

            # formdata
            $res.FormData.Count | Should -Be 1
        }
    }


    #
    # Export-PermissionsFileHC
    #
    Context 'Export-PermissionsFileHC' {

        It 'Calls Export-Excel with correct path' {

            Mock Export-Excel

            $rows = @([pscustomobject]@{A = 1 })
            $path = Join-Path $TestDrive 'perm.xlsx'

            Export-PermissionsFileHC -Rows $rows -Path $path

            Should -Invoke Export-Excel -Times 1 -ParameterFilter { $Path -eq $path }
        }
    }


    #
    # Export-ServiceNowFormDataHC
    #
    Context 'Export-ServiceNowFormDataHC' {

        It 'Calls Export-Excel for FormData' {
            Mock Export-Excel

            $rows = @([pscustomobject]@{F = 1 })
            $path = Join-Path $TestDrive 'form.xlsx'

            Export-ServiceNowFormDataHC -Rows $rows -Path $path

            Should -Invoke Export-Excel -Times 1 -ParameterFilter { $Path -eq $path }
        }
    }


    #
    # Export-OverviewHtmlHC
    #
    Context 'Export-OverviewHtmlHC' {

        It 'Writes HTML file' {
            $path = Join-Path $TestDrive 'overview.html'
            $html = '<html><body>test</body></html>'

            $result = Export-OverviewHtmlHC -Html $html -Path $path

            Test-Path $result | Should -BeTrue
            (Get-Content $path) | Should -Contain 'test'
        }
    }


    #
    # Export-FilesHC
    #
    Context 'Export-FilesHC' {

        It 'Runs all enabled exports' {

            Mock Export-PermissionsFileHC { return 'perm.xlsx' }
            Mock Export-ServiceNowFormDataHC { return 'form.xlsx' }
            Mock Export-OverviewHtmlHC { return 'overview.html' }
            Mock Build-ExportDataHC { return @{ Permissions = @(); FormData = @() } }

            $import = @()
            $settings = @{
                PermissionsExcelFile        = 'perm.xlsx'
                ServiceNowFormDataExcelFile = 'form.xlsx'
                OverviewHtmlFile            = 'overview.html'
            }

            $res = Export-FilesHC `
                -ImportedMatrix $import `
                -ExportSettings $settings `
                -HtmlOverview '<html></html>' `
                -Counters @{ }

            $res.Permissions | Should -Be 'perm.xlsx'
            $res.FormData | Should -Be 'form.xlsx'
            $res.OverviewHtml | Should -Be 'overview.html'

            Should -Invoke Build-ExportDataHC -Times 1
            Should -Invoke Export-PermissionsFileHC -Times 1
            Should -Invoke Export-ServiceNowFormDataHC -Times 1
            Should -Invoke Export-OverviewHtmlHC -Times 1
        }
    }
}