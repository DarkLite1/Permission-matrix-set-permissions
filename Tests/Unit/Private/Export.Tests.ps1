#Requires -Version 7
#requires -Modules Pester

Describe 'Export-FilesHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Export.ps1"
        . "$moduleRoot\Private\Html.ps1"

        $script:FakeMatrices = @(
            [pscustomobject]@{ Id = 1 }
            [pscustomobject]@{ Id = 2 }
        )

        $script:FakeExportData = @{
            Permissions = @(
                [pscustomobject]@{ Path = 'C:\A'; Permission = 'R' }
            )
            FormData    = @(
                [pscustomobject]@{
                    MatrixCategoryName      = 'Cat'
                    MatrixSubCategoryName   = 'Sub'
                    MatrixFolderDisplayName = 'Folder'
                    MatrixFilePath          = 'C:\X.xlsx'
                    MatrixFileName          = 'X.xlsx'
                    MatrixResponsible       = 'a@b.com'
                }
            )
        }
    }

    BeforeEach {
        Mock Build-ExportDataHC { return $FakeExportData }
        Mock Export-PermissionsFileHC { return 'TestDrive:\Permissions.xlsx' }
        Mock Export-ServiceNowFormDataHC { return 'TestDrive:\ServiceNow.xlsx' }
        Mock New-OverviewHtmlHC { return '<html>overview</html>' }
        Mock Export-OverviewHtmlHC { return 'TestDrive:\Overview.html' }
    }

    Context 'Permissions export' {
        It 'writes the permissions file when PermissionsExcelFile is set' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = 'TestDrive:\Permissions.xlsx'
                ServiceNowFormDataExcelFile = $null
                OverviewHtmlFile            = $null
            }

            $result = Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings

            Should -Invoke Export-PermissionsFileHC -Times 1 -ParameterFilter {
                $Path -eq 'TestDrive:\Permissions.xlsx'
            }
            $result.Permissions | Should -Be 'TestDrive:\Permissions.xlsx'
        }

        It 'skips the permissions file when PermissionsExcelFile is null' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = $null
                ServiceNowFormDataExcelFile = $null
                OverviewHtmlFile            = $null
            }

            $result = Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings

            Should -Invoke Export-PermissionsFileHC -Times 0
            $result.Permissions | Should -BeNullOrEmpty
        }

        It 'passes Permissions rows from Build-ExportDataHC to the writer' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = 'TestDrive:\Permissions.xlsx'
                ServiceNowFormDataExcelFile = $null
                OverviewHtmlFile            = $null
            }
 
            Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings | Out-Null
 
            Should -Invoke Export-PermissionsFileHC -Times 1 -ParameterFilter {
                # Array -eq is element-wise filter in PowerShell, not equality.
                # Assert on a property of the rows we expect to have been passed.
                $Rows.Count -eq 1 -and $Rows[0].Path -eq 'C:\A'
            }
        }
    }

    Context 'ServiceNow FormData export' {
        It 'writes the FormData file when ServiceNowFormDataExcelFile is set' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = $null
                ServiceNowFormDataExcelFile = 'TestDrive:\ServiceNow.xlsx'
                OverviewHtmlFile            = $null
            }

            $result = Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings

            Should -Invoke Export-ServiceNowFormDataHC -Times 1 -ParameterFilter {
                $Path -eq 'TestDrive:\ServiceNow.xlsx'
            }
            $result.FormData | Should -Be 'TestDrive:\ServiceNow.xlsx'
        }

        It 'skips the FormData file when ServiceNowFormDataExcelFile is null' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = $null
                ServiceNowFormDataExcelFile = $null
                OverviewHtmlFile            = $null
            }

            Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings | Out-Null

            Should -Invoke Export-ServiceNowFormDataHC -Times 0
        }
    }

    Context 'Overview HTML export' {
        It 'builds HTML internally and writes when OverviewHtmlFile is set' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = $null
                ServiceNowFormDataExcelFile = $null
                OverviewHtmlFile            = 'TestDrive:\Overview.html'
            }

            $result = Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings

            Should -Invoke New-OverviewHtmlHC -Times 1
            Should -Invoke Export-OverviewHtmlHC -Times 1 -ParameterFilter {
                $Path -eq 'TestDrive:\Overview.html'
            }
            $result.OverviewHtml | Should -Be 'TestDrive:\Overview.html'
        }

        It 'passes FormData rows (not ImportedMatrix) to New-OverviewHtmlHC' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = $null
                ServiceNowFormDataExcelFile = $null
                OverviewHtmlFile            = 'TestDrive:\Overview.html'
            }
 
            Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings | Out-Null
 
            Should -Invoke New-OverviewHtmlHC -Times 1 -ParameterFilter {
                # Verify the rows came from FakeExportData.FormData (which has
                # MatrixFileName='X.xlsx'), not from $FakeMatrices (which has
                # only an Id property).
                $FormData.Count -eq 1 -and $FormData[0].MatrixFileName -eq 'X.xlsx'
            }
        }

        It 'passes the generated HTML to the writer' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = $null
                ServiceNowFormDataExcelFile = $null
                OverviewHtmlFile            = 'TestDrive:\Overview.html'
            }

            Mock New-OverviewHtmlHC { return '<html>custom-overview-content</html>' }

            Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings | Out-Null

            Should -Invoke Export-OverviewHtmlHC -Times 1 -ParameterFilter {
                $Html -eq '<html>custom-overview-content</html>'
            }
        }

        It 'skips the HTML overview when OverviewHtmlFile is null' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = $null
                ServiceNowFormDataExcelFile = $null
                OverviewHtmlFile            = $null
            }

            Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings | Out-Null

            Should -Invoke New-OverviewHtmlHC -Times 0
            Should -Invoke Export-OverviewHtmlHC -Times 0
        }
    }

    Context 'overall behavior' {
        It 'returns a hashtable with Permissions, FormData, and OverviewHtml keys' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = $null
                ServiceNowFormDataExcelFile = $null
                OverviewHtmlFile            = $null
            }

            $result = Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings

            $result.Keys | Should -Contain 'Permissions'
            $result.Keys | Should -Contain 'FormData'
            $result.Keys | Should -Contain 'OverviewHtml'
        }

        It 'calls Build-ExportDataHC exactly once regardless of which exports are configured' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = 'TestDrive:\Permissions.xlsx'
                ServiceNowFormDataExcelFile = 'TestDrive:\ServiceNow.xlsx'
                OverviewHtmlFile            = 'TestDrive:\Overview.html'
            }

            Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings | Out-Null

            Should -Invoke Build-ExportDataHC -Times 1
        }

        It 'produces all three artifacts when all settings are configured' {
            $settings = [pscustomobject]@{
                PermissionsExcelFile        = 'TestDrive:\Permissions.xlsx'
                ServiceNowFormDataExcelFile = 'TestDrive:\ServiceNow.xlsx'
                OverviewHtmlFile            = 'TestDrive:\Overview.html'
            }

            $result = Export-FilesHC -ImportedMatrix $FakeMatrices -ExportSettings $settings

            $result.Permissions | Should -Not -BeNullOrEmpty
            $result.FormData | Should -Not -BeNullOrEmpty
            $result.OverviewHtml | Should -Not -BeNullOrEmpty
        }
    }
}

Describe 'Export.ps1 - Export Functions' {

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Export.ps1"
    }

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

    Context 'Export-PermissionsFileHC' {

        It 'Calls Export-Excel with correct path' {

            Mock Export-Excel

            $rows = @([pscustomobject]@{A = 1 })
            $path = Join-Path $TestDrive 'perm.xlsx'

            Export-PermissionsFileHC -Rows $rows -Path $path

            Should -Invoke Export-Excel -Times 1 -ParameterFilter { $Path -eq $path }
        }
    }

    Context 'Export-ServiceNowFormDataHC' {

        It 'Calls Export-Excel for FormData' {
            Mock Export-Excel

            $rows = @([pscustomobject]@{F = 1 })
            $path = Join-Path $TestDrive 'form.xlsx'

            Export-ServiceNowFormDataHC -Rows $rows -Path $path

            Should -Invoke Export-Excel -Times 1 -ParameterFilter { $Path -eq $path }
        }
    }

    Context 'Export-OverviewHtmlHC' {

        It 'Writes HTML file' {
            $path = Join-Path $TestDrive 'overview.html'
            $html = '<html><body>test</body></html>'

            $result = Export-OverviewHtmlHC -Html $html -Path $path

            Test-Path $result | Should -BeTrue
            (Get-Content $path) | Should -Contain 'test'
        }
    }

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
} # -Skip