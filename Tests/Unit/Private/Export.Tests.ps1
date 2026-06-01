#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    . "$moduleRoot\Private\Export.ps1"
    . "$moduleRoot\Private\Html.ps1"
}

Describe 'Export-FilesHC' {
    BeforeAll {
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

Describe 'Build-ExportDataHC' {
    BeforeAll {
        # Builder mirroring the imported-matrix shape Build-ExportDataHC consumes:
        #   $I.File.Item.Name
        #   $I.Settings[].Import.{ComputerName,Path,Action}
        #   $I.Settings[].Check[].Type
        #   $I.FormData.Import[]
        function New-ImportedMatrix {
            param(
                [string]$FileName = 'A.xlsx',
                [object[]]$Settings = @(),
                [object[]]$FormDataImport = $null,
                [switch]$NoFormDataProperty
            )

            $obj = [pscustomobject]@{
                File     = [pscustomobject]@{ Item = [pscustomobject]@{ Name = $FileName } }
                Settings = $Settings
            }

            if (-not $NoFormDataProperty) {
                $obj | Add-Member -NotePropertyName FormData -NotePropertyValue (
                    [pscustomobject]@{ Import = $FormDataImport }
                )
            }

            return $obj
        }

        function New-Setting {
            param(
                [string]$ComputerName = 'PC1',
                [string]$Path = 'C:\',
                [string]$Action = 'Set',
                [object[]]$Check = @()
            )
            return [pscustomobject]@{
                Import = [pscustomobject]@{
                    ComputerName = $ComputerName
                    Path         = $Path
                    Action       = $Action
                }
                Check  = $Check
            }
        }
    }

    Context 'permissions row mapping' {
        It 'maps the matrix file name and import fields onto each permissions row' {
            $imported = @(
                New-ImportedMatrix -FileName 'Team.xlsx' -Settings @(
                    New-Setting -ComputerName 'SRV9' -Path 'D:\data' -Action 'Fix'
                )
            )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.Permissions.Count | Should -Be 1
            $res.Permissions[0].MatrixFile | Should -Be 'Team.xlsx'
            $res.Permissions[0].Computer | Should -Be 'SRV9'
            $res.Permissions[0].Path | Should -Be 'D:\data'
            $res.Permissions[0].Action | Should -Be 'Fix'
        }

        It 'counts FatalError checks into Errors and Warning checks into Warnings' {
            $imported = @(
                New-ImportedMatrix -Settings @(
                    New-Setting -Check @(
                        [pscustomobject]@{ Type = 'FatalError' }
                        [pscustomobject]@{ Type = 'FatalError' }
                        [pscustomobject]@{ Type = 'Warning' }
                        [pscustomobject]@{ Type = 'Information' }
                    )
                )
            )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.Permissions[0].Errors | Should -Be 2
            $res.Permissions[0].Warnings | Should -Be 1
        }

        It 'reports zero Errors and Warnings when a setting has no checks' {
            $imported = @( New-ImportedMatrix -Settings @( New-Setting -Check @() ) )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.Permissions[0].Errors | Should -Be 0
            $res.Permissions[0].Warnings | Should -Be 0
        }

        It 'produces one permissions row per setting across a single matrix' {
            $imported = @(
                New-ImportedMatrix -Settings @(
                    New-Setting -ComputerName 'PC1'
                    New-Setting -ComputerName 'PC2'
                    New-Setting -ComputerName 'PC3'
                )
            )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.Permissions.Count | Should -Be 3
            $res.Permissions.Computer | Should -Be @('PC1', 'PC2', 'PC3')
        }
    }

    Context 'multiple matrices' {
        It 'aggregates permissions rows from every imported matrix' {
            $imported = @(
                New-ImportedMatrix -FileName 'One.xlsx' -Settings @( New-Setting -ComputerName 'A' )
                New-ImportedMatrix -FileName 'Two.xlsx' -Settings @(
                    New-Setting -ComputerName 'B'
                    New-Setting -ComputerName 'C'
                )
            )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.Permissions.Count | Should -Be 3
            ($res.Permissions | Where-Object MatrixFile -EQ 'One.xlsx').Count | Should -Be 1
            ($res.Permissions | Where-Object MatrixFile -EQ 'Two.xlsx').Count | Should -Be 2
        }

        It 'aggregates FormData rows from every imported matrix' {
            $imported = @(
                New-ImportedMatrix -FileName 'One.xlsx' -FormDataImport @(
                    [pscustomobject]@{ Field = 'A'; Value = '1' }
                )
                New-ImportedMatrix -FileName 'Two.xlsx' -FormDataImport @(
                    [pscustomobject]@{ Field = 'B'; Value = '2' }
                    [pscustomobject]@{ Field = 'C'; Value = '3' }
                )
            )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.FormData.Count | Should -Be 3
        }
    }

    Context 'FormData edge cases' {
        It 'preserves each FormData import row object' {
            $row = [pscustomobject]@{ Field = 'Owner'; Value = 'alice' }
            $imported = @( New-ImportedMatrix -FormDataImport @($row) )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.FormData.Count | Should -Be 1
            $res.FormData[0].Field | Should -Be 'Owner'
            $res.FormData[0].Value | Should -Be 'alice'
        }

        It 'adds no FormData rows when FormData.Import is empty' {
            $imported = @( New-ImportedMatrix -FormDataImport @() )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.FormData.Count | Should -Be 0
        }

        It 'adds no FormData rows when the matrix has no FormData property at all' {
            $imported = @( New-ImportedMatrix -Settings @( New-Setting ) -NoFormDataProperty )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.FormData.Count | Should -Be 0
        }
    }

    Context 'empty and sparse inputs' {
        It 'returns empty collections when a matrix has neither settings nor form data' {
            $imported = @( New-ImportedMatrix -Settings @() -FormDataImport @() )

            $res = Build-ExportDataHC -ImportedMatrix $imported

            $res.Permissions.Count | Should -Be 0
            $res.FormData.Count | Should -Be 0
        }

        It 'always returns an object exposing Permissions and FormData properties' {
            $res = Build-ExportDataHC -ImportedMatrix @( New-ImportedMatrix )

            $res.PSObject.Properties.Name | Should -Contain 'Permissions'
            $res.PSObject.Properties.Name | Should -Contain 'FormData'
        }
    }
}

Describe 'Export-PermissionsFileHC' {
    It 'writes to the "Permissions" worksheet with -AutoSize' {
        Mock Export-Excel

        $rows = @([pscustomobject]@{ A = 1 })
        $path = Join-Path $TestDrive 'perm.xlsx'

        Export-PermissionsFileHC -Rows $rows -Path $path | Out-Null

        Should -Invoke Export-Excel -Times 1 -ParameterFilter {
            $Path -eq $path -and $WorksheetName -eq 'Permissions' -and $AutoSize -eq $true
        }
    }

    It 'returns the path it was given' {
        Mock Export-Excel

        $path = Join-Path $TestDrive 'perm.xlsx'
        $result = Export-PermissionsFileHC -Rows @([pscustomobject]@{ A = 1 }) -Path $path

        $result | Should -Be $path
    }

    It 'forwards the supplied rows to Export-Excel' {
        # Capture the bound parameters in the mock body. Export-Excel's
        # pipeline parameter has differed across ImportExcel versions
        # (-TargetData / -InputObject), so we read whichever one bound rather
        # than naming it. $PesterBoundParameters holds every bound parameter.
        $script:capturedRows = [System.Collections.Generic.List[object]]::new()
        Mock Export-Excel {
            $bp = $PesterBoundParameters
            $key = 'TargetData', 'InputObject' | Where-Object { $bp.ContainsKey($_) } | Select-Object -First 1
            if ($key) { foreach ($r in @($bp[$key])) { $script:capturedRows.Add($r) } }
        }

        $rows = @(
            [pscustomobject]@{ Computer = 'PC1' }
            [pscustomobject]@{ Computer = 'PC2' }
        )
        $path = Join-Path $TestDrive 'perm.xlsx'

        Export-PermissionsFileHC -Rows $rows -Path $path | Out-Null

        $script:capturedRows.Count | Should -Be 2
        $script:capturedRows[0].Computer | Should -Be 'PC1'
        $script:capturedRows[1].Computer | Should -Be 'PC2'
    }

    It 'wraps a failure from Export-Excel in a descriptive terminating error' {
        Mock Export-Excel { throw 'disk full' }

        $path = Join-Path $TestDrive 'perm.xlsx'

        { Export-PermissionsFileHC -Rows @([pscustomobject]@{ A = 1 }) -Path $path } |
        Should -Throw -ExpectedMessage '*Failed exporting Permissions Excel file*disk full*'
    }
}

Describe 'Export-ServiceNowFormDataHC' {
    It 'writes to the "FormData" worksheet with -AutoSize' {
        Mock Export-Excel

        $rows = @([pscustomobject]@{ F = 1 })
        $path = Join-Path $TestDrive 'form.xlsx'

        Export-ServiceNowFormDataHC -Rows $rows -Path $path | Out-Null

        Should -Invoke Export-Excel -Times 1 -ParameterFilter {
            $Path -eq $path -and $WorksheetName -eq 'FormData' -and $AutoSize -eq $true
        }
    }

    It 'returns the path it was given' {
        Mock Export-Excel

        $path = Join-Path $TestDrive 'form.xlsx'
        $result = Export-ServiceNowFormDataHC -Rows @([pscustomobject]@{ F = 1 }) -Path $path

        $result | Should -Be $path
    }

    It 'forwards the supplied rows to Export-Excel' {
        $script:capturedRows = [System.Collections.Generic.List[object]]::new()
        Mock Export-Excel {
            $bp = $PesterBoundParameters
            $key = 'TargetData', 'InputObject' | Where-Object { $bp.ContainsKey($_) } | Select-Object -First 1
            if ($key) { foreach ($r in @($bp[$key])) { $script:capturedRows.Add($r) } }
        }

        $rows = @(
            [pscustomobject]@{ Field = 'Owner' }
            [pscustomobject]@{ Field = 'Team' }
        )
        $path = Join-Path $TestDrive 'form.xlsx'

        Export-ServiceNowFormDataHC -Rows $rows -Path $path | Out-Null

        $script:capturedRows.Count | Should -Be 2
        $script:capturedRows[1].Field | Should -Be 'Team'
    }

    It 'wraps a failure from Export-Excel in a descriptive terminating error' {
        Mock Export-Excel { throw 'locked file' }

        $path = Join-Path $TestDrive 'form.xlsx'

        { Export-ServiceNowFormDataHC -Rows @([pscustomobject]@{ F = 1 }) -Path $path } |
        Should -Throw -ExpectedMessage '*Failed exporting ServiceNow FormData Excel*locked file*'
    }
}

Describe 'Export-OverviewHtmlHC' {
    It 'writes the HTML content to the target path' {
        $path = Join-Path $TestDrive 'overview.html'
        $html = '<html><body>test</body></html>'

        Export-OverviewHtmlHC -Html $html -Path $path | Out-Null

        Test-Path -LiteralPath $path | Should -BeTrue
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'test'
    }

    It 'returns the path it was given' {
        $path = Join-Path $TestDrive 'overview-return.html'

        $result = Export-OverviewHtmlHC -Html '<html></html>' -Path $path

        $result | Should -Be $path
    }

    It 'overwrites an existing file rather than appending' {
        $path = Join-Path $TestDrive 'overview-overwrite.html'
        Set-Content -LiteralPath $path -Value 'OLD CONTENT' -Encoding utf8

        Export-OverviewHtmlHC -Html '<html>NEW</html>' -Path $path | Out-Null

        $content = Get-Content -LiteralPath $path -Raw
        $content | Should -Match 'NEW'
        $content | Should -Not -Match 'OLD CONTENT'
    }

    It 'wraps a write failure in a descriptive terminating error' {
        # Writing into a non-existent directory makes Out-File throw inside the
        # try block, exercising the catch/throw path. We assert only on the
        # stable wrapper prefix the function adds, since the inner OS/binding
        # message varies by platform and PowerShell version.
        $badPath = Join-Path $TestDrive 'no-such-dir\overview.html'

        { Export-OverviewHtmlHC -Html '<html></html>' -Path $badPath } |
        Should -Throw -ExpectedMessage '*Failed exporting Overview HTML file*'
    }
}