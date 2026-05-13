#Requires -Version 7
#requires -Modules Pester

Describe 'New-OverviewHtmlHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Html.ps1"

        function New-FormDataRow {
            param(
                [string]$Category = 'CategoryA',
                [string]$SubCategory = 'SubA',
                [string]$FolderDisplayName = 'TeamA-Share',
                [string]$FilePath = 'C:\Matrix\TeamA.xlsx',
                [string]$FileName = 'TeamA.xlsx',
                [string]$Responsible = 'alice@example.com'
            )

            return [pscustomobject]@{
                MatrixCategoryName      = $Category
                MatrixSubCategoryName   = $SubCategory
                MatrixFolderDisplayName = $FolderDisplayName
                MatrixFilePath          = $FilePath
                MatrixFileName          = $FileName
                MatrixResponsible       = $Responsible
            }
        }
    }

    Context 'basic structure' {
        It 'returns a non-empty string' {
            $html = New-OverviewHtmlHC -FormData @( New-FormDataRow )

            $html | Should -Not -BeNullOrEmpty
            $html | Should -BeOfType [string]
        }

        It 'includes the page title heading' {
            $html = New-OverviewHtmlHC -FormData @( New-FormDataRow )

            $html | Should -Match '<h1>Matrix files overview</h1>'
        }

        It 'includes the table header row with all columns' {
            $html = New-OverviewHtmlHC -FormData @( New-FormDataRow )

            'Category', 'Subcategory', 'Folder', 'Link to the matrix', 'Responsible' |
            ForEach-Object {
                $html | Should -Match "<th>$_</th>"
            }
        }

        It 'wraps the rows in a <tbody>' {
            $html = New-OverviewHtmlHC -FormData @( New-FormDataRow )

            $html | Should -Match '<tbody>'
            $html | Should -Match '</tbody>'
        }
    }

    Context 'row generation' {
        It 'produces one <tr> per FormData row in the tbody' {
            $rows = @(
                New-FormDataRow -Category 'A' -FileName 'one.xlsx'
                New-FormDataRow -Category 'B' -FileName 'two.xlsx'
                New-FormDataRow -Category 'C' -FileName 'three.xlsx'
            )

            $html = New-OverviewHtmlHC -FormData $rows

            # Count <tr> occurrences inside <tbody>...</tbody>
            $tbody = [regex]::Match($html, '(?s)<tbody>(.*?)</tbody>').Groups[1].Value
            ([regex]::Matches($tbody, '<tr>')).Count | Should -Be 3
        }

        It 'renders the data values into the row cells' {
            $row = New-FormDataRow `
                -Category 'Finance' `
                -SubCategory 'Payroll' `
                -FileName 'payroll.xlsx'

            $html = New-OverviewHtmlHC -FormData @($row)

            $html | Should -Match '<td>Finance</td>'
            $html | Should -Match '<td>Payroll</td>'
            $html | Should -Match 'payroll\.xlsx'
        }

        It 'links the matrix file as an anchor' {
            $row = New-FormDataRow `
                -FilePath 'C:\Share\My-Matrix.xlsx' `
                -FileName 'My-Matrix.xlsx'

            $html = New-OverviewHtmlHC -FormData @($row)

            $html | Should -Match '<a href="C:\\Share\\My-Matrix\.xlsx">My-Matrix\.xlsx</a>'
        }

        It 'links the folder as an anchor' {
            $row = New-FormDataRow -FolderDisplayName '\\srv01\teamA'

            $html = New-OverviewHtmlHC -FormData @($row)

            $html | Should -Match '<a href="\\\\srv01\\teamA">'
        }
    }

    Context 'sorting' {
        It 'sorts rows by Category, then SubCategory, then FolderDisplayName' {
            $rows = @(
                New-FormDataRow -Category 'B' -SubCategory 'X' -FolderDisplayName 'F1' -FileName 'b-x-f1.xlsx'
                New-FormDataRow -Category 'A' -SubCategory 'Z' -FolderDisplayName 'F1' -FileName 'a-z-f1.xlsx'
                New-FormDataRow -Category 'A' -SubCategory 'Y' -FolderDisplayName 'F2' -FileName 'a-y-f2.xlsx'
                New-FormDataRow -Category 'A' -SubCategory 'Y' -FolderDisplayName 'F1' -FileName 'a-y-f1.xlsx'
            )

            $html = New-OverviewHtmlHC -FormData $rows

            # Expected sort order: a-y-f1, a-y-f2, a-z-f1, b-x-f1
            $positions = @{
                'a-y-f1.xlsx' = $html.IndexOf('a-y-f1.xlsx')
                'a-y-f2.xlsx' = $html.IndexOf('a-y-f2.xlsx')
                'a-z-f1.xlsx' = $html.IndexOf('a-z-f1.xlsx')
                'b-x-f1.xlsx' = $html.IndexOf('b-x-f1.xlsx')
            }

            $positions['a-y-f1.xlsx'] | Should -BeLessThan $positions['a-y-f2.xlsx']
            $positions['a-y-f2.xlsx'] | Should -BeLessThan $positions['a-z-f1.xlsx']
            $positions['a-z-f1.xlsx'] | Should -BeLessThan $positions['b-x-f1.xlsx']
        }
    }

    Context 'responsible-party emails' {
        It 'renders a mailto: link for a single email' {
            $row = New-FormDataRow -Responsible 'bob@example.com'

            $html = New-OverviewHtmlHC -FormData @($row)

            $html | Should -Match '<a href="mailto:bob@example\.com">bob@example\.com</a>'
        }

        It 'renders one link per email when comma-separated' {
            $row = New-FormDataRow -Responsible 'alice@x.com,bob@x.com'

            $html = New-OverviewHtmlHC -FormData @($row)

            $html | Should -Match 'mailto:alice@x\.com'
            $html | Should -Match 'mailto:bob@x\.com'
        }

        It 'trims whitespace around emails before building the mailto link' {
            $row = New-FormDataRow -Responsible 'alice@x.com, bob@x.com, charlie@x.com'

            $html = New-OverviewHtmlHC -FormData @($row)

            $html | Should -Match '<a href="mailto:bob@x\.com">bob@x\.com</a>'
            $html | Should -Not -Match 'mailto: bob'
            $html | Should -Not -Match 'mailto:%20bob'
        }
    }

    Context 'HTML encoding' {
        It 'encodes ampersands in category names' {
            $row = New-FormDataRow -Category 'R&D'

            $html = New-OverviewHtmlHC -FormData @($row)

            $html | Should -Match '<td>R&amp;D</td>'
            $html | Should -Not -Match '<td>R&D</td>'
        }

        It 'encodes angle brackets to prevent tag injection' {
            $row = New-FormDataRow -Category '<script>alert(1)</script>'

            $html = New-OverviewHtmlHC -FormData @($row)

            $html | Should -Not -Match '<script>alert\(1\)</script>'
            $html | Should -Match '&lt;script&gt;'
        }
    }

    Context 'empty input' {
        It 'returns a valid HTML page when FormData is empty' {
            $html = New-OverviewHtmlHC -FormData @()

            $html | Should -Match '<h1>Matrix files overview</h1>'
            $html | Should -Match '<tbody>'
            $html | Should -Match '</tbody>'

            # No data rows
            $tbody = [regex]::Match($html, '(?s)<tbody>(.*?)</tbody>').Groups[1].Value
            ([regex]::Matches($tbody, '<tr>')).Count | Should -Be 0
        }
    }
}

Describe 'Html.ps1 consolidated tests' {

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Html.ps1"
        . "$moduleRoot\Private\Utils.ps1"
    }

    It 'Initializes HTML structure' {
        $h = Initialize-HtmlStructureHC
        $h.Style | Should -Match '<style'
        $h.Templates.SettingsHeader | Should -Match 'Settings'
    }

    It 'Maps error types to CSS classes' {
        Get-HtmlClassProbTypeHC 'FatalError' | Should -Be 'probTypeError'
        Get-HtmlClassProbTypeHC 'Warning' | Should -Be 'probTypeWarning'
    }

    It 'Builds matrix email HTML' {
        $html = Initialize-HtmlStructureHC

        $matrix = @(
            [pscustomobject]@{
                File        = @{
                    Item         = @{ Name = 'A.xlsx' }
                    SaveFullName = 'C:\A.xlsx'
                    ExcelInfo    = @{ LastModifiedBy = 'User'; Modified = Get-Date }
                }
                FormData    = @{ Check = @() }
                Permissions = @{ Check = @() }
                Settings    = @()
            }
        )

        $out = Build-MatrixEmailHtmlHC $matrix $html
        $out | Should -Match 'A.xlsx'
    }

    It 'Writes troubleshooting log' {
        $html = Initialize-HtmlStructureHC
        $folder = Join-Path $TestDrive 'logs'
        New-Item -ItemType Directory -Path $folder | Out-Null

        $matrix = [pscustomobject]@{
            File        = @{ LogFolder = $folder; Check = @() }
            FormData    = @{ Check = @() }
            Permissions = @{ Check = @() }
        }

        $path = Write-MatrixTroubleshootingLogHC $matrix $html
        Test-Path $path | Should -BeTrue
    }
}

