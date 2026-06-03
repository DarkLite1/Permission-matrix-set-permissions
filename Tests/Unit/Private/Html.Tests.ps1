#Requires -Version 7
#requires -Modules Pester

BeforeAll {
    # Load the module code to test
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
    ForEach-Object { . $_.FullName }
}

Describe 'New-OverviewHtmlHC' {
    BeforeAll {
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

            # The anchor carries target/rel attributes; the href is the raw
            # file path and the link text is the (encoded) file name.
            $html | Should -Match '<a href="C:\\Share\\My-Matrix\.xlsx" target=''_blank'' rel=''noopener noreferrer'' >My-Matrix\.xlsx</a>'
        }

        It 'links the folder as an anchor' {
            $row = New-FormDataRow -FolderDisplayName '\\srv01\teamA'

            $html = New-OverviewHtmlHC -FormData @($row)

            $html | Should -Match '<a href="\\\\srv01\\teamA" target=''_blank'' rel=''noopener noreferrer'' >'
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

            $html | Should -Not -Match '<td><script>alert\(1\)</script></td>'
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

Describe 'Initialize-HtmlStructureHC' {

    BeforeEach {
        $script:struct = Initialize-HtmlStructureHC
    }

    It 'returns a hashtable' {
        $struct | Should -BeOfType [hashtable]
    }

    It 'exposes Style as a non-empty <style> block' {
        $struct.Style | Should -Not -BeNullOrEmpty
        $struct.Style | Should -Match '<style type="text/css">'
        $struct.Style | Should -Match '</style>'
    }

    It 'exposes TroubleshootingStyle as a separate <style> block' {
        $struct.TroubleshootingStyle | Should -Not -BeNullOrEmpty
        $struct.TroubleshootingStyle | Should -Match '<style type="text/css">'
    }

    It 'exposes a Templates hashtable' {
        $struct.Templates | Should -BeOfType [hashtable]
        $struct.Templates.ContainsKey('SettingsHeader') | Should -BeTrue
        $struct.Templates.ContainsKey('LegendTable') | Should -BeTrue
    }

    It 'keeps Templates.SettingsHeader as an empty no-op placeholder' {
        # The modern layout embeds settings per file card; this template is
        # retained only for backward compatibility with external callers.
        $struct.Templates.SettingsHeader | Should -BeNullOrEmpty
    }

    It 'keeps Templates.LegendTable as an empty no-op placeholder' {
        $struct.Templates.LegendTable | Should -BeNullOrEmpty
    }

    It 'embeds the themed page background colour in the Style block' {
        $struct.Style | Should -Match 'background-color: #e5e7eb;'
    }
}

Describe 'Get-HtmlClassProbTypeHC' {
    It 'maps <Type> to <Expected>' -ForEach @(
        @{ Type = 'FatalError'; Expected = 'probTypeError' }
        @{ Type = 'Warning'; Expected = 'probTypeWarning' }
        @{ Type = 'Information'; Expected = 'probTypeInfo' }
        @{ Type = ''; Expected = 'probTypeInfo' }
        @{ Type = 'unknown'; Expected = 'probTypeInfo' }
    ) {
        Get-HtmlClassProbTypeHC $Type | Should -Be $Expected
    }
}

Describe 'Format-IssueCountLabelHC' {
    It 'returns "Success" when there are no errors or warnings' {
        Format-IssueCountLabelHC -Errors 0 -Warnings 0 | Should -Be 'Success'
    }

    It 'singularises a single error' {
        Format-IssueCountLabelHC -Errors 1 -Warnings 0 | Should -Be '1 Error'
    }

    It 'pluralises multiple errors' {
        Format-IssueCountLabelHC -Errors 3 -Warnings 0 | Should -Be '3 Errors'
    }

    It 'singularises a single warning' {
        Format-IssueCountLabelHC -Errors 0 -Warnings 1 | Should -Be '1 Warning'
    }

    It 'pluralises multiple warnings' {
        Format-IssueCountLabelHC -Errors 0 -Warnings 2 | Should -Be '2 Warnings'
    }

    It 'joins errors and warnings with a comma' {
        Format-IssueCountLabelHC -Errors 2 -Warnings 1 | Should -Be '2 Errors, 1 Warning'
    }
}

Describe 'Format-LastChangeHC' {
    It 'combines user and date when both are known' {
        $dt = Get-Date '2026-05-19 13:30:00'
        Format-LastChangeHC -LastModifiedBy 'Brecht' -Modified $dt |
        Should -Be 'Last change: Brecht &middot; 19/05/2026 13:30'
    }

    It 'shows only the user when the date is missing' {
        Format-LastChangeHC -LastModifiedBy 'Brecht' -Modified $null |
        Should -Be 'Last change: Brecht'
    }

    It 'shows only the date when the user is missing' {
        $dt = Get-Date '2026-05-19 13:30:00'
        Format-LastChangeHC -LastModifiedBy '' -Modified $dt |
        Should -Be 'Last change: 19/05/2026 13:30'
    }

    It 'treats the literal "Unknown" username as missing' {
        Format-LastChangeHC -LastModifiedBy 'Unknown' -Modified $null |
        Should -Be 'No modification metadata available'
    }

    It 'returns a placeholder when neither value is known' {
        Format-LastChangeHC -LastModifiedBy '' -Modified $null |
        Should -Be 'No modification metadata available'
    }

    It 'treats a non-datetime Modified value as missing' {
        Format-LastChangeHC -LastModifiedBy 'Brecht' -Modified 'not-a-date' |
        Should -Be 'Last change: Brecht'
    }

    It 'treats DateTime.MinValue as missing' {
        Format-LastChangeHC -LastModifiedBy 'Brecht' -Modified ([datetime]::MinValue) |
        Should -Be 'Last change: Brecht'
    }

    It 'HTML-encodes the username component' {
        Format-LastChangeHC -LastModifiedBy 'A&B' -Modified $null |
        Should -Be 'Last change: A&amp;B'
    }

    It 'uses HH:mm (not seconds) for the time component' {
        $dt = Get-Date '2026-05-19 13:30:45'
        Format-LastChangeHC -LastModifiedBy '' -Modified $dt |
        Should -Match '13:30$'
    }
}

Describe 'ConvertTo-FileUrlHC' {
    It 'returns empty string for null or whitespace input' {
        ConvertTo-FileUrlHC -Path $null | Should -Be ''
        ConvertTo-FileUrlHC -Path '   ' | Should -Be ''
    }

    It 'prefixes file:// and converts backslashes to forward slashes' {
        ConvertTo-FileUrlHC -Path 'C:\share\budget.xlsx' |
        Should -Be 'file://C:/share/budget.xlsx'
    }

    It 'percent-encodes spaces' {
        ConvertTo-FileUrlHC -Path 'C:\my files\a b.xlsx' |
        Should -Be 'file://C:/my%20files/a%20b.xlsx'
    }

    It 'converts UNC paths' {
        ConvertTo-FileUrlHC -Path '\\srv01\teamA\m.xlsx' |
        Should -Be 'file:////srv01/teamA/m.xlsx'
    }
}

Describe 'Get-CheckThemeHC' {
    It 'returns the error theme for FatalError' {
        $t = Get-CheckThemeHC 'FatalError'
        $t.Label | Should -Be 'ERROR'
        $t.Symbol | Should -Be '✖'
        $t.Accent | Should -Be '#dc2626'
    }

    It 'returns the warning theme for Warning' {
        $t = Get-CheckThemeHC 'Warning'
        $t.Label | Should -Be 'WARNING'
        $t.Symbol | Should -Be '⚠'
        $t.Accent | Should -Be '#d97706'
    }

    It 'returns the info theme for any other value' {
        $t = Get-CheckThemeHC 'Information'
        $t.Label | Should -Be 'INFO'
        $t.Symbol | Should -Be 'ℹ'
        $t.Accent | Should -Be '#2563eb'
    }
}

Describe 'Get-TruncatedPathHC' {
    It 'returns the path unchanged and $false when within the limit' {
        $result = Get-TruncatedPathHC -Path 'C:\short\path.txt' -MaxChars 64
        $result[0] | Should -Be 'C:\short\path.txt'
        $result[1] | Should -BeFalse
    }

    It 'returns empty input unchanged' {
        $result = Get-TruncatedPathHC -Path '' -MaxChars 10
        $result[0] | Should -Be ''
        $result[1] | Should -BeFalse
    }

    It 'truncates with an ellipsis and reports $true when over the limit' {
        $long = 'C:\a\very\deeply\nested\folder\structure\with\a\file.txt'
        $result = Get-TruncatedPathHC -Path $long -MaxChars 32
        $result[1] | Should -BeTrue
        $result[0] | Should -Match '\\\.\.\.\\'
        $result[0].Length | Should -BeLessOrEqual ($long.Length)
    }

    It 'keeps the final segment visible after truncation' {
        $long = 'C:\a\very\deeply\nested\folder\structure\with\a\report.txt'
        $result = Get-TruncatedPathHC -Path $long -MaxChars 32
        $result[0] | Should -Match 'report\.txt$'
    }
}

Describe 'New-PillHtmlHC' {
    It 'returns empty string for blank text' {
        New-PillHtmlHC -Text '' -Bg '#000000' | Should -Be ''
        New-PillHtmlHC -Text '   ' -Bg '#000000' | Should -Be ''
    }

    It 'renders a span with the supplied text and background colour' {
        $pill = New-PillHtmlHC -Text 'Error' -Bg '#dc2626'
        $pill | Should -Match '<span'
        $pill | Should -Match 'background-color:#dc2626;'
        $pill | Should -Match '>Error</span>'
    }

    It 'defaults the text colour to white' {
        $pill = New-PillHtmlHC -Text 'OK' -Bg '#16a34a'
        $pill | Should -Match 'color:#ffffff;'
    }

    It 'honours a custom text colour' {
        $pill = New-PillHtmlHC -Text 'OK' -Bg '#16a34a' -Color '#000000'
        $pill | Should -Match 'color:#000000;'
    }
}

Describe 'Build-ErrorWarningTableHC' {
    It 'returns empty string when there are no issues' {
        $counter = [pscustomobject]@{ TotalErrors = 0; TotalWarnings = 0 }
        Build-ErrorWarningTableHC -CounterData $counter | Should -Be ''
    }

    It 'renders an error pill when there are errors' {
        $counter = [pscustomobject]@{ TotalErrors = 2; TotalWarnings = 0 }
        $html = Build-ErrorWarningTableHC -CounterData $counter
        $html | Should -Match 'Detected issues'
        $html | Should -Match '2 Errors'
    }

    It 'renders a warning pill when there are warnings' {
        $counter = [pscustomobject]@{ TotalErrors = 0; TotalWarnings = 1 }
        $html = Build-ErrorWarningTableHC -CounterData $counter
        $html | Should -Match '1 Warning'
    }

    It 'renders both pills when there are errors and warnings' {
        $counter = [pscustomobject]@{ TotalErrors = 1; TotalWarnings = 3 }
        $html = Build-ErrorWarningTableHC -CounterData $counter
        $html | Should -Match '1 Error'
        $html | Should -Match '3 Warnings'
    }
}

Describe 'Build-SystemErrorsBlockHC' {
    It 'returns empty string for a null or empty list' {
        Build-SystemErrorsBlockHC -SystemErrors @() | Should -Be ''
        Build-SystemErrorsBlockHC -SystemErrors $null | Should -Be ''
    }

    It 'ignores items that are neither FatalError nor Warning' {
        $items = @(
            [pscustomobject]@{ Type = 'Information'; Name = 'note'; Message = 'fyi' }
        )
        Build-SystemErrorsBlockHC -SystemErrors $items | Should -Be ''
    }

    It 'renders a System Error card for a FatalError item' {
        $items = @(
            [pscustomobject]@{ Type = 'FatalError'; Name = 'Boom'; Message = 'it broke'; Category = 'Matrix' }
        )
        $html = Build-SystemErrorsBlockHC -SystemErrors $items
        $html | Should -Match 'System Error'
        $html | Should -Match 'Boom'
        $html | Should -Match 'it broke'
        $html | Should -Match '1 Error'
    }

    It 'renders a System Warning card for a Warning item' {
        $items = @(
            [pscustomobject]@{ Type = 'Warning'; Name = 'Careful'; Message = 'heads up'; Category = '' }
        )
        $html = Build-SystemErrorsBlockHC -SystemErrors $items
        $html | Should -Match 'System Warning'
        $html | Should -Match '1 Warning'
    }

    It 'HTML-encodes the item name' {
        $items = @(
            [pscustomobject]@{ Type = 'FatalError'; Name = 'a&b'; Message = 'm'; Category = '' }
        )
        $html = Build-SystemErrorsBlockHC -SystemErrors $items
        $html | Should -Match 'a&amp;b'
    }
}

Describe 'Build-FileLevelCheckRowHC' {
    It 'renders the check name, description and sheet label' {
        $check = [pscustomobject]@{ Type = 'Warning'; Name = 'CheckName'; Description = 'CheckDesc' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'Excel File'
        $html | Should -Match 'CheckName'
        $html | Should -Match 'CheckDesc'
        $html | Should -Match 'Excel File'
    }

    It 'uses the error accent colour for FatalError checks' {
        $check = [pscustomobject]@{ Type = 'FatalError'; Name = 'n'; Description = 'd' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'Excel File'
        $html | Should -Match '#dc2626'
        $html | Should -Match 'ERROR'
    }

    It 'includes the 16px inset wrapper by default' {
        $check = [pscustomobject]@{ Type = 'Warning'; Name = 'n'; Description = 'd' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'X'
        $html | Should -Match 'padding:0 16px 8px 16px;'
    }

    It 'omits the inset wrapper when -IncludeWrapper is $false' {
        $check = [pscustomobject]@{ Type = 'Warning'; Name = 'n'; Description = 'd' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'X' -IncludeWrapper $false
        $html | Should -Match 'padding:0 0 8px 0;'
    }

    It 'falls back to placeholder text for a missing name' {
        $check = [pscustomobject]@{ Type = 'Warning'; Name = ''; Description = 'd' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'X'
        $html | Should -Match 'Unnamed check'
    }
}

Describe 'Build-SettingsRowHC' {
    BeforeAll {
        function New-MatrixItem {
            param(
                [string]$ComputerName = 'SRV01',
                [string]$Path = 'D:\data',
                [string]$Action = 'Apply',
                [object[]]$Check = @(),
                [string]$ReportFilePath = ''
            )
            return [pscustomobject]@{
                ID          = 1
                Check       = $Check
                Setting     = [pscustomobject]@{
                    Formatted = [pscustomobject]@{
                        ComputerName = $ComputerName
                        Path         = $Path
                        Action       = $Action
                    }
                }
                JobTime     = [pscustomobject]@{ Duration = $null }
                FileContext = [pscustomobject]@{ ReportFilePath = $ReportFilePath }
            }
        }
    }

    It 'renders the computer name and path' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem -ComputerName 'SRV99' -Path 'E:\foo')
        $html | Should -Match 'SRV99'
        $html | Should -Match 'E:\\foo'
    }

    It 'shows an Error pill when the row has a FatalError check' {
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'FatalError' })
        $html = Build-SettingsRowHC -MatrixItem $item
        $html | Should -Match '>Error</span>'
        $html | Should -Match '#dc2626'
    }

    It 'shows a Warning pill when the row has a Warning check' {
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'Warning' })
        $html = Build-SettingsRowHC -MatrixItem $item
        $html | Should -Match '>Warning</span>'
        $html | Should -Match '#d97706'
    }

    It 'uses the success accent and no status pill for a clean row' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem)
        $html | Should -Match '#16a34a'
        $html | Should -Not -Match 'rr-srow-status'
    }

    It 'shows N/A for a missing duration' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem)
        $html | Should -Match 'N/A'
    }

    It 'links the row to the report file path when present' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem -ReportFilePath 'C:\logs\r.html')
        $html | Should -Match "href='C:\\logs\\r\.html'"
    }
}

Describe 'Build-MatrixEmailHtmlHC' {
    BeforeAll {
        function New-FileResult {
            param(
                [string]$Name = 'A.xlsx',
                [string]$FullName = 'C:\A.xlsx',
                [string]$LastModifiedBy = 'User',
                [datetime]$Modified = (Get-Date '2024-01-15 09:30:00'),
                [object[]]$Check = @(),
                [object[]]$FormDataCheck = @(),
                [object[]]$PermissionsCheck = @(),
                [object[]]$Matrices = @()
            )

            return [pscustomobject]@{
                Item      = [pscustomobject]@{ Name = $Name; FullName = $FullName }
                ExcelInfo = [pscustomobject]@{
                    LastModifiedBy = $LastModifiedBy
                    Modified       = $Modified
                }
                Check     = $Check
                Sheets    = [pscustomobject]@{
                    FormData    = [pscustomobject]@{ Check = $FormDataCheck }
                    Permissions = [pscustomobject]@{ Check = $PermissionsCheck }
                }
                Matrices  = $Matrices
            }
        }

        function New-MatrixRow {
            param(
                [int]$ID = 1,
                [string]$ComputerName = 'SRV01',
                [object[]]$Check = @()
            )
            return [pscustomobject]@{
                ID      = $ID
                Check   = $Check
                Setting = [pscustomobject]@{
                    Formatted = [pscustomobject]@{ ComputerName = $ComputerName; Path = ''; Action = '' }
                }
                JobTime = [pscustomobject]@{ Duration = $null }
            }
        }
    }

    BeforeEach {
        $script:html = Initialize-HtmlStructureHC
    }

    Context 'basic file rendering' {
        It 'renders the filename in the title link text' {
            $files = @( New-FileResult -Name 'Q3-Permissions.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'Q3-Permissions\.xlsx'
        }

        It 'uses a file:// URL derived from Item.FullName as the title link href' {
            $files = @( New-FileResult -FullName 'C:\share\budget.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            # ConvertTo-FileUrlHC turns the Windows path into a file:// URL with
            # forward slashes; the anchor also carries a title tooltip.
            $out | Should -Match '<a href="file://C:/share/budget\.xlsx"'
        }

        It 'puts the raw Windows path in the title tooltip of the header link' {
            $files = @( New-FileResult -FullName 'C:\share\budget.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'title="C:\\share\\budget\.xlsx"'
        }

        It 'renders one file card table per file result' {
            $files = @(
                New-FileResult -Name 'one.xlsx'
                New-FileResult -Name 'two.xlsx'
                New-FileResult -Name 'three.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            # Each card is anchored by its "Open full report" footer link.
            ([regex]::Matches($out, 'Open full report')).Count | Should -Be 3
        }
    }

    Context 'ExcelInfo handling' {
        It 'renders LastModifiedBy in the file info row' {
            $files = @( New-FileResult -LastModifiedBy 'alice@example.com' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'alice@example\.com'
        }

        It 'omits the user but keeps the date when LastModifiedBy is empty' {
            $files = @( New-FileResult -LastModifiedBy '' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            # Format-LastChangeHC drops the user and renders date-only.
            $out | Should -Match 'Last change: 15/01/2024 09:30'
            $out | Should -Not -Match 'Last change: Unknown'
        }

        It 'formats Modified as dd/MM/yyyy HH:mm' {
            $files = @(
                New-FileResult -Modified (Get-Date '2024-03-22 14:05:09')
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            # The layout uses minute precision, not seconds.
            $out | Should -Match '22/03/2024 14:05'
            $out | Should -Not -Match '22/03/2024 14:05:09'
        }

        It 'shows only the user when Modified is not a datetime' {
            $fr = New-FileResult
            # Overwrite Modified with a non-datetime value
            $fr.ExcelInfo.Modified = 'not-a-date'

            $out = Build-MatrixEmailHtmlHC -FileResults @($fr) -Html $html

            $out | Should -Match 'Last change: User'
            $out | Should -Not -Match 'Last change: User &middot;'
        }

        It 'HTML-encodes the filename' {
            $files = @( New-FileResult -Name 'a&b<c>.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'a&amp;b&lt;c&gt;\.xlsx'
            $out | Should -Not -Match '<c>'
        }
    }

    Context 'header status' {
        It 'shows a success header for a file with no issues' {
            $out = Build-MatrixEmailHtmlHC -FileResults @( New-FileResult ) -Html $html
            $out | Should -Match '✓'
            $out | Should -Match 'Success'
        }

        It 'shows a warning header when a matrix row has a Warning' {
            $files = @(
                New-FileResult -Matrices @(
                    New-MatrixRow -Check @([pscustomobject]@{ Type = 'Warning'; Name = 'w'; Description = 'd' })
                )
            )
            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html
            $out | Should -Match '⚠'
            $out | Should -Match '1 Warning'
        }

        It 'shows an error header when a file-level check is a FatalError' {
            $files = @( New-FileResult -Check @([pscustomobject]@{ Type = 'FatalError'; Name = 'e'; Description = 'd' }) )
            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html
            $out | Should -Match '✖'
            $out | Should -Match '1 Error'
        }
    }

    Context 'matrices section' {
        It 'renders a Settings section header with the matrix count' {
            $files = @(
                New-FileResult -Matrices @(
                    New-MatrixRow -ID 1 -ComputerName 'SRV01'
                    New-MatrixRow -ID 2 -ComputerName 'SRV02'
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'Settings \(2\)'
            $out | Should -Match 'SRV01'
            $out | Should -Match 'SRV02'
        }

        It 'shows the empty-state message when there are no matrices and no issues' {
            $files = @( New-FileResult -Matrices @() )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'No settings rows were processed for this file\.'
        }

        It 'renders a File Issues section when a file-level check exists' {
            $files = @(
                New-FileResult -Check @(
                    [pscustomobject]@{ Type = 'Warning'; Name = 'fileCheck'; Description = 'desc' }
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'File Issues \(1\)'
            $out | Should -Match 'fileCheck'
        }
    }
}

Describe 'Get-MailBodyHtmlHC' {
    BeforeEach {
        $script:html = Initialize-HtmlStructureHC
    }

    It 'returns a complete HTML document' {
        $settings = [pscustomobject]@{ ScriptName = 'My Script'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')
        $out | Should -Match '<!DOCTYPE html>'
        $out | Should -Match '</html>'
    }

    It 'renders the script name in an encoded h1' {
        $settings = [pscustomobject]@{ ScriptName = 'R&D Run'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')
        $out | Should -Match '<h1>R&amp;D Run</h1>'
    }

    It 'falls back to a default script name when none is supplied' {
        $settings = [pscustomobject]@{ ScriptName = ''; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')
        $out | Should -Match '<h1>Permission Matrix</h1>'
    }

    It 'renders a footer with Started, Ended and Duration when a start time is given' {
        $settings = [pscustomobject]@{ ScriptName = 'S'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:30:00')
        $out | Should -Match 'Started'
        $out | Should -Match 'Ended'
        $out | Should -Match 'Duration'
        $out | Should -Match '00:30:00'
    }

    It 'includes the MatrixTables fragment passed via the Html hashtable' {
        $html.MatrixTables = '<!-- MATRIX_TABLES_MARKER -->'
        $settings = [pscustomobject]@{ ScriptName = 'S'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')
        $out | Should -Match 'MATRIX_TABLES_MARKER'
    }

    It 'renders system error cards from a [ref] SystemErrors entry' {
        $errors = [System.Collections.Generic.List[object]]::new()
        $errors.Add([pscustomobject]@{ Type = 'FatalError'; Name = 'SysBoom'; Message = 'bad'; Category = 'Matrix' })
        $html.SystemErrors = ([ref]$errors)
        $settings = [pscustomobject]@{ ScriptName = 'S'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')
        $out | Should -Match 'SysBoom'
        $out | Should -Match 'System Error'
    }
}

Describe 'Build-ExecutionDetailsBlockHC' {
    BeforeAll {
        function New-FileResultForDetails {
            param(
                [string]$FullName = 'C:\data\matrix.xlsx',
                [string]$LastModifiedBy = 'User',
                [datetime]$Modified = (Get-Date '2024-01-15 09:30:00')
            )
            return [pscustomobject]@{
                Item      = [pscustomobject]@{ FullName = $FullName }
                ExcelInfo = [pscustomobject]@{ LastModifiedBy = $LastModifiedBy; Modified = $Modified }
            }
        }
    }

    It 'renders the matrix file path as a file:// link' {
        $fr = New-FileResultForDetails -FullName 'C:\data\matrix.xlsx'
        $html = Build-ExecutionDetailsBlockHC -FileResult $fr `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:05:00')
        $html | Should -Match 'href="file://C:/data/matrix\.xlsx"'
        $html | Should -Match 'Matrix file'
    }

    It 'renders the defaults file when a path is supplied' {
        $fr = New-FileResultForDetails
        $html = Build-ExecutionDetailsBlockHC -FileResult $fr -DefaultsFilePath 'C:\d\defaults.json' `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:05:00')
        $html | Should -Match 'Defaults file'
        $html | Should -Match 'defaults\.json'
    }

    It 'skips the defaults file row when no path is supplied' {
        $fr = New-FileResultForDetails
        $html = Build-ExecutionDetailsBlockHC -FileResult $fr -DefaultsFilePath '' `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:05:00')
        $html | Should -Not -Match 'Defaults file'
    }

    It 'formats start and end times with seconds precision' {
        $fr = New-FileResultForDetails
        $html = Build-ExecutionDetailsBlockHC -FileResult $fr `
            -ScriptStartTime (Get-Date '2024-03-22 14:05:09') `
            -ScriptEndTime (Get-Date '2024-03-22 14:06:11')
        $html | Should -Match '22/03/2024 14:05:09'
        $html | Should -Match '22/03/2024 14:06:11'
    }

    It 'strips the "Last change:" prefix from the last-change value' {
        $fr = New-FileResultForDetails -LastModifiedBy 'Brecht' -Modified (Get-Date '2024-01-15 09:30:00')
        $html = Build-ExecutionDetailsBlockHC -FileResult $fr `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:05:00')
        $html | Should -Match 'Last change'
        $html | Should -Match 'Brecht'
    }
}

Describe 'Build-MatrixDetailCardHC' {
    BeforeAll {
        function New-DetailMatrix {
            param(
                [object]$ID = 42,
                [string]$ComputerName = 'SRV01',
                [string]$Path = 'D:\share',
                [string]$Action = 'Apply',
                [object[]]$Check = @()
            )
            return [pscustomobject]@{
                ID      = $ID
                Check   = $Check
                Setting = [pscustomobject]@{
                    Formatted = [pscustomobject]@{
                        ComputerName = $ComputerName
                        Path         = $Path
                        Action       = $Action
                    }
                }
                JobTime = [pscustomobject]@{ Duration = $null }
            }
        }
    }

    It 'renders a compact (header-only) card for a clean matrix row' {
        $html = Build-MatrixDetailCardHC -MatrixItem (New-DetailMatrix)
        $html | Should -Match 'SRV01'
        # No check rows for a clean card
        $html | Should -Not -Match 'WARNING'
        $html | Should -Not -Match 'ERROR'
    }

    It 'renders check rows in full mode when the row has checks' {
        $item = New-DetailMatrix -Check @(
            [pscustomobject]@{ Type = 'Warning'; Name = 'TestCheck'; Description = 'A test description' }
        )
        $html = Build-MatrixDetailCardHC -MatrixItem $item
        $html | Should -Match 'TestCheck'
        $html | Should -Match 'A test description'
        $html | Should -Match 'WARNING'
    }

    It 'shortens a long ID in the visible cell but keeps the full ID in the tooltip' {
        $item = New-DetailMatrix -ID '1234567890ABCDEF'
        $html = Build-MatrixDetailCardHC -MatrixItem $item
        $html | Should -Match '123\.\.\.DEF'
        $html | Should -Match 'title="1234567890ABCDEF"'
    }

    It 'links a check name to its JSON file when JsonFileName is present' {
        $item = New-DetailMatrix -Check @(
            [pscustomobject]@{ Type = 'FatalError'; Name = 'JCheck'; Description = 'd'; JsonFileName = 'C:\j\check.json' }
        )
        $html = Build-MatrixDetailCardHC -MatrixItem $item
        $html | Should -Match "href='C:\\j\\check\.json'"
        $html | Should -Match '>JCheck</a>'
    }
}

Describe 'New-HtmlSectionHC' {
    It 'returns empty when there are no checks and no title' {
        New-HtmlSectionHC -Title '' -Checks @() | Should -Be ''
    }

    It 'renders an encoded section title heading' {
        $html = New-HtmlSectionHC -Title 'R&D Section' -Checks @()
        $html | Should -Match 'R&amp;D Section'
    }

    It 'renders one file-level check row per check' {
        $checks = @(
            [pscustomobject]@{ Type = 'Warning'; Name = 'c1'; Description = 'd1' }
            [pscustomobject]@{ Type = 'FatalError'; Name = 'c2'; Description = 'd2' }
        )
        $html = New-HtmlSectionHC -Title 'Sec' -Checks $checks
        $html | Should -Match 'c1'
        $html | Should -Match 'c2'
    }
}

Describe 'New-HtmlCheckRowHC' {
    It 'renders a row with the prob-type class for the check type' {
        $check = [pscustomobject]@{ Type = 'FatalError'; Name = 'n'; Description = 'd' }
        $html = New-HtmlCheckRowHC -CheckItem $check
        $html | Should -Match "class='probTypeError'"
        $html | Should -Match '>n</td>'
        $html | Should -Match '>d</td>'
    }

    It 'HTML-encodes the name and description' {
        $check = [pscustomobject]@{ Type = 'Warning'; Name = 'a&b'; Description = '<x>' }
        $html = New-HtmlCheckRowHC -CheckItem $check
        $html | Should -Match 'a&amp;b'
        $html | Should -Match '&lt;x&gt;'
    }
}

Describe 'New-SettingsCardHtmlHC' {
    It 'delegates to Build-MatrixDetailCardHC' {
        $item = [pscustomobject]@{
            ID      = 7
            Check   = @()
            Setting = [pscustomobject]@{
                Formatted = [pscustomobject]@{ ComputerName = 'SRV-DELEGATE'; Path = ''; Action = '' }
            }
            JobTime = [pscustomobject]@{ Duration = $null }
        }
        $html = New-SettingsCardHtmlHC -MatrixItem $item
        $html | Should -Match 'SRV-DELEGATE'
    }
}

Describe 'New-SettingsOverviewHtmlHC' {
    It 'is a no-op that returns empty string in the modern layout' {
        New-SettingsOverviewHtmlHC -MatrixRows @() -Html @{} | Should -Be ''
    }
}

Describe 'Write-MatrixExecutionReportHC' {
    BeforeEach {
        $script:html = Initialize-HtmlStructureHC
        $script:logFolder = Join-Path $TestDrive 'reportlogs'
        New-Item -ItemType Directory -Path $logFolder -Force | Out-Null

        $script:fileResult = [pscustomobject]@{
            Item      = [pscustomobject]@{ Name = 'Report.xlsx'; FullName = 'C:\Report.xlsx' }
            ExcelInfo = [pscustomobject]@{ LastModifiedBy = 'User'; Modified = (Get-Date '2024-01-15 09:30:00') }
            Check     = @()
            Sheets    = [pscustomobject]@{
                FormData    = [pscustomobject]@{ Check = @() }
                Permissions = [pscustomobject]@{ Check = @() }
            }
            Matrices  = @(
                [pscustomobject]@{
                    ID      = 1
                    Check   = @([pscustomobject]@{ Type = 'Warning'; Name = 'C'; Description = 'D' })
                    Setting = [pscustomobject]@{
                        Formatted = [pscustomobject]@{ ComputerName = 'SRV01'; Path = ''; Action = '' }
                    }
                    JobTime = [pscustomobject]@{ Duration = $null }
                }
            )
        }
    }

    It 'writes "00 - Execution Report.html" into the log folder' {
        Write-MatrixExecutionReportHC -FileResult $fileResult -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:05:00') `
            -LogFolder $logFolder

        Test-Path -LiteralPath (Join-Path $logFolder '00 - Execution Report.html') | Should -BeTrue
    }

    It 'writes a valid HTML document with the filename in the title' {
        Write-MatrixExecutionReportHC -FileResult $fileResult -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:05:00') `
            -LogFolder $logFolder

        $content = Get-Content -LiteralPath (Join-Path $logFolder '00 - Execution Report.html') -Raw
        $content | Should -Match '<!DOCTYPE html>'
        $content | Should -Match '<title>Execution Report - Report\.xlsx</title>'
    }

    It 'includes a Settings section and the matrix detail card' {
        Write-MatrixExecutionReportHC -FileResult $fileResult -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:05:00') `
            -LogFolder $logFolder

        $content = Get-Content -LiteralPath (Join-Path $logFolder '00 - Execution Report.html') -Raw
        $content | Should -Match 'Settings \(1\)'
        $content | Should -Match 'SRV01'
    }

    It 'returns null when the log folder does not exist' {
        $missing = Join-Path $TestDrive 'no-such-folder'
        $result = Write-MatrixExecutionReportHC -FileResult $fileResult -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:05:00') `
            -LogFolder $missing
        $result | Should -BeNullOrEmpty
    }
}

Describe 'Write-MatrixSettingLogHC' {
    BeforeEach {
        $script:html = Initialize-HtmlStructureHC
        $script:logFolder = Join-Path $TestDrive 'logs'
        New-Item -ItemType Directory -Path $logFolder -Force | Out-Null

        $script:matrix = [pscustomobject]@{
            ID      = 42
            Check   = @(
                [pscustomobject]@{
                    Type        = 'Warning'
                    Name        = 'TestCheck'
                    Description = 'A test description'
                }
            )
            Setting = [pscustomobject]@{
                Formatted = [pscustomobject]@{ ComputerName = ''; Path = ''; Action = '' }
            }
            JobTime = [pscustomobject]@{ Duration = $null }
        }
    }

    It 'writes a file named "ID <id> - Settings.html" in the log folder' {
        Write-MatrixSettingLogHC -Matrix $matrix -Html $html -LogFolder $logFolder

        $expected = Join-Path $logFolder 'ID 42 - Settings.html'
        Test-Path -LiteralPath $expected | Should -BeTrue
    }

    It 'writes a valid HTML document with the matrix ID in the title' {
        Write-MatrixSettingLogHC -Matrix $matrix -Html $html -LogFolder $logFolder

        $path = Join-Path $logFolder 'ID 42 - Settings.html'
        $content = Get-Content -LiteralPath $path -Raw

        $content | Should -Match '<!DOCTYPE html>'
        # The heading uses an em-dash entity, not a hyphen.
        $content | Should -Match '<h1>Settings Log &mdash; ID 42</h1>'
    }

    It 'renders the check detail card for the matrix' {
        Write-MatrixSettingLogHC -Matrix $matrix -Html $html -LogFolder $logFolder

        $path = Join-Path $logFolder 'ID 42 - Settings.html'
        $content = Get-Content -LiteralPath $path -Raw

        $content | Should -Match 'TestCheck'
        $content | Should -Match 'A test description'
    }

    It 'returns null when the log folder does not exist' {
        $missing = Join-Path $TestDrive 'does-not-exist'

        $result = Write-MatrixSettingLogHC -Matrix $matrix -Html $html -LogFolder $missing

        $result | Should -BeNullOrEmpty
    }
}