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

Describe 'Initialize-HtmlStructureHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Html.ps1"
    }

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

    It 'exposes Templates.SettingsHeader containing a Settings <th>' {
        $struct.Templates.SettingsHeader | Should -Match '<th class="matrixHeader" colspan="8">Settings</th>'
    }

    It 'exposes Templates.LegendTable with Error/Warning/Information cells' {
        $struct.Templates.LegendTable | Should -Match 'probTypeError'
        $struct.Templates.LegendTable | Should -Match 'probTypeWarning'
        $struct.Templates.LegendTable | Should -Match 'probTypeInfo'
    }
}

Describe 'Get-HtmlClassProbTypeHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Html.ps1"
    }

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

Describe 'Build-MatrixEmailHtmlHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Html.ps1"
        . "$moduleRoot\Private\Utils.ps1"  # Get-StringOrDefaultHC

        # FileResult shape that matches what Build-MatrixEmailHtmlHC actually
        # reads. Production paths (per Html.ps1 lines 576-621):
        #   .Item.Name            → filename in title
        #   .Item.FullName        → href for the title link
        #   .ExcelInfo.LastModifiedBy
        #   .ExcelInfo.Modified   (datetime)
        #   .Check                → global file-level checks (array)
        #   .Sheets.FormData.Check
        #   .Sheets.Permissions.Check
        #   .Matrices             → array of matrix rows with .ID, .Check, .Setting.Formatted.ComputerName
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
    }

    BeforeEach {
        $script:html = Initialize-HtmlStructureHC

        # Build-MatrixEmailHtmlHC calls New-SettingsOverviewHtmlHC when
        # $Matrices is non-empty; mock to keep this test focused on the
        # orchestration around the file headers, not the matrices subsection.
        Mock New-SettingsOverviewHtmlHC { return '<!-- mocked overview -->' }
    }

    Context 'basic file rendering' {
        It 'renders the filename in the title link text' {
            $files = @( New-FileResult -Name 'Q3-Permissions.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'Q3-Permissions\.xlsx'
        }

        It 'uses Item.FullName as the title link href' {
            $files = @( New-FileResult -FullName 'C:\share\budget.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match '<a href="C:\\share\\budget\.xlsx">'
        }

        It 'renders one <table class="matrixTable"> per file' {
            $files = @(
                New-FileResult -Name 'one.xlsx'
                New-FileResult -Name 'two.xlsx'
                New-FileResult -Name 'three.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            ([regex]::Matches($out, '<table class="matrixTable">')).Count | Should -Be 3
        }
    }

    Context 'ExcelInfo handling' {
        It 'renders LastModifiedBy in the file info row' {
            $files = @( New-FileResult -LastModifiedBy 'alice@example.com' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'alice@example\.com'
        }

        It 'falls back to "Unknown" when LastModifiedBy is empty' {
            $files = @( New-FileResult -LastModifiedBy '' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'Last change: Unknown'
        }

        It 'formats Modified as dd/MM/yyyy HH:mm:ss' {
            $files = @(
                New-FileResult -Modified (Get-Date '2024-03-22 14:05:09')
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match '22/03/2024 14:05:09'
        }

        It 'falls back to "Unknown" when Modified is not a datetime' {
            $fr = New-FileResult
            # Overwrite Modified with a non-datetime value
            $fr.ExcelInfo.Modified = 'not-a-date'

            $out = Build-MatrixEmailHtmlHC -FileResults @($fr) -Html $html

            $out | Should -Match 'Last change: User @ Unknown'
        }

        It 'HTML-encodes the filename' {
            $files = @( New-FileResult -Name 'a&b<c>.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'a&amp;b&lt;c&gt;\.xlsx'
            $out | Should -Not -Match '<c>'
        }
    }

    Context 'matrices section' {
        It 'calls New-SettingsOverviewHtmlHC when the file has matrices' {
            $files = @(
                New-FileResult -Matrices @(
                    [pscustomobject]@{
                        ID      = 1
                        Check   = @()
                        Setting = [pscustomobject]@{
                            Formatted = [pscustomobject]@{ ComputerName = 'SRV01' }
                        }
                    }
                )
            )

            Build-MatrixEmailHtmlHC -FileResults $files -Html $html | Out-Null

            Should -Invoke New-SettingsOverviewHtmlHC -Times 1
        }

        It 'does not call New-SettingsOverviewHtmlHC when Matrices is empty' {
            $files = @( New-FileResult -Matrices @() )

            Build-MatrixEmailHtmlHC -FileResults $files -Html $html | Out-Null

            Should -Invoke New-SettingsOverviewHtmlHC -Times 0
        }

        It 'includes settings details only for matrices that have Check entries' {
            $matrixWithIssues = [pscustomobject]@{
                ID      = 1
                Check   = @(
                    [pscustomobject]@{ Type = 'Warning'; Name = 'thing'; Description = 'something off' }
                )
                Setting = [pscustomobject]@{
                    Formatted = [pscustomobject]@{ ComputerName = 'SRV01' }
                }
            }
            $matrixClean = [pscustomobject]@{
                ID      = 2
                Check   = @()
                Setting = [pscustomobject]@{
                    Formatted = [pscustomobject]@{ ComputerName = 'SRV02' }
                }
            }

            $files = @( New-FileResult -Matrices @($matrixWithIssues, $matrixClean) )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'Settings sheet details \(ID: 1\) - SRV01'
            $out | Should -Not -Match 'Settings sheet details \(ID: 2\) - SRV02'
        }
    }
}

Describe 'Write-MatrixSettingLogHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Html.ps1"
    }

    BeforeEach {
        $script:html = Initialize-HtmlStructureHC
        $script:logFolder = Join-Path $TestDrive 'logs'
        New-Item -ItemType Directory -Path $logFolder -Force | Out-Null

        $script:matrix = [pscustomobject]@{
            ID    = 42
            Check = @(
                [pscustomobject]@{
                    Type        = 'Warning'
                    Name        = 'TestCheck'
                    Description = 'A test description'
                }
            )
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
        $content | Should -Match '<h1>Settings Log - ID 42</h1>'
    }

    It 'includes the legend table in the output' {
        Write-MatrixSettingLogHC -Matrix $matrix -Html $html -LogFolder $logFolder

        $path = Join-Path $logFolder 'ID 42 - Settings.html'
        $content = Get-Content -LiteralPath $path -Raw

        $content | Should -Match 'legendTable'
    }

    It 'returns null when the log folder does not exist' {
        $missing = Join-Path $TestDrive 'does-not-exist'

        $result = Write-MatrixSettingLogHC -Matrix $matrix -Html $html -LogFolder $missing

        $result | Should -BeNullOrEmpty
    }
}
