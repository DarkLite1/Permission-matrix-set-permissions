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
