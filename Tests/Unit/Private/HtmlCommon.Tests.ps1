#Requires -Version 7
#requires -Modules Pester

BeforeAll {
    # Load the module code to test
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
    ForEach-Object { . $_.FullName }
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

