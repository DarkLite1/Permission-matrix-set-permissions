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
        $struct | Should-HaveType ([hashtable])
    }

    It 'exposes Style as a non-empty <style> block' {
        $struct.Style | Should-BeTruthy
        $struct.Style | Should-MatchString '<style type="text/css">'
        $struct.Style | Should-MatchString '</style>'
    }

    It 'exposes TroubleshootingStyle as a separate <style> block' {
        $struct.TroubleshootingStyle | Should-BeTruthy
        $struct.TroubleshootingStyle | Should-MatchString '<style type="text/css">'
    }

    It 'embeds the themed page background colour in the Style block' {
        $struct.Style | Should-MatchString 'background-color: #e5e7eb;'
    }
}

Describe 'Get-FileCheckTallyHC' {
    BeforeAll {
        function New-TallyFileResult {
            param(
                [object[]]$Check = @(),
                [object[]]$FormDataCheck = @(),
                [object[]]$PermissionsCheck = @(),
                [object[]]$MatrixCheck = @()
            )
            return [pscustomobject]@{
                Check    = $Check
                Sheets   = [pscustomobject]@{
                    FormData    = [pscustomobject]@{ Check = $FormDataCheck }
                    Permissions = [pscustomobject]@{ Check = $PermissionsCheck }
                }
                Matrices = @([pscustomobject]@{ ID = 1; Check = $MatrixCheck })
            }
        }
    }

    It 'counts errors and warnings on the file itself' {
        $tally = Get-FileCheckTallyHC -FileResult (New-TallyFileResult -Check @(
                [pscustomobject]@{ Type = 'FatalError'; Name = 'e' }
                [pscustomobject]@{ Type = 'Warning'; Name = 'w' }
                [pscustomobject]@{ Type = 'Warning'; Name = 'w2' }
            ))
        $tally.Errors | Should-Be 1
        $tally.Warnings | Should-Be 2
    }

    It 'includes checks from the FormData and Permissions sheets' {
        $tally = Get-FileCheckTallyHC -FileResult (New-TallyFileResult `
                -FormDataCheck @([pscustomobject]@{ Type = 'FatalError'; Name = 'e' }) `
                -PermissionsCheck @([pscustomobject]@{ Type = 'Warning'; Name = 'w' }))
        $tally.Errors | Should-Be 1
        $tally.Warnings | Should-Be 1
    }

    It 'includes checks from the matrices' {
        $tally = Get-FileCheckTallyHC -FileResult (New-TallyFileResult -MatrixCheck @(
                [pscustomobject]@{ Type = 'Warning'; Name = 'w' }
            ))
        $tally.Warnings | Should-Be 1
    }

    It 'ignores Information checks' {
        $tally = Get-FileCheckTallyHC -FileResult (New-TallyFileResult -MatrixCheck @(
                [pscustomobject]@{ Type = 'Information'; Name = 'i1' }
                [pscustomobject]@{ Type = 'Information'; Name = 'i2' }
            ))
        $tally.Errors | Should-Be 0
        $tally.Warnings | Should-Be 0
    }

    It 'returns zeroes for a file result with no checks' {
        $tally = Get-FileCheckTallyHC -FileResult (New-TallyFileResult)
        $tally.Errors | Should-Be 0
        $tally.Warnings | Should-Be 0
    }

    It 'returns zeroes for a null file result' {
        $tally = Get-FileCheckTallyHC -FileResult $null
        $tally.Errors | Should-Be 0
        $tally.Warnings | Should-Be 0
    }
}

Describe 'Get-MatrixFileNameHC' {
    It 'returns the name from the Item property' {
        $fileResult = [pscustomobject]@{
            Item = [pscustomobject]@{ Name = 'DNK HCPT.xlsx' }
        }
        Get-MatrixFileNameHC -FileResult $fileResult | Should-Be 'DNK HCPT.xlsx'
    }

    It 'falls back to the File property when Item is absent' {
        # Shape produced by the runspace catch block in Invoke-PermissionMatrixBeginHC
        $fileResult = [pscustomobject]@{
            File = [pscustomobject]@{ Name = 'Crashed.xlsx' }
        }
        Get-MatrixFileNameHC -FileResult $fileResult | Should-Be 'Crashed.xlsx'
    }

    It 'prefers Item over File when both are present' {
        $fileResult = [pscustomobject]@{
            Item = [pscustomobject]@{ Name = 'FromItem.xlsx' }
            File = [pscustomobject]@{ Name = 'FromFile.xlsx' }
        }
        Get-MatrixFileNameHC -FileResult $fileResult | Should-Be 'FromItem.xlsx'
    }

    It 'returns the default when neither property yields a name' {
        Get-MatrixFileNameHC -FileResult ([pscustomobject]@{}) -Default '(unknown)' |
        Should-Be '(unknown)'
    }

    It 'returns the default for a null file result' {
        Get-MatrixFileNameHC -FileResult $null -Default '(unknown)' | Should-Be '(unknown)'
    }

    It 'returns an empty string by default' {
        Get-MatrixFileNameHC -FileResult ([pscustomobject]@{}) | Should-Be ''
    }
}

Describe 'Format-IssueCountLabelHC' {
    It 'returns "Success" when there are no errors or warnings' {
        Format-IssueCountLabelHC -Errors 0 -Warnings 0 | Should-Be 'Success'
    }

    It 'singularises a single error' {
        Format-IssueCountLabelHC -Errors 1 -Warnings 0 | Should-Be '1 Error'
    }

    It 'pluralises multiple errors' {
        Format-IssueCountLabelHC -Errors 3 -Warnings 0 | Should-Be '3 Errors'
    }

    It 'singularises a single warning' {
        Format-IssueCountLabelHC -Errors 0 -Warnings 1 | Should-Be '1 Warning'
    }

    It 'pluralises multiple warnings' {
        Format-IssueCountLabelHC -Errors 0 -Warnings 2 | Should-Be '2 Warnings'
    }

    It 'joins errors and warnings with a comma' {
        Format-IssueCountLabelHC -Errors 2 -Warnings 1 | Should-Be '2 Errors, 1 Warning'
    }
}

Describe 'Format-LastChangeHC' {
    It 'combines user and date when both are known' {
        $dt = Get-Date '2026-05-19 13:30:00'
        Format-LastChangeHC -LastModifiedBy 'Brecht' -Modified $dt | Should-Be 'Last change: Brecht &middot; 19/05/2026 13:30'
    }

    It 'shows only the user when the date is missing' {
        Format-LastChangeHC -LastModifiedBy 'Brecht' -Modified $null | Should-Be 'Last change: Brecht'
    }

    It 'shows only the date when the user is missing' {
        $dt = Get-Date '2026-05-19 13:30:00'
        Format-LastChangeHC -LastModifiedBy '' -Modified $dt | Should-Be 'Last change: 19/05/2026 13:30'
    }

    It 'treats the literal "Unknown" username as missing' {
        Format-LastChangeHC -LastModifiedBy 'Unknown' -Modified $null | Should-Be 'No modification metadata available'
    }

    It 'returns a placeholder when neither value is known' {
        Format-LastChangeHC -LastModifiedBy '' -Modified $null | Should-Be 'No modification metadata available'
    }

    It 'treats a non-datetime Modified value as missing' {
        Format-LastChangeHC -LastModifiedBy 'Brecht' -Modified 'not-a-date' | Should-Be 'Last change: Brecht'
    }

    It 'treats DateTime.MinValue as missing' {
        Format-LastChangeHC -LastModifiedBy 'Brecht' -Modified ([datetime]::MinValue) | Should-Be 'Last change: Brecht'
    }

    It 'HTML-encodes the username component' {
        Format-LastChangeHC -LastModifiedBy 'A&B' -Modified $null | Should-Be 'Last change: A&amp;B'
    }

    It 'uses HH:mm (not seconds) for the time component' {
        $dt = Get-Date '2026-05-19 13:30:45'
        Format-LastChangeHC -LastModifiedBy '' -Modified $dt | Should-MatchString '13:30$'
    }
}

Describe 'ConvertTo-FileUrlHC' {
    It 'returns empty string for null or whitespace input' {
        ConvertTo-FileUrlHC -Path $null | Should-Be ''
        ConvertTo-FileUrlHC -Path '   ' | Should-Be ''
    }

    It 'prefixes file:// and converts backslashes to forward slashes' {
        ConvertTo-FileUrlHC -Path 'C:\share\budget.xlsx' | Should-Be 'file://C:/share/budget.xlsx'
    }

    It 'percent-encodes spaces' {
        ConvertTo-FileUrlHC -Path 'C:\my files\a b.xlsx' | Should-Be 'file://C:/my%20files/a%20b.xlsx'
    }

    It 'converts UNC paths' {
        ConvertTo-FileUrlHC -Path '\\srv01\teamA\m.xlsx' | Should-Be 'file:////srv01/teamA/m.xlsx'
    }
}

Describe 'Get-CheckThemeHC' {
    It 'returns the error theme for FatalError' {
        $t = Get-CheckThemeHC 'FatalError'
        $t.Label | Should-Be 'ERROR'
        $t.Symbol | Should-Be '✖'
        $t.Accent | Should-Be '#dc2626'
    }

    It 'returns the warning theme for Warning' {
        $t = Get-CheckThemeHC 'Warning'
        $t.Label | Should-Be 'WARNING'
        $t.Symbol | Should-Be '⚠'
        $t.Accent | Should-Be '#d97706'
    }

    It 'returns the info theme for any other value' {
        $t = Get-CheckThemeHC 'Information'
        $t.Label | Should-Be 'INFO'
        $t.Symbol | Should-Be 'ℹ'
        $t.Accent | Should-Be '#6b7280'
    }
}

Describe 'New-PillHtmlHC' {
    It 'returns empty string for blank text' {
        New-PillHtmlHC -Text '' -Bg '#000000' | Should-Be ''
        New-PillHtmlHC -Text '   ' -Bg '#000000' | Should-Be ''
    }

    It 'renders a span with the supplied text and background colour' {
        $pill = New-PillHtmlHC -Text 'Error' -Bg '#dc2626'
        $pill | Should-MatchString '<span'
        $pill | Should-MatchString 'background-color:#dc2626;'
        $pill | Should-MatchString '>Error</span>'
    }

    It 'defaults the text colour to white' {
        $pill = New-PillHtmlHC -Text 'OK' -Bg '#16a34a'
        $pill | Should-MatchString 'color:#ffffff;'
    }

    It 'honours a custom text colour' {
        $pill = New-PillHtmlHC -Text 'OK' -Bg '#16a34a' -Color '#000000'
        $pill | Should-MatchString 'color:#000000;'
    }
}

Describe 'Build-ErrorWarningTableHC' {
    It 'returns empty string when there are no issues' {
        $counter = [pscustomobject]@{ TotalErrors = 0; TotalWarnings = 0 }
        Build-ErrorWarningTableHC -CounterData $counter | Should-Be ''
    }

    It 'renders an error pill when there are errors' {
        $counter = [pscustomobject]@{ TotalErrors = 2; TotalWarnings = 0 }
        $html = Build-ErrorWarningTableHC -CounterData $counter
        $html | Should-MatchString 'Detected issues'
        $html | Should-MatchString '2 Errors'
    }

    It 'renders a warning pill when there are warnings' {
        $counter = [pscustomobject]@{ TotalErrors = 0; TotalWarnings = 1 }
        $html = Build-ErrorWarningTableHC -CounterData $counter
        $html | Should-MatchString '1 Warning'
    }

    It 'renders both pills when there are errors and warnings' {
        $counter = [pscustomobject]@{ TotalErrors = 1; TotalWarnings = 3 }
        $html = Build-ErrorWarningTableHC -CounterData $counter
        $html | Should-MatchString '1 Error'
        $html | Should-MatchString '3 Warnings'
    }
}

Describe 'Build-FileLevelCheckRowHC' {
    It 'renders the check name, description and sheet label' {
        $check = [pscustomobject]@{ Type = 'Warning'; Name = 'CheckName'; Description = 'CheckDesc' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'Excel File'
        $html | Should-MatchString 'CheckName'
        $html | Should-MatchString 'CheckDesc'
        $html | Should-MatchString 'Excel File'
    }

    It 'uses the error accent colour for FatalError checks' {
        $check = [pscustomobject]@{ Type = 'FatalError'; Name = 'n'; Description = 'd' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'Excel File'
        $html | Should-MatchString '#dc2626'
        $html | Should-MatchString 'ERROR'
    }

    It 'includes the 16px inset wrapper by default' {
        $check = [pscustomobject]@{ Type = 'Warning'; Name = 'n'; Description = 'd' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'X'
        $html | Should-MatchString 'padding:0 16px 8px 16px;'
    }

    It 'omits the inset wrapper when -IncludeWrapper is $false' {
        $check = [pscustomobject]@{ Type = 'Warning'; Name = 'n'; Description = 'd' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'X' -IncludeWrapper $false
        $html | Should-MatchString 'padding:0 0 8px 0;'
    }

    It 'falls back to placeholder text for a missing name' {
        $check = [pscustomobject]@{ Type = 'Warning'; Name = ''; Description = 'd' }
        $html = Build-FileLevelCheckRowHC -Check $check -SheetLabel 'X'
        $html | Should-MatchString 'Unnamed check'
    }
}