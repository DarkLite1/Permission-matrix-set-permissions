#Requires -Version 7
#requires -Modules Pester

BeforeAll {
    # Load the module code to test
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
    ForEach-Object { . $_.FullName }
}

Describe 'Build-SystemErrorsBlockHC' {
    It 'returns empty string for a null or empty list' {
        Build-SystemErrorsBlockHC -SystemErrors @() | Should-Be ''
        Build-SystemErrorsBlockHC -SystemErrors $null | Should-Be ''
    }

    It 'ignores items that are neither FatalError nor Warning' {
        $items = @(
            [pscustomobject]@{ Type = 'Information'; Name = 'note'; Message = 'fyi' }
        )
        Build-SystemErrorsBlockHC -SystemErrors $items | Should-Be ''
    }

    It 'renders a System Error card for a FatalError item' {
        $items = @(
            [pscustomobject]@{ Type = 'FatalError'; Name = 'Boom'; Message = 'it broke'; Category = 'Matrix' }
        )
        $html = Build-SystemErrorsBlockHC -SystemErrors $items
        $html | Should-MatchString 'System Error'
        $html | Should-MatchString 'Boom'
        $html | Should-MatchString 'it broke'
        $html | Should-MatchString '1 Error'
    }

    It 'renders a System Warning card for a Warning item' {
        $items = @(
            [pscustomobject]@{ Type = 'Warning'; Name = 'Careful'; Message = 'heads up'; Category = '' }
        )
        $html = Build-SystemErrorsBlockHC -SystemErrors $items
        $html | Should-MatchString 'System Warning'
        $html | Should-MatchString '1 Warning'
    }

    It 'HTML-encodes the item name' {
        $items = @(
            [pscustomobject]@{ Type = 'FatalError'; Name = 'a&b'; Message = 'm'; Category = '' }
        )
        $html = Build-SystemErrorsBlockHC -SystemErrors $items
        $html | Should-MatchString 'a&amp;b'
    }
}

Describe 'Build-MailTopLinksBlockHC' {
    BeforeEach {
        $script:html = Initialize-HtmlStructureHC
    }

    It 'returns empty string when no browser or export links are available' {
        Build-MailTopLinksBlockHC | Should-Be ''
    }

    It 'renders a browser-view link when the mail body log path is known' {
        $out = Build-MailTopLinksBlockHC -BrowserViewFilePath 'C:\logs\Mail - Run.html'

        $out | Should-MatchString 'If this mail is not visible'
        $out | Should-MatchString 'click here to view it in the browser'
        $out | Should-MatchString "href='file://C:/logs/Mail%20-%20Run\.html'"
        $out | Should-MatchString 'title="C:\\logs\\Mail - Run\.html"'
    }

    It 'renders export file links from the ordered dictionary returned by Export-FilesHC' {
        $exportedFiles = [ordered]@{
            Permissions  = 'C:\reports\Permissions.xlsx'
            FormData     = 'C:\ServiceNow\FormData.xlsx'
            OverviewHtml = 'C:\reports\Overview.html'
        }

        $out = Build-MailTopLinksBlockHC -ExportedFiles $exportedFiles

        $out | Should-MatchString 'Export files:'
        $out | Should-MatchString 'Permissions Excel'
        $out | Should-MatchString 'ServiceNow FormData Excel'
        $out | Should-MatchString 'Overview HTML'
        $out | Should-MatchString "href='file://C:/reports/Permissions\.xlsx'"
        $out | Should-MatchString "href='file://C:/ServiceNow/FormData\.xlsx'"
        $out | Should-MatchString "href='file://C:/reports/Overview\.html'"
    }

    It 'omits export links whose configured export path was not written' {
        $out = Build-MailTopLinksBlockHC -ExportedFiles ([pscustomobject]@{
                Permissions  = 'C:\reports\Permissions.xlsx'
                FormData     = $null
                OverviewHtml = ''
            })

        $out | Should-MatchString 'Permissions Excel'
        $out | Should-NotMatchString 'ServiceNow FormData Excel'
        $out | Should-NotMatchString 'Overview HTML'
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
                [string]$ReportFilePath = '',
                [timespan]$Duration
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
                JobTime     = [pscustomobject]@{
                    Duration = if ($PSBoundParameters.ContainsKey('Duration')) { $Duration } else { $null }
                }
                FileContext = [pscustomobject]@{ ReportFilePath = $ReportFilePath }
            }
        }
    }

    It 'renders the computer name and path' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem -ComputerName 'SRV99' -Path 'E:\foo')
        $html | Should-MatchString 'SRV99'
        $html | Should-MatchString 'E:\\foo'
    }

    It 'shows an Error pill when the row has a FatalError check' {
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'FatalError' })
        $html = Build-SettingsRowHC -MatrixItem $item
        $html | Should-MatchString '>Error</span>'
        $html | Should-MatchString '#dc2626'
    }

    It 'shows a Warning pill when the row has a Warning check' {
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'Warning' })
        $html = Build-SettingsRowHC -MatrixItem $item
        $html | Should-MatchString '>Warning</span>'
        $html | Should-MatchString '#d97706'
    }

    It 'shows an Incorrect pill when the row has an Incorrect check' {
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'Incorrect' })
        $html = Build-SettingsRowHC -MatrixItem $item
        $html | Should-MatchString '>Incorrect</span>'
        $html | Should-MatchString '#ea580c'
    }

    It 'shows a Fixed pill on a green row when the run corrected the drift' {
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'Fixed' })
        $html = Build-SettingsRowHC -MatrixItem $item
        $html | Should-MatchString '>Fixed</span>'
        $html | Should-MatchString '#16a34a'
    }

    It 'lets Incorrect outrank Warning on the same row' {
        $item = New-MatrixItem -Check @(
            [pscustomobject]@{ Type = 'Warning' }
            [pscustomobject]@{ Type = 'Incorrect' }
        )
        $html = Build-SettingsRowHC -MatrixItem $item
        $html | Should-MatchString '>Incorrect</span>'
        $html | Should-NotMatchString '>Warning</span>'
    }

    It 'uses the success accent and no status pill for a clean row' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem)
        $html | Should-MatchString '#16a34a'
        $html | Should-NotMatchString 'rr-srow-status'
    }

    It 'flags an info-only matrix with a blue info glyph next to the name' {
        # A matrix-level Information check earns no pill and never reaches the
        # file-level tally, so without the glyph the row looks perfectly clean.
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'Information' })
        $html = Build-SettingsRowHC -MatrixItem $item

        $html | Should-MatchString '&#8505;'
        $html | Should-MatchString '#2563eb'
        # Still a clean green row: the glyph is a hint, not a status.
        $html | Should-MatchString '#16a34a'
        $html | Should-NotMatchString 'rr-srow-status'
    }

    It 'shows the info glyph alongside an existing Warning pill' {
        # BNL-MTX-STAFF-HR case: one Warning AND one Information on the same
        # matrix. The pill shows the warning, the glyph reveals the rest.
        $item = New-MatrixItem -Check @(
            [pscustomobject]@{ Type = 'Warning' }
            [pscustomobject]@{ Type = 'Information' }
        )
        $html = Build-SettingsRowHC -MatrixItem $item

        $html | Should-MatchString '>Warning</span>'
        $html | Should-MatchString '&#8505;'
    }

    It 'shows no info glyph when the row has only errors or warnings' {
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'FatalError' })
        $html = Build-SettingsRowHC -MatrixItem $item
        $html | Should-NotMatchString '&#8505;'
    }

    It 'shows no info glyph on a row without checks' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem)
        $html | Should-NotMatchString '&#8505;'
    }

    It 'pluralizes the info glyph tooltip and counts every notice' {
        $item = New-MatrixItem -Check @(
            [pscustomobject]@{ Type = 'Information' }
            [pscustomobject]@{ Type = 'Information' }
        )
        $html = Build-SettingsRowHC -MatrixItem $item
        $html | Should-MatchString 'title="2 information notices on this matrix'
    }

    It 'keeps the info glyph inside the name line at the line''s own metrics' {
        # The identifier cell drives the row height; a larger inline run would
        # change Word's line box and skew the pill/meta centring.
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'Information' })
        $html = Build-SettingsRowHC -MatrixItem $item

        $html | Should-MatchString "font-size:13px; font-weight:400; line-height:15px; mso-line-height-rule:exactly;'>&#8505;</span></div>"
    }

    It 'shows N/A for a missing duration' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem)
        $html | Should-MatchString 'N/A'
    }

    It 'centres the duration in its own column so N/A lines up with timestamps' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem)
        $html | Should-MatchString "align='center'"
        $html | Should-MatchString 'text-align:center;.*>N/A<'
    }

    It 'keeps the Action cell identical whether or not a duration is present' {
        # The bug: Action and Duration shared one right-aligned cell, so a short
        # 'N/A' shifted the Action label right. Separate fixed-width cells mean
        # the Action markup is byte-for-byte the same in both cases.
        $pattern = "(?s)(<td[^>]*rr-srow-meta'[^>]*>Fix</td>)"

        $withDuration = Build-SettingsRowHC -MatrixItem (
            New-MatrixItem -Action 'Fix' -Duration ([timespan]::FromSeconds(15)))
        $withoutDuration = Build-SettingsRowHC -MatrixItem (New-MatrixItem -Action 'Fix')

        $withDuration | Should-MatchString '>00:00:15<'
        $withoutDuration | Should-MatchString '>N/A<'

        ([regex]::Match($withDuration, $pattern).Value) |
        Should-Be ([regex]::Match($withoutDuration, $pattern).Value)
    }

    It 'links the row to the report file path when present' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem -ReportFilePath 'C:\logs\r.html')
        $html | Should-MatchString "href='C:\\logs\\r\.html'"
    }

    It 'marks a clean row as Skipped (grey) when the file has a fatal error' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem) -FileHasError $true
        $html | Should-MatchString '#6b7280'
        $html | Should-MatchString '>Skipped</span>'
        $html | Should-NotMatchString '#16a34a'
    }

    It 'keeps a row with its own error red even when the file has a fatal error' {
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'FatalError' })
        $html = Build-SettingsRowHC -MatrixItem $item -FileHasError $true
        $html | Should-MatchString '>Error</span>'
        $html | Should-NotMatchString '>Skipped</span>'
    }

    It 'shows Skipped rather than Fixed when the file has a fatal error' {
        # The accent bar is grey here, so a green 'Fixed' pill beside it would
        # claim a success this row never had.
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'Fixed' })
        $html = Build-SettingsRowHC -MatrixItem $item -FileHasError $true
        $html | Should-MatchString '>Skipped</span>'
        $html | Should-NotMatchString '>Fixed</span>'
        $html | Should-MatchString '#6b7280'
    }

    It 'middle-aligns the Outlook row chrome with exact line heights' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem)
        $html | Should-MatchString 'table-layout:fixed;'
        $html | Should-MatchString "valign='middle' width='20' style='vertical-align:middle; padding:4px 0 4px 12px;"
        $html | Should-MatchString "valign='middle' class='rr-srow-ident' style='vertical-align:middle; padding:4px 8px;'"
        $html | Should-MatchString '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="rr-srow" style="border-collapse:separate; width:100%; max-width:100%; margin:0 0 4px 0; table-layout:fixed;'
        $html | Should-MatchString "valign='middle' align='right' class='rr-srow-meta' width='44' style='vertical-align:middle; padding:4px 0 4px 10px;"
        $html | Should-MatchString "valign='middle' align='center' nowrap='nowrap' class='rr-srow-meta rr-srow-dur' width='68' style='vertical-align:middle; padding:4px 10px 4px 8px;"
        $html | Should-NotMatchString '<td height="6" style="font-size:0; line-height:0;">&#160;</td>'
    }

    It 'keeps path wrapping styles valid inside single-quoted attributes' {
        $html = Build-SettingsRowHC -MatrixItem (New-MatrixItem -Path 'E:\very\long\path')

        $html | Should-MatchString 'font-family:Consolas, Menlo, monospace; font-size:11px;'
        $html | Should-MatchString 'white-space:normal; overflow-wrap:anywhere; word-break:break-all;'
    }

    It 'splits the pill into a nested-table Outlook cell and a normal browser cell' {
        $item = New-MatrixItem -Check @([pscustomobject]@{ Type = 'Warning' })
        $html = Build-SettingsRowHC -MatrixItem $item

        # Outlook cell: gated by [if mso], wraps the VML in a 3-row nested table
        # whose top(8px)/bottom(4px) spacers make it the tallest cell (drives the
        # row height) and nudge the pill down to centre; line-height:26px avoids clipping.
        $html | Should-MatchString '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="rr-srow" style="border-collapse:separate; width:100%; max-width:100%; margin:0 0 4px 0; table-layout:fixed'
        # Browser cell: gated by [if !mso], keeps a normal font-size (line-height
        # 16px) so the CSS span stays perfectly centred as it was before.
        $html | Should-MatchString '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="rr-srow" style="border-collapse:separate; width:100%; max-width:100%; margin:0 0 4px 0; table-layout:fixed;'
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
                [object[]]$Matrices = @(),
                [string]$ReportFilePath = '',
                [string]$LogMatrixFilePath = ''
            )

            $obj = [pscustomobject]@{
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

            # Added conditionally so the default fixtures keep the old
            # object shape, exercising the absent-property code path
            if ($ReportFilePath) {
                $obj | Add-Member -NotePropertyName ReportFilePath `
                    -NotePropertyValue $ReportFilePath
            }
            if ($LogMatrixFilePath) {
                $obj | Add-Member -NotePropertyName LogMatrixFilePath `
                    -NotePropertyValue $LogMatrixFilePath
            }

            return $obj
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

            $out | Should-MatchString 'Q3-Permissions\.xlsx'
        }

        It 'uses a file:// URL derived from Item.FullName as the title link href' {
            $files = @( New-FileResult -FullName 'C:\share\budget.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            # ConvertTo-FileUrlHC turns the Windows path into a file:// URL with
            # forward slashes; the anchor also carries a title tooltip.
            $out | Should-MatchString '<a href="file://C:/share/budget\.xlsx"'
        }

        It 'puts the raw Windows path in the title tooltip of the header link' {
            $files = @( New-FileResult -FullName 'C:\share\budget.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'title="C:\\share\\budget\.xlsx"'
        }

        It 'renders one file card table per file result' {
            $files = @(
                New-FileResult -Name 'one.xlsx'
                New-FileResult -Name 'two.xlsx'
                New-FileResult -Name 'three.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            ([regex]::Matches($out, 'width="100%" bgcolor="#ffffff"')).Count | Should-Be 3
        }
    }

    Context 'footer links' {
        It 'renders an execution report link when ReportFilePath is set' {
            $files = @(
                New-FileResult `
                    -ReportFilePath 'C:\logs\00 - Execution Report.html'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'Open execution report'
            $out | Should-MatchString "href='file://C:/logs/00%20-%20Execution%20Report\.html'"
            # No fallback link when a log artifact exists
            $out | Should-NotMatchString 'Open matrix file'
        }

        It 'renders a matrix log copy link when LogMatrixFilePath is set' {
            $files = @(
                New-FileResult -LogMatrixFilePath 'C:\logs\A.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'Open matrix log copy'
            $out | Should-MatchString "href='file://C:/logs/A\.xlsx'"
            $out | Should-NotMatchString 'Open matrix file'
            $out | Should-NotMatchString 'Open execution report'
        }

        It 'puts the raw Windows path in the tooltip of the matrix log copy link' {
            $files = @(
                New-FileResult -LogMatrixFilePath 'C:\logs\A.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'title="C:\\logs\\A\.xlsx"'
        }

        It 'renders both links separated by a middot when both artifacts exist' {
            $files = @(
                New-FileResult `
                    -ReportFilePath 'C:\logs\report.html' `
                    -LogMatrixFilePath 'C:\logs\A.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'Open execution report'
            $out | Should-MatchString 'Open matrix log copy'
            # The browser variant joins the two anchors with a padded middot span
            $out | Should-MatchString 'Open execution report &rarr;</a><span[^>]*>&middot;</span><a'
            # The Outlook (Word) variant ignores span padding, so the middot is
            # spaced with non-breaking spaces instead
            $out | Should-MatchString 'Open execution report &rarr;</a><span[^>]*>&nbsp;&nbsp;&middot;&nbsp;&nbsp;</span><a'
        }

        It 'uses separate Outlook and browser footer spacing' {
            $files = @(
                New-FileResult `
                    -ReportFilePath 'C:\logs\report.html' `
                    -LogMatrixFilePath 'C:\logs\A.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString '<!--\[if mso\]>'
            $out | Should-MatchString "valign='middle' style='padding:6px 16px 6px 16px; text-align:center; font-size:12px; line-height:16px; mso-line-height-rule:exactly; color:#6b7280;"
            $out | Should-MatchString "<p style='margin:0; mso-line-height-rule:exactly; line-height:16px;'>"
            $out | Should-MatchString '<!--\[if !mso\]><!-->'
            $out | Should-MatchString 'padding:4px 16px 12px 16px; text-align:center; font-size:12px; line-height:16px; color:#6b7280;'
            $out | Should-MatchString '<td height="16" style="font-size:0; line-height:16px; mso-line-height-rule:exactly;">&#160;</td>'
        }

        It 'falls back to the source matrix file link when no log artifacts exist' {
            $files = @( New-FileResult -FullName 'C:\share\A.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'Open matrix file'
            # The fallback footer anchor uses single-quoted href, unlike
            # the double-quoted header link to the same file URL
            $out | Should-MatchString "href='file://C:/share/A\.xlsx'"
            $out | Should-NotMatchString 'Open execution report'
            $out | Should-NotMatchString 'Open matrix log copy'
        }
    }

    Context 'ExcelInfo handling' {
        It 'renders LastModifiedBy in the file info row' {
            $files = @( New-FileResult -LastModifiedBy 'alice@example.com' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'alice@example\.com'
        }

        It 'omits the user but keeps the date when LastModifiedBy is empty' {
            $files = @( New-FileResult -LastModifiedBy '' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            # Format-LastChangeHC drops the user and renders date-only.
            $out | Should-MatchString 'Last change: 15/01/2024 09:30'
            $out | Should-NotMatchString 'Last change: Unknown'
        }

        It 'formats Modified as dd/MM/yyyy HH:mm' {
            $files = @(
                New-FileResult -Modified (Get-Date '2024-03-22 14:05:09')
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            # The layout uses minute precision, not seconds.
            $out | Should-MatchString '22/03/2024 14:05'
            $out | Should-NotMatchString '22/03/2024 14:05:09'
        }

        It 'shows only the user when Modified is not a datetime' {
            $fr = New-FileResult
            # Overwrite Modified with a non-datetime value
            $fr.ExcelInfo.Modified = 'not-a-date'

            $out = Build-MatrixEmailHtmlHC -FileResults @($fr) -Html $html

            $out | Should-MatchString 'Last change: User'
            $out | Should-NotMatchString 'Last change: User &middot;'
        }

        It 'HTML-encodes the filename' {
            $files = @( New-FileResult -Name 'a&b<c>.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'a&amp;b&lt;c&gt;\.xlsx'
            $out | Should-NotMatchString '<c>'
        }
    }

    Context 'header status' {
        It 'shows a success header for a file with no issues' {
            $out = Build-MatrixEmailHtmlHC -FileResults @( New-FileResult ) -Html $html
            $out | Should-MatchString '✓'
            $out | Should-MatchString 'Success'
        }

        It 'reserves enough header space for the status label in Outlook' {
            $out = Build-MatrixEmailHtmlHC -FileResults @( New-FileResult ) -Html $html

            $out | Should-MatchString "valign='middle' align='right' width='112' style='padding:14px 12px 14px 6px; white-space:nowrap; width:112px;'"
        }

        It 'shows a warning header when a matrix row has a Warning' {
            $files = @(
                New-FileResult -Matrices @(
                    New-MatrixRow -Check @([pscustomobject]@{ Type = 'Warning'; Name = 'w'; Description = 'd' })
                )
            )
            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html
            $out | Should-MatchString '⚠'
            $out | Should-MatchString '1 Warning'
        }

        It 'shows an error header when a file-level check is a FatalError' {
            $files = @( New-FileResult -Check @([pscustomobject]@{ Type = 'FatalError'; Name = 'e'; Description = 'd' }) )
            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html
            $out | Should-MatchString '✖'
            $out | Should-MatchString '1 Error'
        }

        It 'renders the Outlook header without a VML roundrect wrapper' {
            $out = Build-MatrixEmailHtmlHC -FileResults @( New-FileResult ) -Html $html

            $out | Should-NotMatchString 'v:roundrect'
            $out | Should-NotMatchString 'v-text-anchor:top'
            $out | Should-MatchString "valign='middle' width='34'"
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

            $out | Should-MatchString 'Settings \(2\)'
            $out | Should-MatchString 'SRV01'
            $out | Should-MatchString 'SRV02'
            # The section label is wrapped in a margin-zero paragraph so Outlook
            # doesn't add ~12px of extra space between the title and the rows
            $out | Should-MatchString "<p style='margin:0; mso-line-height-rule:exactly; line-height:14px;'>Settings \(2\)</p>"
            ([regex]::Matches($out, 'bgcolor="#ffffff" height="4" style="font-size:0; line-height:4px; mso-line-height-rule:exactly; background-color:#ffffff;">&#160;</td>')).Count | Should-Be 1
            $out | Should-NotMatchString 'height="6" style="font-size:0; line-height:0;">&#160;</td>'
        }

        It 'shows the empty-state message when there are no matrices and no issues' {
            $files = @( New-FileResult -Matrices @() )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'No settings rows were processed for this file\.'
        }

        It 'renders a File Issues section when a file-level check exists' {
            $files = @(
                New-FileResult -Check @(
                    [pscustomobject]@{ Type = 'Warning'; Name = 'fileCheck'; Description = 'desc' }
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString 'File Issues \(1\)'
            $out | Should-MatchString 'fileCheck'
        }

        It 'renders settings rows as Skipped when the file has a file-level error' {
            $files = @(
                New-FileResult -Check @(
                    [pscustomobject]@{ Type = 'FatalError'; Name = 'Runspace processing failed'; Description = 'boom' }
                ) -Matrices @(
                    New-MatrixRow -ID 1 -ComputerName 'SRV01'
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString '>Skipped</span>'
            $out | Should-MatchString '#6b7280'
        }
    }

    Context 'card ordering' {
        BeforeAll {
            function Get-RenderedCardOrder {
                param([string]$Html)
                return @(
                    [regex]::Matches($Html, 'text-decoration:none;">([^<]+\.xlsx)</a>') |
                    ForEach-Object { $_.Groups[1].Value }
                )
            }

            # These tests are about ORDER, so they compare a single joined
            # string rather than a collection. Should-BeCollection compares
            # membership only - it would pass no matter how the cards were
            # arranged, making every assertion below meaningless.
            function Get-RenderedCardOrderText {
                param([string]$Html)
                return ((Get-RenderedCardOrder -Html $Html) -join ' > ')
            }
        }

        It 'sorts the cards alphabetically by matrix file name' {
            # Regression: $Context.FileResults arrives in runspace completion
            # order, so the cards used to appear in a different order on
            # every run.
            $files = @(
                New-FileResult -Name 'NOR KYN.xlsx'
                New-FileResult -Name 'DNK HCPT.xlsx'
                New-FileResult -Name 'SWE CEM CR.xlsx'
                New-FileResult -Name 'NOR CON.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            Get-RenderedCardOrderText -Html $out |
            Should-Be 'DNK HCPT.xlsx > NOR CON.xlsx > NOR KYN.xlsx > SWE CEM CR.xlsx'
        }

        It 'floats error cards above warning cards above clean cards' {
            $files = @(
                New-FileResult -Name 'Clean.xlsx'
                New-FileResult -Name 'Warned.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'Warning'; Name = 'w'; Description = 'd' }
                )
                New-FileResult -Name 'Errored.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'FatalError'; Name = 'e'; Description = 'd' }
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            Get-RenderedCardOrderText -Html $out |
            Should-Be 'Errored.xlsx > Warned.xlsx > Clean.xlsx'
        }

        It 'ranks incorrect cards between the errors and the warnings' {
            $files = @(
                New-FileResult -Name 'Clean.xlsx'
                New-FileResult -Name 'Warned.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'Warning'; Name = 'w'; Description = 'd' }
                )
                New-FileResult -Name 'Drifted.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'Incorrect'; Name = 'i'; Description = 'd' }
                )
                New-FileResult -Name 'Errored.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'FatalError'; Name = 'e'; Description = 'd' }
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            Get-RenderedCardOrderText -Html $out |
            Should-Be 'Errored.xlsx > Drifted.xlsx > Warned.xlsx > Clean.xlsx'
        }

        It 'does not promote a card that only has Fixed checks' {
            # A corrected permission is an outcome, not an outstanding issue.
            $files = @(
                New-FileResult -Name 'aClean.xlsx'
                New-FileResult -Name 'zFixed.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'Fixed'; Name = 'f'; Description = 'd' }
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            Get-RenderedCardOrderText -Html $out | Should-Be 'aClean.xlsx > zFixed.xlsx'
        }

        It 'gives an incorrect-only card the orange header and its own glyph' {
            $files = @(
                New-FileResult -Name 'Drifted.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'Incorrect'; Name = 'i'; Description = 'd' }
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString '#ea580c'
            $out | Should-MatchString '≠'
            $out | Should-MatchString '1 Incorrect'
        }

        It 'gives a fixed-only card the success header' {
            $files = @(
                New-FileResult -Name 'Corrected.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'Fixed'; Name = 'f'; Description = 'd' }
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should-MatchString '✓'
            $out | Should-MatchString '1 Fixed'
        }

        It 'sorts alphabetically within each severity group' {
            $files = @(
                New-FileResult -Name 'zClean.xlsx'
                New-FileResult -Name 'zErr.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'FatalError'; Name = 'e'; Description = 'd' }
                )
                New-FileResult -Name 'aClean.xlsx'
                New-FileResult -Name 'zWarn.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'Warning'; Name = 'w'; Description = 'd' }
                )
                New-FileResult -Name 'aErr.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'FatalError'; Name = 'e'; Description = 'd' }
                )
                New-FileResult -Name 'aWarn.xlsx' -Check @(
                    [pscustomobject]@{ Type = 'Warning'; Name = 'w'; Description = 'd' }
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            Get-RenderedCardOrderText -Html $out |
            Should-Be 'aErr.xlsx > zErr.xlsx > aWarn.xlsx > zWarn.xlsx > aClean.xlsx > zClean.xlsx'
        }

        It 'does not promote a card that only has Information checks' {
            # Info notices are not issues - they must stay in the
            # alphabetical run with the other clean files.
            $files = @(
                New-FileResult -Name 'aClean.xlsx'
                New-FileResult -Name 'zInfo.xlsx' -Matrices @(
                    New-MatrixRow -ID 1 -Check @(
                        [pscustomobject]@{ Type = 'Information'; Name = 'i'; Description = 'd' }
                    )
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            Get-RenderedCardOrderText -Html $out | Should-Be 'aClean.xlsx > zInfo.xlsx'
        }

        It 'ranks a card by a check on one of its matrices' {
            $files = @(
                New-FileResult -Name 'aClean.xlsx'
                New-FileResult -Name 'zMatrixErr.xlsx' -Matrices @(
                    New-MatrixRow -ID 1 -Check @(
                        [pscustomobject]@{ Type = 'FatalError'; Name = 'e'; Description = 'd' }
                    )
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            Get-RenderedCardOrderText -Html $out | Should-Be 'zMatrixErr.xlsx > aClean.xlsx'
        }

        It 'ranks a card by a check on one of its sheets' {
            $files = @(
                New-FileResult -Name 'aClean.xlsx'
                New-FileResult -Name 'zSheetWarn.xlsx' -PermissionsCheck @(
                    [pscustomobject]@{ Type = 'Warning'; Name = 'w'; Description = 'd' }
                )
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            Get-RenderedCardOrderText -Html $out | Should-Be 'zSheetWarn.xlsx > aClean.xlsx'
        }

        It 'produces the same output regardless of the input order' {
            $names = @('Zeta.xlsx', 'Alpha.xlsx', 'Mike.xlsx')

            $forward = Build-MatrixEmailHtmlHC -Html $html -FileResults @(
                $names | ForEach-Object { New-FileResult -Name $_ }
            )
            $reversed = Build-MatrixEmailHtmlHC -Html $html -FileResults @(
                $names | Sort-Object -Descending | ForEach-Object { New-FileResult -Name $_ }
            )

            $forward | Should-Be $reversed
        }

        It 'sorts case-insensitively' {
            $files = @(
                New-FileResult -Name 'beta.xlsx'
                New-FileResult -Name 'Alpha.xlsx'
                New-FileResult -Name 'CHARLIE.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            Get-RenderedCardOrderText -Html $out |
            Should-Be 'Alpha.xlsx > beta.xlsx > CHARLIE.xlsx'
        }

        It 'renders every card exactly once' {
            $files = @(
                New-FileResult -Name 'One.xlsx'
                New-FileResult -Name 'Two.xlsx'
                New-FileResult -Name 'Three.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            (Get-RenderedCardOrder -Html $out).Count | Should-Be 3
        }

        It 'still renders a single card' {
            $out = Build-MatrixEmailHtmlHC -Html $html -FileResults @(
                New-FileResult -Name 'Only.xlsx'
            )

            Get-RenderedCardOrderText -Html $out | Should-Be 'Only.xlsx'
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
        $out | Should-MatchString '<!DOCTYPE html>'
        $out | Should-MatchString '</html>'
    }

    It 'renders the script name in an encoded h1' {
        $settings = [pscustomobject]@{ ScriptName = 'R&D Run'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')
        $out | Should-MatchString '<h1>R&amp;D Run</h1>'
    }

    It 'falls back to a default script name when none is supplied' {
        $settings = [pscustomobject]@{ ScriptName = ''; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')
        $out | Should-MatchString '<h1>Permission Matrix</h1>'
    }

    It 'renders a footer with Started, Ended and Duration when a start time is given' {
        $settings = [pscustomobject]@{ ScriptName = 'S'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00') `
            -ScriptEndTime (Get-Date '2024-01-01 08:30:00')
        $out | Should-MatchString 'Started'
        $out | Should-MatchString 'Ended'
        $out | Should-MatchString 'Duration'
        $out | Should-MatchString '00:30:00'
    }

    It 'includes the MatrixTables fragment passed via the Html hashtable' {
        $html.MatrixTables = '<!-- MATRIX_TABLES_MARKER -->'
        $settings = [pscustomobject]@{ ScriptName = 'S'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')
        $out | Should-MatchString 'MATRIX_TABLES_MARKER'
    }

    It 'renders system error cards from a [ref] SystemErrors entry' {
        $errors = [System.Collections.Generic.List[object]]::new()
        $errors.Add([pscustomobject]@{ Type = 'FatalError'; Name = 'SysBoom'; Message = 'bad'; Category = 'Matrix' })
        $html.SystemErrors = ([ref]$errors)
        $settings = [pscustomobject]@{ ScriptName = 'S'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')
        $out | Should-MatchString 'SysBoom'
        $out | Should-MatchString 'System Error'
    }

    It 'includes browser-view and export links when supplied' {
        $settings = [pscustomobject]@{ ScriptName = 'S'; SendMail = [pscustomobject]@{ Body = '' } }
        $out = Get-MailBodyHtmlHC -Settings $settings -Html $html `
            -BrowserViewFilePath 'C:\logs\Mail - Run.html' `
            -ExportedFiles ([ordered]@{ Permissions = 'C:\reports\Permissions.xlsx' }) `
            -ScriptStartTime (Get-Date '2024-01-01 08:00:00')

        $out | Should-MatchString 'click here to view it in the browser'
        $out | Should-MatchString 'Permissions Excel'
    }
}