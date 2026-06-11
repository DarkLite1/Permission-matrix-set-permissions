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

            # Each card is anchored by its footer link list. These
            # fixtures have no ReportFilePath or LogMatrixFilePath, so
            # the fallback 'Open matrix file' link renders once per card.
            ([regex]::Matches($out, 'Open matrix file')).Count | Should -Be 3
        }
    }

    Context 'footer links' {
        It 'renders an execution report link when ReportFilePath is set' {
            $files = @(
                New-FileResult `
                    -ReportFilePath 'C:\logs\00 - Execution Report.html'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'Open execution report'
            $out | Should -Match "href='file://C:/logs/00%20-%20Execution%20Report\.html'"
            # No fallback link when a log artifact exists
            $out | Should -Not -Match 'Open matrix file'
        }

        It 'renders a matrix Excel file link when LogMatrixFilePath is set' {
            $files = @(
                New-FileResult -LogMatrixFilePath 'C:\logs\A.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'Open matrix Excel file'
            $out | Should -Match "href='file://C:/logs/A\.xlsx'"
            $out | Should -Not -Match 'Open matrix file'
            $out | Should -Not -Match 'Open execution report'
        }

        It 'puts the raw Windows path in the tooltip of the matrix Excel link' {
            $files = @(
                New-FileResult -LogMatrixFilePath 'C:\logs\A.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'title="C:\\logs\\A\.xlsx"'
        }

        It 'renders both links separated by a middot when both artifacts exist' {
            $files = @(
                New-FileResult `
                    -ReportFilePath 'C:\logs\report.html' `
                    -LogMatrixFilePath 'C:\logs\A.xlsx'
            )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'Open execution report'
            $out | Should -Match 'Open matrix Excel file'
            # The two anchors are joined by a middot separator span
            $out | Should -Match 'Open execution report &rarr;</a><span[^>]*>&middot;</span><a'
        }

        It 'falls back to the source matrix file link when no log artifacts exist' {
            $files = @( New-FileResult -FullName 'C:\share\A.xlsx' )

            $out = Build-MatrixEmailHtmlHC -FileResults $files -Html $html

            $out | Should -Match 'Open matrix file'
            # The fallback footer anchor uses single-quoted href, unlike
            # the double-quoted header link to the same file URL
            $out | Should -Match "href='file://C:/share/A\.xlsx'"
            $out | Should -Not -Match 'Open execution report'
            $out | Should -Not -Match 'Open matrix Excel file'
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