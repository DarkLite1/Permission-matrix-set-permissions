#requires -Modules Pester

Describe 'Html.ps1 consolidated tests' {

    BeforeAll {
        . "$PSScriptRoot/../Modules/Toolbox.PermissionMatrixHC/Private/Utils.ps1"
        . "$PSScriptRoot/../Modules/Toolbox.PermissionMatrixHC/Private/Html.ps1"
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