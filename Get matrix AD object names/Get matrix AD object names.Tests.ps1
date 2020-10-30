#Requires -Module Assert, Pester

BeforeAll {
    # $VerbosePreference = 'SilentlyContinue'
    $VerbosePreference = 'Continue'
    $WarningPreference = 'SilentlyContinue'

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    $ScriptAdmin = 'Brecht.Gijbels@heidelbergcement.com'

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and 
        ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }
    
    $Params = @{
        ScriptName  = 'Test'
        Path        = (New-Item -Path "TestDrive:\Matrix.xlsx" -ItemType File -EA Ignore).FullName
        MailTo      = $ScriptAdmin
        ScriptAdmin = $ScriptAdmin
        LogFolder   = (New-Item -Path "TestDrive:\Log" -ItemType Directory -EA Ignore).FullName
    }

    Mock Send-MailHC
    Mock Write-EventLog
}

Describe 'error handling' {    
    Context 'mandatory parameters' {
        It '<Name>' -TestCases @(
            @{Name = 'ScriptName' }
            @{Name = 'Path' }
            @{Name = 'MailTo' }
        ) {
            (Get-Command $testScript).Parameters[$Name].Attributes.Mandatory |
            Should -BeTrue
        }
    }
    Context 'not found' {
        It 'LogFolder' {
            $testParams = $Params.Clone()
            $testParams.LogFolder = 'NotExistingLogFolder'
            . $testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -Scope It -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Path 'NotExistingLogFolder' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -Scope It -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'Excel file' {
            $testParams = $Params.Clone()
            $testParams.Path = 'NotExistingExcelFile.xlsx'
            . $testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -Scope It -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Path 'NotExistingExcelFile.xlsx' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -Scope It -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
    }
}

Describe "mail the Excel file to the user in 'MailTo'" {
    It 'one Excel input file, mail attachment has one worksheet' {
        Mock Get-MatrixAdObjectNamesHC { 
            [PSCustomObject]@{
                fileName = 'MatrixA.xlsx'
                AdObject = @('bob', 'mike')
            }
        }
                    
        . $testScript @Params

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($MailTo -eq $MailTo) -and 
            ($Subject -eq 'Success') -and 
            ($Attachments -like '*.xlsx')
        }
        Should -Invoke Get-MatrixAdObjectNamesHC -Exactly 1
                    
        $Expected = @(
            [PSCustomObject]@{ 'MatrixA.xlsx' = 'bob' }
            [PSCustomObject]@{ 'MatrixA.xlsx' = 'mike' }
        )
                    
        $Actual = Import-Excel $ExportParams.Path
        Assert-Equivalent -Actual $Actual -Expected $Expected
    }
    It 'two Excel input files, mail attachment has two worksheets' {
        Mock Get-MatrixAdObjectNamesHC { 
            [PSCustomObject]@{
                fileName = 'MatrixA.xlsx'
                AdObject = @('bob', 'mike')
            }
            [PSCustomObject]@{
                fileName = 'MatrixB.xlsx'
                AdObject = @('jake', 'chuck', 'victor')
            }
        }
                    
        .$testScript @Params

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($MailTo -eq $MailTo) -and 
            ($Subject -eq 'Success') -and 
            ($Attachments -like '*.xlsx')
        }
        Should -Invoke Get-MatrixAdObjectNamesHC -Exactly 1
                    
        $Expected = @(
            [PSCustomObject]@{ 'MatrixA.xlsx' = 'bob' }
            [PSCustomObject]@{ 'MatrixA.xlsx' = 'mike' }
        )
        $Actual = Import-Excel $ExportParams.Path -WorksheetName "1 MatrixA.xlsx"
        Assert-Equivalent -Actual $Actual -Expected $Expected
         
        $Expected = @(
            [PSCustomObject]@{ 'MatrixB.xlsx' = 'jake' }
            [PSCustomObject]@{ 'MatrixB.xlsx' = 'chuck' }
            [PSCustomObject]@{ 'MatrixB.xlsx' = 'victor' }
        )
 
        $Actual = Import-Excel $ExportParams.Path -WorksheetName "2 MatrixB.xlsx"
        Assert-Equivalent -Actual $Actual -Expected $Expected
    }
}
