# tests for mail sending, log folder creation, counters

<# 
Email sending
Counters
Error/warning matrices
HTML output
Cleanup
Event log writing
 #>

BeforeDiscovery {
    # Dot-source common helpers only
    $root = Split-Path -Parent $MyInvocation.MyCommand.Path
    . (Join-Path $root '../Helpers/Helpers.HC.ps1')
    . (Join-Path $root '../Helpers/Fixtures.Excel.ps1')
    . (Join-Path $root '../Helpers/Fixtures.Json.ps1')

    # Define static TestCases here
    $InvalidPathTests = @(
        @{ Property = 'Matrix.FolderPath'  ; Value = 'x:\NotExisting' }
        @{ Property = 'Matrix.DefaultsFile'; Value = 'x:\NotExisting.xlsx' }
    )
}

BeforeAll {
    # Dynamic TestDrive content belongs here.
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    $jsonFile = New-Item 'TestDrive:/input.json' -ItemType File
    $settingsFile = New-Item 'TestDrive:/Defaults.xlsx' -Type File
    $matrixFolder = New-Item 'TestDrive:/Matrices' -Type Directory

    $testInputJson = New-JsonFixture     # from Fixtures.Json.ps1
    $testExcel = New-ExcelDefaults   # from Fixtures.Excel.ps1

    $testExcel | Export-Excel -Path $settingsFile.FullName -WorksheetName Settings

    $testInputJson.Matrix.FolderPath = $matrixFolder.FullName
    $testInputJson.Matrix.DefaultsFile = $settingsFile.FullName

    $testInputJson | ConvertTo-Json -Depth 10 | Set-Content $jsonFile.FullName
}

BeforeEach {
    # Clear mocks for isolation
    Mock Invoke-Command
    Mock Send-MailKitMessageHC
    Mock Write-EventLog
}