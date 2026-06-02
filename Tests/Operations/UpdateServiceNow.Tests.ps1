#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

<#
    Tests for Operations\UpdateServiceNow.ps1

    Approach:
        UpdateServiceNow.ps1 is a *script* (param + begin/process/end), invoked
        with the call operator. Its helpers (Get-StringValueHC,
        New-ServiceNowSessionHC) live inside begin{} so they can't be dot-sourced
        and unit-tested in isolation; instead the script is run end to end and its
        dependencies are mocked.

        - ServiceNow cmdlets are mocked: New-ServiceNowSession, Get-ServiceNowRecord,
          Remove-ServiceNowRecord, New-ServiceNowRecord. Because the calls happen in
          the script (not inside an imported module), test-scope mocks intercept
          them without -ModuleName.
        - ImportExcel is used for real: fixtures are written with Export-Excel and
          read back by the script's Import-Excel.
        - Start-Sleep is mocked so retry tests don't wait.
        - The credentials JSON is a real file under TestDrive.

    The ServiceNow and ImportExcel modules must be installed (the script #Requires
    them, and Pester can only mock commands that exist).
#>

BeforeAll {
    $script:ScriptPath = "$PSScriptRoot\..\..\Scripts\Operations\UpdateServiceNow.ps1"

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Script under test not found: '$ScriptPath'. Adjust the path resolution for this test's location."
    }

    foreach ($module in 'ImportExcel', 'ServiceNow') {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Module '$module' is required to run these tests. Install it with: Install-Module $module"
        }
        Import-Module $module -ErrorAction Stop
    }

    function New-ValidSnowEnvironment {
        @{
            Uri          = 'https://prod.example.service-now.com'
            UserName     = 'snow-user'
            Password     = 'snow-pass'
            ClientId     = 'client-id'
            ClientSecret = 'client-secret'
        }
    }

    function New-SnowCredsFile {
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][hashtable]$Environment,
            [string]$Name = 'Prod'
        )
        @{ $Name = $Environment } |
            ConvertTo-Json -Depth 5 |
            Set-Content -LiteralPath $Path -Encoding UTF8
    }

    function New-SampleRecords {
        param([int]$Count = 2)
        1..$Count | ForEach-Object {
            [PSCustomObject]@{
                u_matrixfilename = "matrix$_.xlsx"
                u_adobjectname   = "GRP-$_"
                u_adobjectid     = "S-1-5-$_"
            }
        }
    }

    function New-RecordsXlsx {
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][object[]]$Records,
            [string]$WorksheetName = 'SnowFormData'
        )
        $Records | Export-Excel -Path $Path -WorksheetName $WorksheetName
    }
}

Describe 'UpdateServiceNow.ps1' {
    BeforeEach {
        Remove-Item (Join-Path $TestDrive '*') -Recurse -Force -ErrorAction Ignore

        $credsFile = Join-Path $TestDrive 'creds.json'
        $recordsFile = Join-Path $TestDrive 'records.xlsx'

        New-SnowCredsFile -Path $credsFile -Environment (New-ValidSnowEnvironment)
        New-RecordsXlsx -Path $recordsFile -Records (New-SampleRecords -Count 2)

        $params = @{
            CredentialsFilePath    = $credsFile
            Environment            = 'Prod'
            FormDataExcelFilePath  = $recordsFile
            ExcelFileWorksheetName = 'SnowFormData'
            TableName              = 'u_bnl_roles'
            MaxRetries             = 3
        }

        # Default: a populated table (two existing records) so the happy path
        # exercises removal. Tests override where needed.
        Mock New-ServiceNowSession {}
        Mock Get-ServiceNowRecord {
            @(
                [PSCustomObject]@{ sys_id = 'rec-1' }
                [PSCustomObject]@{ sys_id = 'rec-2' }
            )
        }
        Mock Remove-ServiceNowRecord {}
        Mock New-ServiceNowRecord {}
        Mock Start-Sleep {}
        Mock Write-Warning {}
    }

    AfterEach {
        Remove-Item Env:\SNOW_TEST_URI -ErrorAction Ignore
        Remove-Item Env:\SNOW_MISSING_VAR -ErrorAction Ignore
    }

    Context 'credentials & environment validation' {
        It 'throws when the credentials file does not exist' {
            $params.CredentialsFilePath = Join-Path $TestDrive 'missing.json'

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*ServiceNow credentials file*'

            Should -Invoke New-ServiceNowSession -Times 0
        }

        It 'throws when the requested environment is not in the file' {
            $params.Environment = 'Dev'   # file only contains 'Prod'

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage "*Failed to find environment 'Dev'*"
        }

        It 'throws when the <Property> property is missing for the environment' -TestCases @(
            @{ Property = 'Uri' }
            @{ Property = 'UserName' }
            @{ Property = 'Password' }
            @{ Property = 'ClientId' }
            @{ Property = 'ClientSecret' }
        ) {
            param($Property)

            $environment = New-ValidSnowEnvironment
            $environment.Remove($Property)
            New-SnowCredsFile -Path $params.CredentialsFilePath -Environment $environment

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage "*Property '$Property' not found*"
        }
    }

    Context 'reading records from Excel' {
        It 'throws a clear error when the Excel file cannot be read' {
            $params.FormDataExcelFilePath = Join-Path $TestDrive 'no-such-file.xlsx'

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*Failed to import records to upload*'
        }

        It 'does nothing when there are no records to upload' {
            # A genuinely empty worksheet is awkward to build with Export-Excel,
            # so for this single case Import-Excel is mocked to return nothing.
            Mock Import-Excel {}

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke New-ServiceNowSession -Times 0
            Should -Invoke Get-ServiceNowRecord -Times 0
            Should -Invoke Remove-ServiceNowRecord -Times 0
            Should -Invoke New-ServiceNowRecord -Times 0
        }
    }

    Context 'creating the ServiceNow session' {
        It 'creates a session using the configured Uri' {
            & $ScriptPath @params

            Should -Invoke New-ServiceNowSession -Times 1
            Should -Invoke New-ServiceNowSession -Times 1 -ParameterFilter {
                $Url -eq 'https://prod.example.service-now.com'
            }
        }

        It 'resolves an ENV: credential value from the environment' {
            $env:SNOW_TEST_URI = 'https://env.example.service-now.com'

            $environment = New-ValidSnowEnvironment
            $environment.Uri = 'ENV:SNOW_TEST_URI'
            New-SnowCredsFile -Path $params.CredentialsFilePath -Environment $environment

            & $ScriptPath @params

            Should -Invoke New-ServiceNowSession -Times 1 -ParameterFilter {
                $Url -eq 'https://env.example.service-now.com'
            }
        }

        It 'throws when an ENV: credential points to a missing variable' {
            $environment = New-ValidSnowEnvironment
            $environment.Uri = 'ENV:SNOW_MISSING_VAR'
            New-SnowCredsFile -Path $params.CredentialsFilePath -Environment $environment

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage "*Environment variable 'SNOW_MISSING_VAR' not found*"
        }

        It 'throws a clear error when the session cannot be created' {
            Mock New-ServiceNowSession { throw 'connection refused' }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*Failed to create a ServiceNow session*'

            Should -Invoke New-ServiceNowRecord -Times 0
        }
    }

    Context 'removing existing records' {
        It 'removes every existing record before uploading' {
            & $ScriptPath @params

            Should -Invoke Get-ServiceNowRecord -Times 1 -ParameterFilter {
                $Table -eq 'u_bnl_roles'
            }
            Should -Invoke Remove-ServiceNowRecord -Times 2
        }

        It 'skips removal when the table is already empty' {
            Mock Get-ServiceNowRecord {}

            & $ScriptPath @params

            Should -Invoke Remove-ServiceNowRecord -Times 0
            # Upload still happens.
            Should -Invoke New-ServiceNowRecord -Times 2
        }

        It 'retries a failed removal up to MaxRetries and then continues' {
            $params.MaxRetries = 2
            Mock Get-ServiceNowRecord { @([PSCustomObject]@{ sys_id = 'rec-1' }) }
            Mock Remove-ServiceNowRecord { throw 'transient remove error' }

            { & $ScriptPath @params } | Should -Not -Throw

            # One record, two attempts, sleep between each failed attempt.
            Should -Invoke Remove-ServiceNowRecord -Times 2
            Should -Invoke Start-Sleep -Times 2
            # A removal that never succeeds is non-fatal: upload still proceeds.
            Should -Invoke New-ServiceNowRecord -Times 2
        }
    }

    Context 'creating new records' {
        It 'creates a record for every row in the worksheet' {
            & $ScriptPath @params

            Should -Invoke New-ServiceNowRecord -Times 2 -ParameterFilter {
                $Table -eq 'u_bnl_roles'
            }
        }

        It 'retries a transient creation failure and then succeeds' {
            $singleRecordFile = Join-Path $TestDrive 'one-record.xlsx'
            New-RecordsXlsx -Path $singleRecordFile -Records (New-SampleRecords -Count 1)
            $params.FormDataExcelFilePath = $singleRecordFile
            $params.MaxRetries = 3
            Mock Get-ServiceNowRecord {}   # isolate creation from removal

            $script:createCalls = 0
            Mock New-ServiceNowRecord {
                $script:createCalls++
                if ($script:createCalls -eq 1) { throw 'transient create error' }
            }

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke New-ServiceNowRecord -Times 2   # one failure + one success
            Should -Invoke Start-Sleep -Times 1
        }

        It 'throws CRITICAL FAILURE after exhausting MaxRetries on creation' {
            $singleRecordFile = Join-Path $TestDrive 'one-record.xlsx'
            New-RecordsXlsx -Path $singleRecordFile -Records (New-SampleRecords -Count 1)
            $params.FormDataExcelFilePath = $singleRecordFile
            $params.MaxRetries = 2
            Mock Get-ServiceNowRecord {}
            Mock New-ServiceNowRecord { throw 'persistent create error' }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*CRITICAL FAILURE*'

            Should -Invoke New-ServiceNowRecord -Times 2
            Should -Invoke Start-Sleep -Times 1   # only between attempts, not after the final throw
        }
    }

    Context 'happy path' {
        It 'creates a session, clears the table, and uploads all records' {
            & $ScriptPath @params

            Should -Invoke New-ServiceNowSession -Times 1
            Should -Invoke Get-ServiceNowRecord -Times 1
            Should -Invoke Remove-ServiceNowRecord -Times 2
            Should -Invoke New-ServiceNowRecord -Times 2
        }
    }
}
