#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

<#
    Tests for Operations\UpdateServiceNow.ps1

    Approach:
        UpdateServiceNow.ps1 is a *script* (param + begin/process/end), invoked
        with the call operator. Its helpers (Get-StringValueHC,
        New-ServiceNowSessionHC, Get-RecordComparisonKeyHC, Test-RecordActiveHC,
        Invoke-WithRetryHC) live inside begin{} so they can't be dot-sourced and
        unit-tested in isolation; instead the script is run end to end and its
        dependencies are mocked.

        The script now performs a *differential sync* rather than a destructive
        drop-and-reload:
            - source row absent from the table   -> New-ServiceNowRecord (u_active TRUE)
            - source row matching an inactive row -> Update-ServiceNowRecord (u_active TRUE)
            - table row absent from the source    -> Update-ServiceNowRecord (u_active FALSE)
            - table row already inactive & absent -> left untouched
        Matching is on every source column except u_active and sys_* fields.

        - ServiceNow cmdlets are mocked: New-ServiceNowSession, Get-ServiceNowRecord,
          New-ServiceNowRecord, Update-ServiceNowRecord. Because the calls happen in
          the script (not inside an imported module), test-scope mocks intercept
          them without -ModuleName.
        - ImportExcel is used for real: fixtures are written with Export-Excel and
          read back by the script's Import-Excel.
        - Start-Sleep is mocked so retry tests don't wait.
        - The credentials JSON is a real file under TestDrive.

    The ServiceNow and ImportExcel modules must be installed (the script #Requires
    them, and Pester can only mock commands that exist).

    Assumptions about the ServiceNow module surface (adjust if your installed
    version differs):
        - Get-ServiceNowRecord supports -First and -Skip paging.
        - Update-ServiceNowRecord accepts -Table, -ID and -Values <hashtable>.

    Note on -Exactly:
        Every Should -Invoke below uses -Exactly. In Pester 5, `-Times N` without
        `-Exactly` means "at least N", so `-Times 0` ("at least 0") never fails and
        `-Times 2` would still pass if the command were called 3 times. -Exactly
        turns these into the strict counts the tests actually intend.
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

    # Build a table row equivalent to source record N (so its comparison key
    # matches), with a controllable sys_id and u_active state.
    function New-MatchingTableRecord {
        param(
            [Parameter(Mandatory)][int]$Index,
            [Parameter(Mandatory)][string]$SysId,
            [object]$Active = $true
        )
        [PSCustomObject]@{
            sys_id           = $SysId
            u_matrixfilename = "matrix$Index.xlsx"
            u_adobjectname   = "GRP-$Index"
            u_adobjectid     = "S-1-5-$Index"
            u_active         = $Active
        }
    }

    # Build a table row that does NOT correspond to any source record.
    function New-StaleTableRecord {
        param(
            [Parameter(Mandatory)][string]$SysId,
            [object]$Active = $true
        )
        [PSCustomObject]@{
            sys_id           = $SysId
            u_matrixfilename = 'matrix-stale.xlsx'
            u_adobjectname   = 'GRP-STALE'
            u_adobjectid     = 'S-1-5-999'
            u_active         = $Active
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

        # Default: an empty table. With the two default source rows, the happy
        # path therefore creates two new (active) records and updates nothing.
        # Tests that need existing rows override Get-ServiceNowRecord. Note the
        # script reads the table up front in chunks of -ChunkSize (default 200)
        # using -First/-Skip. A short/empty chunk ends the loop, so a single
        # returned batch of fewer than 200 rows is fetched in one Get call.
        Mock New-ServiceNowSession {}
        Mock Get-ServiceNowRecord {}
        Mock New-ServiceNowRecord {}
        Mock Update-ServiceNowRecord {}
        Mock Start-Sleep {}
        Mock Write-Warning {}
    }

    AfterEach {
        Remove-Item Env:\SNOW_TEST_URI -ErrorAction Ignore
        Remove-Item Env:\SNOW_MISSING_VAR -ErrorAction Ignore
    }

    Context 'parameter validation' {
        It 'rejects a MaxRetries below 1' {
            $params.MaxRetries = 0

            { & $ScriptPath @params } | Should -Throw

            Should -Invoke New-ServiceNowSession -Exactly -Times 0
        }
    }

    Context 'credentials & environment validation' {
        It 'throws when the credentials file does not exist' {
            $params.CredentialsFilePath = Join-Path $TestDrive 'missing.json'

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*ServiceNow credentials file*'

            Should -Invoke New-ServiceNowSession -Exactly -Times 0
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

            Should -Invoke New-ServiceNowSession -Exactly -Times 0
            Should -Invoke Get-ServiceNowRecord -Exactly -Times 0
            Should -Invoke New-ServiceNowRecord -Exactly -Times 0
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 0
        }
    }

    Context 'creating the ServiceNow session' {
        It 'creates a session using the configured Uri' {
            & $ScriptPath @params

            Should -Invoke New-ServiceNowSession -Exactly -Times 1
            Should -Invoke New-ServiceNowSession -Exactly -Times 1 -ParameterFilter {
                $Url -eq 'https://prod.example.service-now.com'
            }
        }

        It 'resolves an ENV: credential value from the environment' {
            $env:SNOW_TEST_URI = 'https://env.example.service-now.com'

            $environment = New-ValidSnowEnvironment
            $environment.Uri = 'ENV:SNOW_TEST_URI'
            New-SnowCredsFile -Path $params.CredentialsFilePath -Environment $environment

            & $ScriptPath @params

            Should -Invoke New-ServiceNowSession -Exactly -Times 1 -ParameterFilter {
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

            Should -Invoke New-ServiceNowRecord -Exactly -Times 0
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 0
        }
    }

    Context 'reading the existing table' {
        It 'reads the target table for comparison' {
            & $ScriptPath @params

            Should -Invoke Get-ServiceNowRecord -Exactly -Times 1 -ParameterFilter {
                $Table -eq 'u_bnl_roles'
            }
        }

        It 'reads the existing table in chunks of -ChunkSize until drained' {
            # Chunks of 2: two full chunks then a short chunk of 1 ends the loop.
            # The reader pages with -First/-Skip; the mock ignores -Skip and just
            # serves the next batch by call count. Rows carry no business fields,
            # so they never match a source row (and are inactive, so they trigger
            # no deactivation) - this test is purely about how many fetches happen.
            $params.ChunkSize = 2

            $script:getCalls = 0
            Mock Get-ServiceNowRecord {
                $script:getCalls = ([int]$script:getCalls) + 1
                if ($script:getCalls -le 2) {
                    [PSCustomObject]@{ sys_id = "p$($script:getCalls)-a"; u_active = $false }
                    [PSCustomObject]@{ sys_id = "p$($script:getCalls)-b"; u_active = $false }
                }
                else {
                    [PSCustomObject]@{ sys_id = 'p-final'; u_active = $false }
                }
            }

            & $ScriptPath @params

            # Three fetches (2 + 2 + 1) proves the table is read a chunk at a time.
            Should -Invoke Get-ServiceNowRecord -Exactly -Times 3
        }

        It 'retries a failed chunk read and then aborts (a partial read is fatal)' {
            # An unreadable table must NOT be treated as "all rows are gone", so a
            # chunk that fails all retries throws rather than continuing. With
            # MaxRetries 2 the single chunk is attempted twice (one sleep between)
            # before the script gives up.
            $params.MaxRetries = 2
            Mock Get-ServiceNowRecord { throw 'network down' }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*Failed to retrieve records*'

            Should -Invoke Get-ServiceNowRecord -Exactly -Times 2
            Should -Invoke Start-Sleep -Exactly -Times 1
            # Nothing is written when the table can't be read.
            Should -Invoke New-ServiceNowRecord -Exactly -Times 0
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 0
        }
    }

    Context 'adding rows that are not yet in the table' {
        It 'creates every source row that is absent from the table' {
            # Default Get returns nothing -> the table is empty.
            & $ScriptPath @params

            Should -Invoke New-ServiceNowRecord -Exactly -Times 2 -ParameterFilter {
                $Table -eq 'u_bnl_roles'
            }
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 0

            # TODO: assert the created payload carries u_active = TRUE. New rows
            # are piped into New-ServiceNowRecord, so the assertion depends on the
            # pipeline-bound parameter name for the installed ServiceNow module
            # version. Add a -ParameterFilter once that name is confirmed.
        }
    }

    Context 'rows already present and active' {
        It 'leaves a matching, active row untouched' {
            Mock Get-ServiceNowRecord {
                @(
                    New-MatchingTableRecord -Index 1 -SysId 'rec-1' -Active $true
                    New-MatchingTableRecord -Index 2 -SysId 'rec-2' -Active $true
                )
            }

            & $ScriptPath @params

            Should -Invoke New-ServiceNowRecord -Exactly -Times 0
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 0
        }
    }

    Context 'reactivating a matched-but-inactive row' {
        It 'sets u_active TRUE for a matching row that was deactivated' {
            Mock Get-ServiceNowRecord {
                @(
                    New-MatchingTableRecord -Index 1 -SysId 'rec-1' -Active $false
                    New-MatchingTableRecord -Index 2 -SysId 'rec-2' -Active $true
                )
            }

            & $ScriptPath @params

            # Only the inactive match (rec-1) is reactivated; nothing is created
            # and nothing is deactivated.
            Should -Invoke New-ServiceNowRecord -Exactly -Times 0
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 1
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 1 -ParameterFilter {
                $ID -eq 'rec-1' -and $Values.u_active -eq $true
            }
        }
    }

    Context 'deactivating rows no longer in the collection' {
        It 'sets u_active FALSE for an active table row absent from the source' {
            Mock Get-ServiceNowRecord {
                @(
                    New-MatchingTableRecord -Index 1 -SysId 'rec-1' -Active $true
                    New-MatchingTableRecord -Index 2 -SysId 'rec-2' -Active $true
                    New-StaleTableRecord -SysId 'rec-stale' -Active $true
                )
            }

            & $ScriptPath @params

            Should -Invoke New-ServiceNowRecord -Exactly -Times 0
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 1
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 1 -ParameterFilter {
                $ID -eq 'rec-stale' -and $Values.u_active -eq $false
            }
        }

        It 'leaves an already-inactive straggler untouched' {
            Mock Get-ServiceNowRecord {
                @(
                    New-MatchingTableRecord -Index 1 -SysId 'rec-1' -Active $true
                    New-MatchingTableRecord -Index 2 -SysId 'rec-2' -Active $true
                    New-StaleTableRecord -SysId 'rec-stale' -Active $false
                )
            }

            & $ScriptPath @params

            Should -Invoke New-ServiceNowRecord -Exactly -Times 0
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 0
        }
    }

    Context 'a record that changed in one or more fields' {
        It 'deactivates the old row and adds the changed row as new' {
            # Source row 1 is matrix1/GRP-1/S-1-5-1. The table holds a row with the
            # same matrix and group but a different u_adobjectid, so it no longer
            # matches: the old row is deactivated and the source row is created.
            Mock Get-ServiceNowRecord {
                @(
                    [PSCustomObject]@{
                        sys_id           = 'rec-old'
                        u_matrixfilename = 'matrix1.xlsx'
                        u_adobjectname   = 'GRP-1'
                        u_adobjectid     = 'S-1-5-OLD'
                        u_active         = $true
                    }
                    New-MatchingTableRecord -Index 2 -SysId 'rec-2' -Active $true
                )
            }

            & $ScriptPath @params

            Should -Invoke New-ServiceNowRecord -Exactly -Times 1
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 1 -ParameterFilter {
                $ID -eq 'rec-old' -and $Values.u_active -eq $false
            }
        }
    }

    Context 'numeric source values (round-trip safety)' {
        It 'matches a numeric source value against its plain-string stored value' {
            # Excel stores the AD name 158774 as a number; ServiceNow stored it as
            # the string '158774'. These must compare equal - the row must NOT be
            # recreated on every run.
            $numericFile = Join-Path $TestDrive 'numeric.xlsx'
            [PSCustomObject]@{
                u_matrixfilename = 'matrix-n.xlsx'
                u_adobjectname   = 158774          # numeric -> read back as a number
                u_adobjectid     = 'S-1-5-7'
            } | Export-Excel -Path $numericFile -WorksheetName 'SnowFormData'
            $params.FormDataExcelFilePath = $numericFile

            Mock Get-ServiceNowRecord {
                @(
                    [PSCustomObject]@{
                        sys_id           = 'rec-n'
                        u_matrixfilename = 'matrix-n.xlsx'
                        u_adobjectname   = '158774'   # stored as a plain string
                        u_adobjectid     = 'S-1-5-7'
                        u_active         = $true
                    }
                )
            }

            & $ScriptPath @params

            Should -Invoke New-ServiceNowRecord -Exactly -Times 0
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 0
        }

        It 'cleans up a stored value corrupted with a trailing .0' {
            # A value stored as '158774.0' by an earlier (buggy) upload no longer
            # matches the source '158774', so it is deactivated and a clean copy is
            # created - a one-time correction, after which the row is stable.
            $numericFile = Join-Path $TestDrive 'numeric.xlsx'
            [PSCustomObject]@{
                u_matrixfilename = 'matrix-n.xlsx'
                u_adobjectname   = 158774
                u_adobjectid     = 'S-1-5-7'
            } | Export-Excel -Path $numericFile -WorksheetName 'SnowFormData'
            $params.FormDataExcelFilePath = $numericFile

            Mock Get-ServiceNowRecord {
                @(
                    [PSCustomObject]@{
                        sys_id           = 'rec-dirty'
                        u_matrixfilename = 'matrix-n.xlsx'
                        u_adobjectname   = '158774.0'   # corrupted by an earlier run
                        u_adobjectid     = 'S-1-5-7'
                        u_active         = $true
                    }
                )
            }

            & $ScriptPath @params

            Should -Invoke New-ServiceNowRecord -Exactly -Times 1
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 1 -ParameterFilter {
                $ID -eq 'rec-dirty' -and $Values.u_active -eq $false
            }
        }
    }

    Context 'retry behaviour' {
        It 'retries a transient creation failure and then succeeds' {
            $singleRecordFile = Join-Path $TestDrive 'one-record.xlsx'
            New-RecordsXlsx -Path $singleRecordFile -Records (New-SampleRecords -Count 1)
            $params.FormDataExcelFilePath = $singleRecordFile
            $params.MaxRetries = 3
            # Empty table (default) -> the single source row is created.

            $script:createCalls = 0
            Mock New-ServiceNowRecord {
                $script:createCalls++
                if ($script:createCalls -eq 1) { throw 'transient create error' }
            }

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke New-ServiceNowRecord -Exactly -Times 2   # one failure + one success
            Should -Invoke Start-Sleep -Exactly -Times 1
        }

        It 'throws CRITICAL FAILURE after exhausting MaxRetries on creation' {
            $singleRecordFile = Join-Path $TestDrive 'one-record.xlsx'
            New-RecordsXlsx -Path $singleRecordFile -Records (New-SampleRecords -Count 1)
            $params.FormDataExcelFilePath = $singleRecordFile
            $params.MaxRetries = 2
            Mock New-ServiceNowRecord { throw 'persistent create error' }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*CRITICAL FAILURE*'

            Should -Invoke New-ServiceNowRecord -Exactly -Times 2
            Should -Invoke Start-Sleep -Exactly -Times 1   # only between attempts, not after the final throw
        }

        It 'treats a failed deactivation as non-fatal and continues' {
            $params.MaxRetries = 3
            # One active straggler to deactivate; the two source rows are new.
            Mock Get-ServiceNowRecord {
                @( New-StaleTableRecord -SysId 'rec-stale' -Active $true )
            }
            Mock Update-ServiceNowRecord { throw 'permanent update error' }

            { & $ScriptPath @params } | Should -Not -Throw

            # Deactivation is retried up to MaxRetries (3 attempts, 2 sleeps) then
            # abandoned with a warning - the creates still happen.
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 3
            Should -Invoke Start-Sleep -Exactly -Times 2
            Should -Invoke New-ServiceNowRecord -Exactly -Times 2
        }

        It 'treats a failed reactivation as non-fatal and continues' {
            $params.MaxRetries = 3
            # Source row 1 matches an inactive row (reactivation target); source
            # row 2 has no match and must be created.
            Mock Get-ServiceNowRecord {
                @( New-MatchingTableRecord -Index 1 -SysId 'rec-1' -Active $false )
            }
            Mock Update-ServiceNowRecord { throw 'permanent update error' }

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke Update-ServiceNowRecord -Exactly -Times 3
            Should -Invoke Start-Sleep -Exactly -Times 2
            Should -Invoke New-ServiceNowRecord -Exactly -Times 1
        }
    }

    Context 'happy path' {
        It 'creates a session, then adds, reactivates and deactivates as needed' {
            $threeRecordFile = Join-Path $TestDrive 'three-records.xlsx'
            New-RecordsXlsx -Path $threeRecordFile -Records (New-SampleRecords -Count 3)
            $params.FormDataExcelFilePath = $threeRecordFile

            # Table state:
            #   rec-1   matches source 1, active   -> no change
            #   rec-2   matches source 2, inactive -> reactivate
            #   rec-st  not in source, active      -> deactivate
            # Source 3 has no matching row         -> create
            Mock Get-ServiceNowRecord {
                @(
                    New-MatchingTableRecord -Index 1 -SysId 'rec-1' -Active $true
                    New-MatchingTableRecord -Index 2 -SysId 'rec-2' -Active $false
                    New-StaleTableRecord -SysId 'rec-st' -Active $true
                )
            }

            & $ScriptPath @params

            Should -Invoke New-ServiceNowSession -Exactly -Times 1
            Should -Invoke Get-ServiceNowRecord -Exactly -Times 1
            Should -Invoke New-ServiceNowRecord -Exactly -Times 1
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 2
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 1 -ParameterFilter {
                $ID -eq 'rec-2' -and $Values.u_active -eq $true
            }
            Should -Invoke Update-ServiceNowRecord -Exactly -Times 1 -ParameterFilter {
                $ID -eq 'rec-st' -and $Values.u_active -eq $false
            }
        }
    }
}