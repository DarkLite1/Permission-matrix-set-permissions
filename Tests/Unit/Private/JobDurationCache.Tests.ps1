#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

<#
    Tests for Private\JobDurationCache.ps1

    The central promise of this file is that the cache is NEVER load-bearing.
    Most of what follows therefore checks failure modes rather than the happy
    path: a missing file, corrupt JSON, truncated JSON, an unexpected schema, a
    future version, unreadable entries. Every one of them has to degrade to
    "no ordering information" without throwing, because a run that dies while
    reading a performance hint would be far worse than a run that is slow.

    Note on assertions:
        $Cache.Count -eq 0 and 'key not present' are treated as the same thing
        by the caller, so several tests assert the estimate is MaxValue rather
        than inspecting the hashtable. That is the behaviour the queue ordering
        actually depends on.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    . "$moduleRoot\Private\JobDurationCache.ps1"

    function New-MatrixItem {
        <#
        .SYNOPSIS
            A minimal stand-in for one flattened Settings row.

        .DESCRIPTION
            Only the four fields the cache touches are present:
            Setting.Formatted.ComputerName / .Path / .Action and
            JobTime.Duration. Passing $null for Seconds produces a row with no
            duration, which is how a skipped or failed job looks.
        #>
        param(
            [string]$ComputerName = 'server01',
            [string]$Path = '\\server01\share\folder',
            [string]$Action = 'Fix',
            $Seconds = 100
        )

        $jobTime = @{}

        if ($null -ne $Seconds) {
            $jobTime.Duration = [timespan]::FromSeconds($Seconds)
        }

        [PSCustomObject]@{
            Setting = @{
                Formatted = [PSCustomObject]@{
                    ComputerName = $ComputerName
                    Path         = $Path
                    Action       = $Action
                }
            }
            JobTime = $jobTime
        }
    }

    function Get-CacheFileContent {
        param([Parameter(Mandatory)][string]$LogFolder)
        Get-Content -LiteralPath (Join-Path $LogFolder 'JobDurations.json') -Raw |
            ConvertFrom-Json
    }
}

Describe 'Get-JobDurationCacheKeyHC' {
    It 'builds a lowercase pipe-joined triplet' {
        Get-JobDurationCacheKeyHC -ComputerName 'SERVER01' -Path '\\SERVER01\Share' -Action 'Fix' |
            Should -BeExactly 'server01|\\server01\share|fix'
    }

    It 'treats a trailing <Separator> as the same job' -TestCases @(
        @{ Separator = 'backslash'; Path = '\\server01\share\' }
        @{ Separator = 'forward slash'; Path = '\\server01\share/' }
    ) {
        param($Path)

        $withSeparator = Get-JobDurationCacheKeyHC -ComputerName 'server01' -Path $Path -Action 'Fix'
        $without = Get-JobDurationCacheKeyHC -ComputerName 'server01' -Path '\\server01\share' -Action 'Fix'

        $withSeparator | Should -BeExactly $without
    }

    It 'ignores surrounding whitespace' {
        Get-JobDurationCacheKeyHC -ComputerName '  server01 ' -Path ' \\server01\share ' -Action ' Fix ' |
            Should -BeExactly 'server01|\\server01\share|fix'
    }

    It 'distinguishes jobs that differ only by <Field>' -TestCases @(
        @{ Field = 'computer'; ComputerName = 'server02'; Path = '\\server01\share'; Action = 'Fix' }
        @{ Field = 'path'; ComputerName = 'server01'; Path = '\\server01\other'; Action = 'Fix' }
        @{ Field = 'action'; ComputerName = 'server01'; Path = '\\server01\share'; Action = 'Check' }
    ) {
        param($ComputerName, $Path, $Action)

        # One matrix file can target several computers and paths, so the key has
        # to be the whole triplet rather than the file name.
        $baseline = Get-JobDurationCacheKeyHC -ComputerName 'server01' -Path '\\server01\share' -Action 'Fix'
        $variant = Get-JobDurationCacheKeyHC -ComputerName $ComputerName -Path $Path -Action $Action

        $variant | Should -Not -BeExactly $baseline
    }

    It 'tolerates null input without throwing' {
        { Get-JobDurationCacheKeyHC -ComputerName $null -Path $null -Action $null } |
            Should -Not -Throw
    }
}

Describe 'Get-JobDurationCacheHC' {
    Context 'nothing usable to read' {
        It 'returns empty when the log folder is <Description>' -TestCases @(
            @{ Description = 'null'; LogFolder = $null }
            @{ Description = 'an empty string'; LogFolder = '' }
            @{ Description = 'whitespace'; LogFolder = '   ' }
        ) {
            param($LogFolder)

            $cache = Get-JobDurationCacheHC -LogFolder $LogFolder

            $cache | Should -BeOfType [hashtable]
            $cache.Count | Should -Be 0
        }

        It 'returns empty when the log folder does not exist' {
            $cache = Get-JobDurationCacheHC -LogFolder (Join-Path $TestDrive 'no-such-folder')

            $cache.Count | Should -Be 0
        }

        It 'returns empty on the first ever run, when no cache file exists yet' {
            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            $cache.Count | Should -Be 0
        }
    }

    Context 'a valid cache file' {
        BeforeEach {
            $cacheFile = Join-Path $TestDrive 'JobDurations.json'

            @{
                Version = 1
                Jobs    = @{
                    'server01|\\server01\share|fix'  = @{
                        TotalSeconds = 1465.32
                        LastSeen = '2026-07-27T02:14:33.0000000Z'
                    }
                    'server02|\\server02\other|fix' = @{
                        TotalSeconds = 12
                        LastSeen = '2026-07-27T02:14:33.0000000Z'
                    }
                }
            } | ConvertTo-Json -Depth 5 |
                Set-Content -LiteralPath $cacheFile -Encoding UTF8
        }

        AfterEach {
            Remove-Item (Join-Path $TestDrive '*') -Recurse -Force -ErrorAction Ignore
        }

        It 'reads every entry' {
            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            $cache.Count | Should -Be 2
        }

        It 'preserves fractional seconds' {
            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            $cache['server01|\\server01\share|fix'] | Should -Be 1465.32
        }

        It 'returns a value usable as a sort key' {
            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            $cache['server01|\\server01\share|fix'] |
                Should -BeGreaterThan $cache['server02|\\server02\other|fix']
        }
    }

    Context 'damaged or unexpected files - must degrade, never throw' {
        AfterEach {
            Remove-Item (Join-Path $TestDrive '*') -Recurse -Force -ErrorAction Ignore
        }

        It 'returns empty for <Description>' -TestCases @(
            @{ Description = 'malformed JSON'; Content = '{ this is not json' }
            @{ Description = 'JSON truncated mid-write'; Content = '{ "Version": 1, "Jobs": { "a|b|c": { "Tot' }
            @{ Description = 'an empty file'; Content = '' }
            @{ Description = 'whitespace only'; Content = "   `n  " }
            @{ Description = 'a JSON array instead of an object'; Content = '[1,2,3]' }
            @{ Description = 'valid JSON with no Jobs property'; Content = '{ "Version": 1 }' }
            @{ Description = 'a future schema version'; Content = '{ "Version": 99, "Jobs": { "a|b|c": { "TotalSeconds": 5 } } }' }
        ) {
            param($Content)

            Set-Content -LiteralPath (Join-Path $TestDrive 'JobDurations.json') `
                -Value $Content -Encoding UTF8

            # Two calls on purpose. Should -Not -Throw runs its scriptblock in a
            # child scope, so an assignment made inside it never reaches this
            # one: the variable would still be $null here and the type
            # assertion would fail even though the function behaved correctly.
            # Reading the cache is cheap, so assert the two properties
            # separately rather than smuggling a value out of the scriptblock.
            { Get-JobDurationCacheHC -LogFolder $TestDrive } | Should -Not -Throw

            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            $cache | Should -BeOfType [hashtable]
            $cache.Count | Should -Be 0
        }

        It 'keeps the good entries when one record is malformed' {
            # One bad record must not cost the whole optimisation.
            $content = @'
{
  "Version": 1,
  "Jobs": {
    "good|path|fix": { "TotalSeconds": 500, "LastSeen": "2026-07-27T02:00:00.0000000Z" },
    "bad|path|fix": { "TotalSeconds": "not-a-number", "LastSeen": "2026-07-27T02:00:00.0000000Z" },
    "alsogood|path|fix": { "TotalSeconds": 20, "LastSeen": "2026-07-27T02:00:00.0000000Z" }
  }
}
'@
            Set-Content -LiteralPath (Join-Path $TestDrive 'JobDurations.json') `
                -Value $content -Encoding UTF8

            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            $cache.Count | Should -Be 2
            $cache.ContainsKey('good|path|fix') | Should -BeTrue
            $cache.ContainsKey('bad|path|fix') | Should -BeFalse
        }

        It 'rejects a <Description> duration' -TestCases @(
            @{ Description = 'zero'; Value = '0' }
            @{ Description = 'negative'; Value = '-5' }
        ) {
            param($Value)

            # A non-positive duration is meaningless and would corrupt the
            # ordering rather than improve it.
            $content = '{ "Version": 1, "Jobs": { "a|b|c": { "TotalSeconds": ' + $Value + ', "LastSeen": "2026-07-27T02:00:00.0000000Z" } } }'

            Set-Content -LiteralPath (Join-Path $TestDrive 'JobDurations.json') `
                -Value $content -Encoding UTF8

            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            $cache.Count | Should -Be 0
        }
    }
}

Describe 'Get-JobDurationEstimateHC' {
    It 'returns the cached duration when the job is known' {
        $cache = @{ 'server01|\\server01\share|fix' = 1465.32 }

        Get-JobDurationEstimateHC -Cache $cache -ComputerName 'server01' `
            -Path '\\server01\share' -Action 'Fix' |
            Should -Be 1465.32
    }

    It 'sorts an unknown job first by returning MaxValue' {
        # Deliberate: an unknown job may be trivial or enormous. Starting a large
        # one early saves the entire ordering benefit; starting a small one early
        # costs almost nothing.
        $cache = @{ 'other|path|fix' = 10 }

        Get-JobDurationEstimateHC -Cache $cache -ComputerName 'server01' `
            -Path '\\server01\share' -Action 'Fix' |
            Should -Be ([double]::MaxValue)
    }

    It 'returns MaxValue for <Description>, so a first run keeps its original order' -TestCases @(
        @{ Description = 'an empty cache'; Cache = @{} }
        @{ Description = 'a null cache'; Cache = $null }
    ) {
        param($Cache)

        Get-JobDurationEstimateHC -Cache $Cache -ComputerName 'server01' `
            -Path '\\server01\share' -Action 'Fix' |
            Should -Be ([double]::MaxValue)
    }

    It 'matches a cached job regardless of casing or trailing separator' {
        $cache = @{ 'server01|\\server01\share|fix' = 42 }

        Get-JobDurationEstimateHC -Cache $cache -ComputerName 'SERVER01' `
            -Path '\\SERVER01\Share\' -Action 'FIX' |
            Should -Be 42
    }
}

Describe 'Save-JobDurationCacheHC' {
    AfterEach {
        Remove-Item (Join-Path $TestDrive '*') -Recurse -Force -ErrorAction Ignore
    }

    Context 'no-ops that must stay silent' {
        It 'does nothing when <Description>' -TestCases @(
            @{ Description = 'no log folder is given'; LogFolder = ''; Matrices = @() }
            @{ Description = 'the log folder is null'; LogFolder = $null; Matrices = @() }
        ) {
            param($LogFolder, $Matrices)

            { Save-JobDurationCacheHC -LogFolder $LogFolder -Matrices $Matrices } |
                Should -Not -Throw
        }

        It 'does not create the log folder when it is missing' {
            $missing = Join-Path $TestDrive 'not-created'

            { Save-JobDurationCacheHC -LogFolder $missing -Matrices @((New-MatrixItem)) } |
                Should -Not -Throw

            Test-Path $missing | Should -BeFalse
        }

        It 'writes nothing when there are no matrices' {
            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices @()

            Test-Path (Join-Path $TestDrive 'JobDurations.json') | Should -BeFalse
        }

        It 'writes nothing when no job completed with a duration' {
            # Every job failed or was skipped: there is nothing to learn, and the
            # previous file must be left exactly as it was.
            Save-JobDurationCacheHC -LogFolder $TestDrive `
                -Matrices @((New-MatrixItem -Seconds $null))

            Test-Path (Join-Path $TestDrive 'JobDurations.json') | Should -BeFalse
        }
    }

    Context 'writing observations' {
        It 'creates the file with the expected shape' {
            Save-JobDurationCacheHC -LogFolder $TestDrive `
                -Matrices @((New-MatrixItem -Seconds 250))

            $json = Get-CacheFileContent -LogFolder $TestDrive

            $json.Version | Should -Be 1
            $json.Updated | Should -Not -BeNullOrEmpty
            $json.Jobs.'server01|\\server01\share\folder|fix'.TotalSeconds | Should -Be 250
        }

        It 'records one entry per job, not per matrix file' {
            $matrices = @(
                New-MatrixItem -ComputerName 'server01' -Path '\\server01\a' -Seconds 10
                New-MatrixItem -ComputerName 'server01' -Path '\\server01\b' -Seconds 20
                New-MatrixItem -ComputerName 'server02' -Path '\\server02\a' -Seconds 30
            )

            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices $matrices

            $json = Get-CacheFileContent -LogFolder $TestDrive

            @($json.Jobs.PSObject.Properties).Count | Should -Be 3
        }

        It 'stores the WHOLE duration, not the seconds component' {
            # The bug this guards: [timespan].Seconds is the 0-59 component,
            # while .TotalSeconds is the whole duration. A 24m25s job stored via
            # .Seconds would be recorded as 25 and sort last - the exact
            # inversion the cache exists to prevent. Anything over a minute
            # distinguishes the two; 1465s is the real duration that motivated
            # this feature.
            Save-JobDurationCacheHC -LogFolder $TestDrive `
                -Matrices @((New-MatrixItem -Seconds 1465))

            $json = Get-CacheFileContent -LogFolder $TestDrive

            $json.Jobs.'server01|\\server01\share\folder|fix'.TotalSeconds |
                Should -Be 1465

            # 1465s is 24m25s, so the component value would be 25.
            $json.Jobs.'server01|\\server01\share\folder|fix'.TotalSeconds |
                Should -Not -Be 25
        }

        It 'rounds the duration to two decimals' {
            $matrices = @((New-MatrixItem -Seconds 123.456789))

            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices $matrices

            $json = Get-CacheFileContent -LogFolder $TestDrive

            $json.Jobs.'server01|\\server01\share\folder|fix'.TotalSeconds | Should -Be 123.46
        }

        It 'stores a round-trippable UTC timestamp' {
            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices @((New-MatrixItem))

            $json = Get-CacheFileContent -LogFolder $TestDrive
            $lastSeen = $json.Jobs.'server01|\\server01\share\folder|fix'.LastSeen

            { [datetime]::Parse($lastSeen, [cultureinfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::RoundtripKind) } |
                Should -Not -Throw
        }

        It 'writes keys in sorted order so run-to-run diffs are readable' {
            $matrices = @(
                New-MatrixItem -ComputerName 'zebra' -Path '\\zebra\a' -Seconds 10
                New-MatrixItem -ComputerName 'alpha' -Path '\\alpha\a' -Seconds 20
            )

            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices $matrices

            $names = @((Get-CacheFileContent -LogFolder $TestDrive).Jobs.PSObject.Properties.Name)

            $names[0] | Should -BeLike 'alpha*'
        }

        It 'leaves no temporary file behind' {
            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices @((New-MatrixItem))

            @(Get-ChildItem -Path $TestDrive -Filter '*.tmp').Count | Should -Be 0
        }
    }

    Context 'merging with what is already stored' {
        BeforeEach {
            @{
                Version = 1
                Jobs    = @{
                    'server01|\\server01\share\folder|fix' = @{
                        TotalSeconds = 999
                        LastSeen = (Get-Date).ToUniversalTime().AddDays(-1).ToString('o')
                    }
                    'absent|\\absent\share|fix'            = @{
                        TotalSeconds = 777
                        LastSeen = (Get-Date).ToUniversalTime().AddDays(-2).ToString('o')
                    }
                }
            } | ConvertTo-Json -Depth 5 |
                Set-Content -LiteralPath (Join-Path $TestDrive 'JobDurations.json') -Encoding UTF8
        }

        It 'overwrites the duration of a job that ran again' {
            Save-JobDurationCacheHC -LogFolder $TestDrive `
                -Matrices @((New-MatrixItem -Seconds 250))

            $json = Get-CacheFileContent -LogFolder $TestDrive

            $json.Jobs.'server01|\\server01\share\folder|fix'.TotalSeconds | Should -Be 250
        }

        It 'keeps a job that did not run this time' {
            # A matrix file disabled or missing tonight keeps its remembered
            # duration for whenever it comes back.
            Save-JobDurationCacheHC -LogFolder $TestDrive `
                -Matrices @((New-MatrixItem -Seconds 250))

            $json = Get-CacheFileContent -LogFolder $TestDrive

            $json.Jobs.'absent|\\absent\share|fix'.TotalSeconds | Should -Be 777
        }

        It 'does not overwrite a remembered duration when the job failed' {
            # A failed job has no duration. Writing zero would demote a
            # genuinely expensive job in tomorrow's ordering.
            $matrices = @(
                New-MatrixItem -Seconds $null
                New-MatrixItem -ComputerName 'other' -Path '\\other\a' -Seconds 5
            )

            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices $matrices

            $json = Get-CacheFileContent -LogFolder $TestDrive

            $json.Jobs.'server01|\\server01\share\folder|fix'.TotalSeconds | Should -Be 999
        }

        It 'prunes entries not seen within the retention window' {
            Save-JobDurationCacheHC -LogFolder $TestDrive `
                -Matrices @((New-MatrixItem -Seconds 250)) `
                -RetentionDays 1

            $json = Get-CacheFileContent -LogFolder $TestDrive

            # Two days old, so it falls outside a one-day window.
            $json.Jobs.PSObject.Properties.Name | Should -Not -Contain 'absent|\\absent\share|fix'
            # Refreshed this run, so it stays.
            $json.Jobs.PSObject.Properties.Name | Should -Contain 'server01|\\server01\share\folder|fix'
        }

        It 'starts fresh when the existing file cannot be parsed' {
            Set-Content -LiteralPath (Join-Path $TestDrive 'JobDurations.json') `
                -Value '{ corrupt' -Encoding UTF8

            { Save-JobDurationCacheHC -LogFolder $TestDrive `
                    -Matrices @((New-MatrixItem -Seconds 250)) } | Should -Not -Throw

            $json = Get-CacheFileContent -LogFolder $TestDrive

            @($json.Jobs.PSObject.Properties).Count | Should -Be 1
            $json.Jobs.'server01|\\server01\share\folder|fix'.TotalSeconds | Should -Be 250
        }
    }

    Context 'a full round trip' {
        It 'reads back exactly what it wrote' {
            $matrices = @(
                New-MatrixItem -ComputerName 'busy' -Path '\\busy\share' -Seconds 1465
                New-MatrixItem -ComputerName 'idle' -Path '\\idle\share' -Seconds 12
            )

            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices $matrices

            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            $cache.Count | Should -Be 2

            Get-JobDurationEstimateHC -Cache $cache -ComputerName 'busy' `
                -Path '\\busy\share' -Action 'Fix' | Should -Be 1465
        }

        It 'orders the expensive job first, which is the whole point' {
            $matrices = @(
                New-MatrixItem -ComputerName 'srv' -Path '\\srv\small' -Seconds 12
                New-MatrixItem -ComputerName 'srv' -Path '\\srv\huge' -Seconds 1465
                New-MatrixItem -ComputerName 'srv' -Path '\\srv\medium' -Seconds 300
            )

            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices $matrices

            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            # Deliberately fed in cheapest-first to prove the sort does the work.
            $sorted = @(
                @('\\srv\small', '\\srv\medium', '\\srv\huge') |
                Sort-Object -Property {
                    Get-JobDurationEstimateHC -Cache $cache -ComputerName 'srv' `
                        -Path $_ -Action 'Fix'
                } -Descending
            )

            $sorted[0] | Should -BeExactly '\\srv\huge'
            $sorted[-1] | Should -BeExactly '\\srv\small'
        }

        It 'puts an unknown job ahead of every known one' {
            $matrices = @(
                New-MatrixItem -ComputerName 'srv' -Path '\\srv\known' -Seconds 1465
            )

            Save-JobDurationCacheHC -LogFolder $TestDrive -Matrices $matrices

            $cache = Get-JobDurationCacheHC -LogFolder $TestDrive

            $sorted = @(
                @('\\srv\known', '\\srv\brand-new') |
                Sort-Object -Property {
                    Get-JobDurationEstimateHC -Cache $cache -ComputerName 'srv' `
                        -Path $_ -Action 'Fix'
                } -Descending
            )

            $sorted[0] | Should -BeExactly '\\srv\brand-new'
        }
    }
}