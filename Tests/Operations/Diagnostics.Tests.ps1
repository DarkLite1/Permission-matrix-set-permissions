#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

<#
    Guards the telemetry pipeline added for run-over-run performance
    diagnostics.

    The regression these tests exist to prevent is quiet rather than loud: if a
    'Telemetry' record ever reaches $matrix.Check, nothing throws. Every clean
    settings row simply grows a notice card, the issue tally starts counting
    diagnostics as findings, and the summary mail fills with noise. That is the
    kind of breakage that ships.
#>

Describe 'Execution telemetry' {
    BeforeAll {
        <#
            Walk up from this file until the repository root is found, so the
            test keeps working wherever it is filed inside Tests\. The same
            idiom as DuplicatedHelpers.Tests.ps1.

            A fixed "$PSScriptRoot\..\..\.." would silently resolve to a
            DIFFERENT existing folder when the file is moved between
            Tests\Unit\Private and Tests\Operations, and the failure surfaces
            as 'ConvertTo-StructuredObjectHC is not recognized' rather than as
            a path problem — so it is worth the extra few lines.
        #>
        $script:repoRoot = $null
        $candidate = [System.IO.DirectoryInfo]::new($PSScriptRoot)

        while ($candidate) {
            $hasModules = Test-Path -LiteralPath (Join-Path $candidate.FullName 'Modules\PermissionMatrix') -PathType Container
            $hasScripts = Test-Path -LiteralPath (Join-Path $candidate.FullName 'Scripts\Operations') -PathType Container

            if ($hasModules -and $hasScripts) {
                $script:repoRoot = $candidate.FullName
                break
            }

            $candidate = $candidate.Parent
        }

        if (-not $script:repoRoot) {
            throw "Could not locate the repository root above '$PSScriptRoot'. Expected a folder holding both 'Modules\PermissionMatrix' and 'Scripts\Operations'."
        }

        $moduleRoot = Join-Path $script:repoRoot 'Modules\PermissionMatrix'

        $privateFolder = Join-Path $moduleRoot 'Private'
        $dotSourced = @(Get-ChildItem $privateFolder -Filter '*.ps1' -File)

        if ($dotSourced.Count -eq 0) {
            throw "No private functions found in '$privateFolder'."
        }

        $dotSourced | ForEach-Object { . $_.FullName }

        # Fail loudly and immediately if the functions under test did not load,
        # instead of letting every It report 'term is not recognized'.
        foreach ($required in @(
                'ConvertTo-StructuredObjectHC',
                'Write-MatrixDiagnosticsJsonHC',
                'Write-RunDiagnosticsJsonHC',
                'Write-RunPathDiagnosticsJsonHC',
                'Write-DiagnosticsFieldReferenceHC',
                'Get-DiagnosticsFieldReferenceHC',
                'Build-MatrixDetailCardHC'
            )) {
            if (-not (Get-Command $required -ErrorAction SilentlyContinue)) {
                throw "Required function '$required' was not loaded from '$privateFolder'."
            }
        }

        # Minimal stand-in for a matrix object as built by Import-MatrixFileHC.
        function New-TestMatrixObject {
            param(
                [string]$Path = 'E:\DEPARTMENTS\STAFF\SCM',
                [string]$ComputerName = 'BELSFFRAN0001',
                [string]$Action = 'Fix',
                [hashtable]$Telemetry,
                [int]$DurationSeconds = 90
            )

            [pscustomobject]@{
                ID                  = [guid]::NewGuid().ToString()
                Setting             = @{
                    Formatted = [pscustomobject]@{
                        ComputerName = $ComputerName
                        Path         = $Path
                        Action       = $Action
                    }
                }
                Check               = [System.Collections.Generic.List[pscustomobject]]::new()
                Matrix              = [System.Collections.Generic.List[pscustomobject]]::new()
                AdObjects           = @{}
                JobTime             = @{
                    Start    = Get-Date
                    End      = (Get-Date).AddSeconds($DurationSeconds)
                    Duration = New-TimeSpan -Seconds $DurationSeconds
                }
                Telemetry           = $Telemetry
                DiagnosticsFileName = $null
                FileContext         = [pscustomobject]@{
                    Item = [pscustomobject]@{ Name = 'BNL-MTX-STAFF-SCM.xlsx' }
                }
            }
        }

        function New-TestPathRows {
            param([long]$Big = 100000, [long]$Small = 20000)

            @(
                [ordered]@{ Path = 'E:\SCM\Archive'; ItemsWalked = $Big; AclReadMsPerItem = 1.44; IncorrectItems = 0; AceCountMean = 6.2; Walked = $true }
                [ordered]@{ Path = 'E:\SCM\Projects'; ItemsWalked = $Small; AclReadMsPerItem = 0.38; IncorrectItems = 2; AceCountMean = 6.1; Walked = $true }
                [ordered]@{ Path = 'E:\SCM\Ignored'; ItemsWalked = 0; AclReadMsPerItem = 0; IncorrectItems = 0; AceCountMean = 0; Walked = $false }
            )
        }

        function New-TestTelemetry {
            param([long]$ItemsWalked = 120000, [double]$AclReadMsPerItem = 0.42)

            [ordered]@{
                Paths            = (New-TestPathRows)
                Path             = 'E:\DEPARTMENTS\STAFF\SCM'
                Action           = 'Fix'
                ItemsWalked      = $ItemsWalked
                FoldersWalked    = 20000
                FilesWalked      = ($ItemsWalked - 20000)
                AclReads         = $ItemsWalked
                AclReadMsPerItem = $AclReadMsPerItem
                AclWrites        = 0
                IncorrectItems   = 0
                AceCountMean     = 6.5
                AceCountMax      = 11
            }
        }
    }

    Context 'Routing: telemetry is not a check' {
        <#
         Mirrors the split in Invoke-PermissionMatrixProcessHC. Kept as a
         behavioural test of the .Where(..., 'Split') contract rather than a
         call into the orchestrator, which would need a live PSSession.
        #>
        BeforeAll {
            $script:remoteOutput = @(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Type     = 'Warning'
                    Name     = 'Inherited permissions incorrect'
                    Value    = @('E:\a', 'E:\b')
                }
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Type     = 'Telemetry'
                    Name     = 'Execution telemetry'
                    Value    = (New-TestTelemetry)
                }
            )
        }

        It 'splits Telemetry records away from the check stream' {
            $structured = @($remoteOutput | ConvertTo-StructuredObjectHC)

            $telemetryEntries, $checkEntries = $structured.Where(
                { $_.Type -eq 'Telemetry' }, 'Split'
            )

            $telemetryEntries.Count | Should -Be 1
            $checkEntries.Count | Should -Be 1
            $checkEntries[0].Type | Should -Be 'Warning'
        }

        It 'leaves no Telemetry record in Check' {
            $matrix = New-TestMatrixObject
            $structured = @($remoteOutput | ConvertTo-StructuredObjectHC)

            $telemetryEntries, $checkEntries = $structured.Where(
                { $_.Type -eq 'Telemetry' }, 'Split'
            )
            foreach ($entry in $checkEntries) { $matrix.Check.Add($entry) }
            if ($telemetryEntries) {
                $matrix.Telemetry = $telemetryEntries[-1].Value
            }

            @($matrix.Check | Where-Object Type -EQ 'Telemetry').Count |
            Should -Be 0
            $matrix.Telemetry.ItemsWalked | Should -Be 120000
        }

        It 'does not make a clean row render as a notice row' {
            # Build-MatrixDetailCardHC counts anything that is neither
            # FatalError nor Warning as an info notice, which switches the card
            # out of compact mode. A row whose ONLY extra record is telemetry
            # must stay compact.
            $matrix = New-TestMatrixObject -Telemetry (New-TestTelemetry)

            $info = @($matrix.Check | Where-Object {
                    $_.Type -ne 'FatalError' -and $_.Type -ne 'Warning'
                }).Count

            $info | Should -Be 0
        }
    }

    Context 'Write-MatrixDiagnosticsJsonHC' {
        BeforeEach {
            $script:logFolder = Join-Path 'TestDrive:' ([guid]::NewGuid())
            New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
        }

        It 'writes the file and stamps DiagnosticsFileName' {
            $matrix = New-TestMatrixObject -Telemetry (New-TestTelemetry)

            Write-MatrixDiagnosticsJsonHC -Matrix $matrix -LogFolder $logFolder

            $matrix.DiagnosticsFileName |
            Should -Be "ID $($matrix.ID) - Diagnostics.json"

            $file = Join-Path $logFolder $matrix.DiagnosticsFileName
            Test-Path -LiteralPath $file | Should -BeTrue

            $json = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
            $json.Path | Should -Be 'E:\DEPARTMENTS\STAFF\SCM'
            $json.Telemetry.ItemsWalked | Should -Be 120000
            $json.Duration | Should -Be '00:01:30'
        }

        It 'writes nothing for a row that never executed' {
            # A skipped row has no telemetry, which is different from a row
            # that walked nothing. The report must not link to a missing file.
            $matrix = New-TestMatrixObject -Telemetry $null

            Write-MatrixDiagnosticsJsonHC -Matrix $matrix -LogFolder $logFolder

            $matrix.DiagnosticsFileName | Should -BeNullOrEmpty
            @(Get-ChildItem -Path $logFolder -File).Count | Should -Be 0
        }
    }

    Context 'Write-RunDiagnosticsJsonHC' {
        BeforeEach {
            $script:runFolder = Join-Path 'TestDrive:' ([guid]::NewGuid())
            New-Item -ItemType Directory -Path $runFolder -Force | Out-Null
        }

        It 'flattens telemetry onto one row per Settings row' {
            $matrices = @(
                New-TestMatrixObject -Path 'E:\A' -Telemetry (New-TestTelemetry -ItemsWalked 100)
                New-TestMatrixObject -Path 'E:\B' -Telemetry (New-TestTelemetry -ItemsWalked 200)
            )

            Write-RunDiagnosticsJsonHC `
                -Matrices $matrices `
                -LogFolder $runFolder `
                -RunStartTime (Get-Date '2026-08-16 22:00:04')

            $file = Join-Path $runFolder 'Diagnostics.json'
            Test-Path -LiteralPath $file | Should -BeTrue

            $rows = @(Get-Content -LiteralPath $file -Raw | ConvertFrom-Json)
            $rows.Count | Should -Be 2

            # Flat, not nested: every field must be a top-level column so the
            # file pipes straight into Group-Object / Export-Csv.
            $rows[0].PSObject.Properties.Name | Should -Contain 'ItemsWalked'
            $rows[0].PSObject.Properties.Name | Should -Contain 'DurationSeconds'
            $rows[0].PSObject.Properties.Name | Should -Contain 'RunStartTime'
            $rows[0].ItemsWalked | Should -Be 100
            $rows[1].ItemsWalked | Should -Be 200
        }

        It 'stays a JSON array for a single-row run' {
            # Without -AsArray a one-matrix run collapses to a bare object and
            # every consumer needs a special case.
            $matrices = @(New-TestMatrixObject -Telemetry (New-TestTelemetry))

            Write-RunDiagnosticsJsonHC -Matrices $matrices -LogFolder $runFolder

            $raw = Get-Content `
                -LiteralPath (Join-Path $runFolder 'Diagnostics.json') -Raw
            $raw.Trim().StartsWith('[') | Should -BeTrue
        }

        It 'writes nothing when no row produced telemetry' {
            $matrices = @(New-TestMatrixObject -Telemetry $null)

            Write-RunDiagnosticsJsonHC -Matrices $matrices -LogFolder $runFolder

            Test-Path -LiteralPath (Join-Path $runFolder 'Diagnostics.json') |
            Should -BeFalse
        }
    }

    Context 'Diagnostics.Fields.json' {
        <#
         A field reference that has drifted from the data is worse than no
         reference: it is confidently wrong, and the reader has no way to tell.
         These tests make drift a build failure rather than a discovery.
        #>
        BeforeAll {
            $script:Reference = Get-DiagnosticsFieldReferenceHC

            # The field names the remote script ACTUALLY emits, read from its
            # AST rather than from a duplicated list. Locates the telemetry
            # record by shape (a hashtable carrying Type/Name/Description/Value)
            # and takes the key names of the hashtable assigned to 'Value'.
            $setPermissions = Join-Path $script:repoRoot 'Scripts\Operations\SetPermissions.ps1'
            $ast = [System.Management.Automation.Language.Parser]::ParseFile(
                $setPermissions, [ref]$null, [ref]$null
            )

            $script:EmittedFields = @()

            $ast.FindAll(
                { $args[0] -is [System.Management.Automation.Language.HashtableAst] }, $true
            ) | ForEach-Object {
                $keys = @($_.KeyValuePairs | ForEach-Object { $_.Item1.Extent.Text.Trim("'`"") })

                if (
                    ($keys -contains 'Type') -and ($keys -contains 'Name') -and
                    ($keys -contains 'Value') -and ($keys -contains 'Description')
                ) {
                    $valuePair = $_.KeyValuePairs |
                    Where-Object { $_.Item1.Extent.Text.Trim("'`"") -eq 'Value' }

                    $inner = $valuePair.Item2.FindAll(
                        { $args[0] -is [System.Management.Automation.Language.HashtableAst] }, $true
                    ) | Select-Object -First 1

                    if ($inner) {
                        $script:EmittedFields = @(
                            $inner.KeyValuePairs |
                            ForEach-Object { $_.Item1.Extent.Text.Trim("'`"") }
                        )
                    }
                }
            }
        }

        It 'found the telemetry record in SetPermissions.ps1' {
            # Guards the guard: if the AST search stops matching, the two tests
            # below would pass vacuously against an empty list.
            $script:EmittedFields.Count | Should -BeGreaterThan 20
        }

        It 'documents every field the telemetry record emits' {
            $undocumented = @(
                $script:EmittedFields |
                Where-Object { $_ -notin @($script:Reference.TelemetryFields.Keys) }
            )

            $undocumented -join ', ' | Should -BeNullOrEmpty
        }

        It 'documents no field the telemetry record does not emit' {
            $stale = @(
                @($script:Reference.TelemetryFields.Keys) |
                Where-Object { $_ -notin $script:EmittedFields }
            )

            $stale -join ', ' | Should -BeNullOrEmpty
        }

        It 'gives every telemetry field a unit and a meaning' {
            foreach ($name in @($script:Reference.TelemetryFields.Keys)) {
                $entry = $script:Reference.TelemetryFields[$name]
                $entry.Unit | Should -Not -BeNullOrEmpty -Because "'$name' needs a unit"
                $entry.Meaning | Should -Not -BeNullOrEmpty -Because "'$name' needs a meaning"
            }
        }

        It 'documents the identity fields written by both writers' {
            # Top level of the per-row file, plus the two extra columns the
            # roll-up adds.
            foreach ($required in @(
                    'ID', 'MatrixFile', 'ComputerName', 'Path', 'Action',
                    'Start', 'End', 'Duration',
                    'RunStartTime', 'DurationSeconds'
                )) {
                @($script:Reference.RecordFields.Keys) |
                Should -Contain $required
            }
        }

        It 'writes valid, round-trippable JSON' {
            $folder = Join-Path 'TestDrive:' ([guid]::NewGuid())
            New-Item -ItemType Directory -Path $folder -Force | Out-Null

            Write-DiagnosticsFieldReferenceHC -LogFolder $folder

            $file = Join-Path $folder 'Diagnostics.Fields.json'
            Test-Path -LiteralPath $file | Should -BeTrue

            $parsed = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
            $parsed.About | Should -Not -BeNullOrEmpty
            $parsed.HowToRead | Should -Not -BeNullOrEmpty
            $parsed.Caveats.Count | Should -BeGreaterThan 0
            @($parsed.TelemetryFields.PSObject.Properties.Name).Count |
            Should -Be @($script:Reference.TelemetryFields.Keys).Count
        }
    }

    Context 'Write-RunPathDiagnosticsJsonHC' {
        BeforeEach {
            $script:pathFolder = Join-Path 'TestDrive:' ([guid]::NewGuid())
            New-Item -ItemType Directory -Path $pathFolder -Force | Out-Null
        }

        It 'writes one flat row per matrix folder' {
            $matrices = @(New-TestMatrixObject -Telemetry (New-TestTelemetry))

            Write-RunPathDiagnosticsJsonHC `
                -Matrices $matrices -LogFolder $pathFolder `
                -RunStartTime (Get-Date '2026-08-17 22:00:04')

            $rows = @(
                Get-Content -LiteralPath (Join-Path $pathFolder 'Diagnostics.Paths.json') -Raw |
                ConvertFrom-Json
            )

            $rows.Count | Should -Be 3
            $rows[0].PSObject.Properties.Name | Should -Contain 'Path'
            $rows[0].PSObject.Properties.Name | Should -Contain 'SettingPath'
            $rows[0].PSObject.Properties.Name | Should -Contain 'ItemsWalked'
        }

        It 'carries the owning ID on every row so it joins to Diagnostics.json' {
            # Without this the drill-down file cannot be tied back to the row it
            # decomposes, which is the only thing that makes two grains usable.
            $matrix = New-TestMatrixObject -Telemetry (New-TestTelemetry)

            Write-RunPathDiagnosticsJsonHC -Matrices @($matrix) -LogFolder $pathFolder

            $rows = @(
                Get-Content -LiteralPath (Join-Path $pathFolder 'Diagnostics.Paths.json') -Raw |
                ConvertFrom-Json
            )

            @($rows | Where-Object { $_.ID -ne $matrix.ID }).Count | Should -Be 0
        }

        It 'keeps the nested breakdown out of the flat settings-level roll-up' {
            # Diagnostics.json must stay CSV-convertible. An array in a column
            # would break that, and the two files hold different grains anyway.
            $matrix = New-TestMatrixObject -Telemetry (New-TestTelemetry)

            Write-RunDiagnosticsJsonHC -Matrices @($matrix) -LogFolder $pathFolder

            $rows = @(
                Get-Content -LiteralPath (Join-Path $pathFolder 'Diagnostics.json') -Raw |
                ConvertFrom-Json
            )

            $rows[0].PSObject.Properties.Name | Should -Not -Contain 'Paths'
        }

        It 'writes nothing when no row has a per-path breakdown' {
            $telemetry = New-TestTelemetry
            $telemetry.Remove('Paths')

            Write-RunPathDiagnosticsJsonHC `
                -Matrices @(New-TestMatrixObject -Telemetry $telemetry) `
                -LogFolder $pathFolder

            Test-Path -LiteralPath (Join-Path $pathFolder 'Diagnostics.Paths.json') |
            Should -BeFalse
        }
    }

    Context 'Build-MatrixDetailCardHC diagnostics chip' {
        It 'renders the chip when a diagnostics file was written' {
            $matrix = New-TestMatrixObject -Telemetry (New-TestTelemetry)
            $matrix.DiagnosticsFileName = 'ID abc - Diagnostics.json'

            $html = Build-MatrixDetailCardHC -MatrixItem $matrix

            $html | Should -Match 'Diagnostics'
            $html | Should -Match ([regex]::Escape('ID abc - Diagnostics.json'))

            # The '?' beside the chip must point at the field reference one
            # level up, in the run folder.
            $html | Should -Match ([regex]::Escape('../Diagnostics.Fields.json'))
        }

        It 'renders no chip when no diagnostics file exists' {
            $matrix = New-TestMatrixObject -Telemetry $null

            $html = Build-MatrixDetailCardHC -MatrixItem $matrix

            $html | Should -Not -Match 'Diagnostics'
        }
    }
}