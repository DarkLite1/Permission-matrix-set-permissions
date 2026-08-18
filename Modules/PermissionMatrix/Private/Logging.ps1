function Remove-OldLogsHC {
    <#
    .SYNOPSIS
        Purges old log files and the directories left empty behind them.

    .DESCRIPTION
        Deletes files older than the retention threshold, then walks the
        directory tree bottom-up (sorted descending) removing any subdirectory
        that is now empty.

    .NOTES
        - Age is measured on CreationTime, NOT LastWriteTime. A file that is
          still being appended to is deleted once it is old enough.
        - RetentionDays of 0 or less disables the cleanup entirely.
        - Deletion failures (a log open in another process, access denied) are
          appended to SystemErrors as WARNINGS and never throw, so cleanup
          problems cannot crash the orchestrator.

    .EXAMPLE
        $sysErrors = [System.Collections.Generic.List[pscustomobject]]::new()

        Remove-OldLogsHC `
            -LogFolder 'C:\MatrixLogs' `
            -RetentionDays 30 `
            -SystemErrors ([ref]$sysErrors)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$LogFolder,
        [int]$RetentionDays,
        [Parameter(Mandatory)] [ref]$SystemErrors
    )

    # Disabled or folder missing → nothing to do
    if ($RetentionDays -le 0 -or -not $LogFolder) { return }

    try {
        if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return }

        $cutoff = (Get-Date).AddDays(-$RetentionDays)

        # --- 1. Delete old files ---
        Get-ChildItem -LiteralPath $LogFolder -Recurse -File -ErrorAction Stop |
        Where-Object { $_.CreationTime -lt $cutoff } |
        ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
            }
            catch {
                Add-ErrorHC `
                    -Type 'Warning' `
                    -Name 'Log cleanup failed' `
                    -Message "Failed to delete log file '$($_.FullName)': $_" `
                    -Category 'Logging' `
                    -SystemErrors $SystemErrors
            }
        }

        # --- 2. Empty folder cleanup (bottom-up) ---
        Get-ChildItem -LiteralPath $LogFolder -Recurse -Directory -ErrorAction Stop |
        Sort-Object FullName -Descending |
        ForEach-Object {
            if (-not $_.GetFileSystemInfos().Count) {
                try {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                }
                catch {
                    Add-ErrorHC `
                        -Type 'Warning' `
                        -Name 'Log cleanup failed' `
                        -Message "Failed to remove empty folder '$($_.FullName)': $_" `
                        -Category 'Logging' `
                        -SystemErrors $SystemErrors
                }
            }
        }
    }
    catch {
        Add-ErrorHC `
            -Type 'Warning' `
            -Name 'Log cleanup failed' `
            -Message "General log cleanup failure: $_" `
            -Category 'Logging' `
            -SystemErrors $SystemErrors
    }
}

function Write-CheckDetailJsonHC {
    <#
    .SYNOPSIS
        Writes a single check object to its detailed JSON log file.

    .DESCRIPTION
        Stamps the check with 'JsonFileName' and 'JsonFilePath', then, when the
        check carries a 'Value', serializes a copy (excluding those two link
        properties) to disk as JSON.

    .NOTES
        - The passed-in check object is MUTATED in place.
        - ErrorRecord/Exception values are rendered to their string form first,
          to avoid serialization depth failures.
        - When the check has no 'Value', or serialization fails, both link
          properties are reset to $null so downstream reporting never links to a
          file that was not written. A failure is also appended to the check's
          'Description'.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Check,
        [Parameter(Mandatory)] [string]$JsonFileName,
        [Parameter(Mandatory)] [string]$LogFolder
    )

    $Check | Add-Member -NotePropertyMembers @{
        JsonFileName = $JsonFileName
        JsonFilePath = Join-Path -Path $LogFolder -ChildPath $JsonFileName
    } -Force

    if (-not $Check.Value) {
        $Check.JsonFileName = $null
        $Check.JsonFilePath = $null
        return
    }

    try {
        $cForJson = $Check | Select-Object -ExcludeProperty JsonFilePath, JsonFileName

        if (
            $cForJson.Value -is [System.Management.Automation.ErrorRecord] -or
            $cForJson.Value -is [Exception]
        ) {
            $cForJson.Value = ($cForJson.Value | Out-String).Trim()
        }

        $cForJson | ConvertTo-Json -Depth 10 |
        Out-File -FilePath $Check.JsonFilePath -Encoding UTF8 -Force
    }
    catch {
        $Check.Description += "[Detailed JSON log failed to generate: $($_)]"
        $Check.JsonFileName = $null
        $Check.JsonFilePath = $null
    }
}

function Write-MatrixDiagnosticsJsonHC {
    <#
    .SYNOPSIS
        Writes one Settings row's telemetry to its own JSON file.

    .DESCRIPTION
        Produces 'ID <guid> - Diagnostics.json' next to the row's detail files
        and stamps 'DiagnosticsFileName' on the matrix object so the execution
        report can render a link to it.

        This is the DRILL-DOWN artifact: everything known about one path in one
        run. For the cross-run view, see Write-RunDiagnosticsJsonHC.

    .NOTES
        - The passed-in matrix object is MUTATED in place.
        - A row without telemetry (never executed, or a FatalError before the
          remote script ran) writes nothing and leaves the property $null, so
          the report never links to a file that is not there.
        - Failures are swallowed. Diagnostics that cannot be written are not
          worth failing a run over, and not worth a line in the summary mail.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Matrix,
        [Parameter(Mandatory)] [string]$LogFolder
    )

    if (-not $Matrix.Telemetry) { return }

    $fileName = "ID $($Matrix.ID) - Diagnostics.json"

    try {
        $duration = if ($Matrix.JobTime.Duration) {
            '{0:00}:{1:00}:{2:00}' -f `
                $Matrix.JobTime.Duration.Hours,
            $Matrix.JobTime.Duration.Minutes,
            $Matrix.JobTime.Duration.Seconds
        }
        else { $null }

        # The identity fields are repeated inside the file on purpose. The
        # file is meant to be readable on its own, and to survive being copied
        # out of its folder into a ticket or a mail.
        $payload = [ordered]@{
            ID           = $Matrix.ID
            MatrixFile   = $Matrix.FileContext.Item.Name
            ComputerName = $Matrix.Setting.Formatted.ComputerName
            Path         = $Matrix.Setting.Formatted.Path
            Action       = $Matrix.Setting.Formatted.Action
            Start        = $Matrix.JobTime.Start
            End          = $Matrix.JobTime.End
            Duration     = $duration
            Telemetry    = $Matrix.Telemetry
        }

        $payload | ConvertTo-Json -Depth 10 |
        Out-File `
            -FilePath (Join-Path -Path $LogFolder -ChildPath $fileName) `
            -Encoding UTF8 -Force

        $Matrix.DiagnosticsFileName = $fileName
    }
    catch {
        Write-Verbose "Failed writing diagnostics JSON for ID '$($Matrix.ID)': $_"
        $Matrix.DiagnosticsFileName = $null
    }
}

function Write-RunPathDiagnosticsJsonHC {
    <#
    .SYNOPSIS
        Writes 'Diagnostics.Paths.json': one flat row per matrix folder, across
        the whole run.

    .DESCRIPTION
        The drill-down companion to 'Diagnostics.json'.

        'Diagnostics.json' answers "which Settings row got slower". This file
        answers "and where inside it", which is the question that actually leads
        somewhere: a Settings row covering 67 matrix folders can double because
        one child tree grew, and the row-level total cannot tell you which.

        WHY A SEPARATE FILE RATHER THAN MORE COLUMNS
        The two files hold different GRAINS. Mixing one-row-per-setting and
        one-row-per-folder in a single table would break every aggregate taken
        over it — sum a column and the folders get counted twice, once on their
        own row and once inside the setting total. Separate files keep both
        tables individually summable, and 'ID' joins them.

        WHY THE SETTINGS-LEVEL FILE KEEPS ITS SHAPE
        Anything already written against 'Diagnostics.json' keeps working. This
        is additive.

    .NOTES
        Failures are swallowed, as with the other diagnostics writers.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Matrices,
        [Parameter(Mandatory)] [string]$LogFolder,
        [Parameter()] [datetime]$RunStartTime
    )

    try {
        $rows = foreach ($m in $Matrices) {
            if (-not $m.Telemetry) { continue }
            if (-not $m.Telemetry.Paths) { continue }

            foreach ($pathRow in $m.Telemetry.Paths) {
                $row = [ordered]@{
                    RunStartTime = $(
                        if ($RunStartTime) {
                            $RunStartTime.ToString('yyyy-MM-ddTHH:mm:ss')
                        }
                        else { $null }
                    )
                    # Joins back to Diagnostics.json.
                    ID           = $m.ID
                    MatrixFile   = $m.FileContext.Item.Name
                    ComputerName = $m.Telemetry.ComputerName
                    Action       = $m.Telemetry.Action
                    # The Settings-row path, so a folder can be traced to the
                    # row that owns it without a lookup.
                    SettingPath  = $m.Telemetry.Path
                }

                foreach ($field in $pathRow.GetEnumerator()) {
                    $row[$field.Key] = $field.Value
                }

                [PSCustomObject]$row
            }
        }

        $rows = @($rows)

        if ($rows.Count -eq 0) { return }

        $rows | ConvertTo-Json -Depth 10 -AsArray |
        Out-File `
            -FilePath (Join-Path -Path $LogFolder -ChildPath 'Diagnostics.Paths.json') `
            -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Failed writing the per-path diagnostics JSON: $_"
    }
}

function Write-RunDiagnosticsJsonHC {
    <#
    .SYNOPSIS
        Writes one flat array holding every Settings row's telemetry for the
        whole run.

    .DESCRIPTION
        Produces 'Diagnostics.json' in the dated run folder.

        This is the TREND artifact, and it is the one that answers 'is this
        getting worse?'. The per-row files are fine for inspecting a single
        path, but comparing five nights across 101 settings means opening 505
        files. One array per run makes that a one-liner:

            Get-ChildItem '<log root>\*\Diagnostics.json' |
                ForEach-Object {
                    $run = $_.Directory.Name
                    Get-Content $_ -Raw | ConvertFrom-Json |
                    Where-Object Path -eq 'E:\DEPARTMENTS\STAFF\SCM' |
                    Select-Object @{n='Run';e={$run}},
                                  DurationSeconds, ItemsWalked,
                                  AclReadMsPerItem, AceCountMean
                }

        Duration is repeated here as a plain number, not a formatted string,
        because this file is meant to be sorted and charted rather than read.

    .NOTES
        - Failures are swallowed for the same reason as the per-row file.
        - Rows without telemetry are skipped rather than emitted as nulls, so
          the array holds only rows that actually executed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Matrices,
        [Parameter(Mandatory)] [string]$LogFolder,
        [Parameter()] [datetime]$RunStartTime
    )

    try {
        $rows = foreach ($m in $Matrices) {
            if (-not $m.Telemetry) { continue }

            $row = [ordered]@{
                RunStartTime    = $(
                    if ($RunStartTime) {
                        $RunStartTime.ToString('yyyy-MM-ddTHH:mm:ss')
                    }
                    else { $null }
                )
                ID              = $m.ID
                MatrixFile      = $m.FileContext.Item.Name
                DurationSeconds = $(
                    if ($m.JobTime.Duration) {
                        [math]::Round($m.JobTime.Duration.TotalSeconds, 1)
                    }
                    else { $null }
                )
            }

            # Flatten the telemetry onto the row so every field is a
            # top-level column. Nested objects are what make a JSON awkward to
            # pipe into Group-Object / Export-Csv, and trending is the whole
            # purpose of this file.
            foreach ($t in $m.Telemetry.GetEnumerator()) {
                # 'Paths' is the nested per-folder breakdown. It belongs in
                # Diagnostics.Paths.json, not here: embedding an array in a
                # column would stop this file converting cleanly to CSV, which
                # is the whole reason it is flat.
                if ($t.Key -eq 'Paths') { continue }

                $row[$t.Key] = $t.Value
            }

            [PSCustomObject]$row
        }

        $rows = @($rows)

        if ($rows.Count -eq 0) { return }

        # -AsArray keeps a single-row run as a JSON array rather than
        # collapsing to a bare object, so consumers never need to special-case
        # the one-matrix run.
        $rows | ConvertTo-Json -Depth 10 -AsArray |
        Out-File `
            -FilePath (Join-Path -Path $LogFolder -ChildPath 'Diagnostics.json') `
            -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Failed writing the run diagnostics JSON: $_"
    }
}

function Get-DiagnosticsFieldReferenceHC {
    <#
    .SYNOPSIS
        The documented meaning of every field written to the diagnostics files.

    .DESCRIPTION
        Returns an ordered structure describing the artifacts, how to read them,
        and every field they contain.

        WHY THIS IS A SEPARATE FILE AND NOT COMMENTS IN THE DATA
        'Diagnostics.json' is built to be piped straight into Group-Object,
        Export-Csv or a chart. Interleaving description keys with the numbers
        would break that: every consumer would have to filter documentation out
        of its own data. So the data stays clean and the explanation sits beside
        it in its own file, written once per run.

        WHY IT IS REGENERATED EVERY RUN RATHER THAN CHECKED IN
        Log folders get zipped and mailed around, and by the time someone reads
        one they usually do not have the repository open. A run folder that
        explains itself is worth the few KB.

    .NOTES
        Tests\Unit\Private\Diagnostics.Tests.ps1 asserts that this reference and
        the telemetry record emitted by SetPermissions.ps1 describe exactly the
        same field names. A field added to one and not the other fails the
        build, because a stale field reference is worse than none at all.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param()

    return [ordered]@{
        About       = [ordered]@{
            Purpose  = 'Volume and cost counters per Settings row, so a change in run time can be attributed to a change in the amount of data, the cost of each storage operation, or neither.'
            Files    = [ordered]@{
                'Diagnostics.json'                  = 'This folder. One flat row per Settings row for the whole run. Every field is a top-level column, so the file sorts, groups and charts directly. Use this one to compare runs.'
                'ID <guid> - Diagnostics.json'      = 'Inside each matrix subfolder. One Settings row in full, with its identity and timings. Use this one to inspect a single path.'
                'Diagnostics.Paths.json'            = 'This folder. One flat row per MATRIX FOLDER, for the whole run. Answers "where inside the Settings row did it happen": a row covering dozens of folders can double because one child tree grew. Join to Diagnostics.json on ID.'
                'Diagnostics.html'                  = 'This folder. Self-contained sortable page over both grains above. Open two runs side by side to compare nights: rows start in a stable order (computer, then path) so the two windows line up, and everything is embedded so no other log file is needed.'
                'Diagnostics.Fields.json'           = 'This file.'
            }
            KeyIdea  = 'A duration on its own cannot tell you why a job got slower. Read the duration together with ItemsWalked and AclReadMsPerItem: the first says how long, the second says how much there was, the third says what each operation cost.'
        }

        HowToRead   = [ordered]@{
            'AccountedPct well below 100'                        = 'The counters do not explain this row: most of its time went somewhere nothing measures. Do not reason about the cost fields on a row like this, and do not conclude the storage is fine because AclReadMsPerItem looks normal. It is a gap in the instrumentation, not a finding about the share.'
            'Duration up, ItemsWalked up in proportion'          = 'The share grew. Expected, nothing to fix.'
            'Duration up, ItemsWalked flat, AclReadMsPerItem up' = 'Each storage operation became more expensive. Look at the file server (backup, anti-virus, deduplication, snapshot pressure), not at the matrix.'
            'AceCountMean rising run over run'                   = 'The ACLs themselves are growing, which means permissions are being appended rather than replaced. This compounds every night and is worth fixing before anything else.'
            'IncorrectItems the same non-zero value every run'   = 'The tree never converges: the same items are corrected every night. Either something outside the matrix keeps changing them back, or the fix is not taking.'
            'AclReadDenied or AclReadFailed above zero'          = 'Items were skipped or needed an ownership takeover. These are also the slowest items, so a rise here can explain a rise in duration on its own.'
            'Comparing two nights'                              = 'Open Diagnostics.html from both run folders in two windows. Leave the sort at its default so the rows line up, filter both to the same path, and read the differences off the screen. Sort by cost only once you know which row you are chasing, since cost order differs between nights.'
            'A Settings row got slower but you cannot see why'   = 'Open Diagnostics.Paths.json and filter on that ID. Its rows are already sorted by measured cost, so the folder responsible is at the top. Compare the same folder between runs the same way you would compare a Settings row.'
            'A path row shows cost but ItemsWalked is zero'      = 'Walked is false: the folder ACL was checked but its subtree was never walked, because the folder is ignored or every child belongs to another matrix folder. Not an error.'
            TrendOneLiner                                        = "Get-ChildItem '<log root>\*\Diagnostics.json' | ForEach-Object { `$run = `$_.Directory.Name; Get-Content `$_ -Raw | ConvertFrom-Json | Where-Object Path -eq '<path>' | Select-Object @{n='Run';e={`$run}}, DurationSeconds, ItemsWalked, AclReadMsPerItem, AceCountMean }"
        }

        Caveats     = @(
            'Millisecond totals sum CONCURRENT work, so they can exceed the job wall clock by roughly the folders-per-matrix throttle. They measure cost, not elapsed time, and are comparable between runs only while MaxConcurrent is unchanged. This is why AccountedPct can legitimately exceed 100 and why UnaccountedMs can be negative; neither is an error.'
            'AccountedPct is meaningless on a row that finished in about a second: fixed per-job cost (session setup, module load, matrix marshalling) dominates a wall clock that small, so the percentage reads near zero on rows where there is nothing to explain. Ignore it unless ItemsWalked and WallClockMs are large enough for the row to be worth investigating.'
            'AccountedPct grades the instrumentation, not the run. A low value never means the server was idle: it means the time went into work that carries no counter. Compare it between runs as well, because a row whose AccountedPct falls while its duration rises has changed in a way the current fields cannot describe.'
            'AclReadMsPerItem is a SAMPLED mean (see SampleEvery and SampleWarmup). Always check AclReadSamples and AclReadBasis before trusting it: a handful of samples can land anywhere, and a basis of warmup+stride means the figure leans on the cold early items.'
            'Even with a few hundred samples the mean carries several percent of run-to-run noise, because ACL read times are heavy-tailed and whether the sample catches a slow outlier is luck. A single run moving 10% is not a signal; a trend across several runs is.'
            'AclReadMsEstimated covers ACL reads only, not enumeration, comparison or loop overhead, so it sits well below the job duration by design. It is an extrapolation, not a measurement.'
            'Items skipped before they are counted (reparse points, DFS links, system items, and folders listed in the matrix or marked to ignore) appear in EnumeratedDirs but not in ItemsWalked. EnumeratedDirs larger than ItemsWalked is normal.'
            'Counters describe the run that produced them. A row that never executed writes no diagnostics file at all, which is different from a row that executed and walked nothing.'
        )

        PathFields   = [ordered]@{
            SettingPath = 'Path from the Settings sheet that owns this folder, so a folder traces back to its row without a lookup.'
            Path        = 'The matrix folder this row describes. The unit of comparison between runs when localising a regression.'
            Walked      = 'False when the folder ACL was checked but its subtree was never walked (folder ignored, or every child belongs to another matrix folder). Explains a row with cost but no ItemsWalked.'
            Note        = 'All other columns carry the same meaning as the matching entry under TelemetryFields, scoped to this folder subtree instead of the whole Settings row.'
        }

        RecordFields = [ordered]@{
            ID           = 'Identifier of the Settings row, matching the ID shown on the execution report card and the "ID <guid> - Detail N.json" files.'
            MatrixFile   = 'File name of the matrix the Settings row came from.'
            ComputerName = 'Server the permissions were applied on. This is where the counters were measured.'
            Path         = 'Parent folder from the Settings row. The unit of comparison between runs.'
            Action       = 'New, Check or Fix. Check reads without writing, so its write counters stay at zero by definition.'
            Start        = 'When this Settings row started.'
            End          = 'When this Settings row finished.'
            Duration     = 'Wall clock for this Settings row as hh:mm:ss (per-row file only).'
            RunStartTime = 'When the whole run started. Identical on every row, so it can be used as the x-axis when charting several runs (roll-up only).'
            DurationSeconds = 'Wall clock for this Settings row in seconds, as a number rather than a formatted string, so it sorts and charts (roll-up only).'
        }

        TelemetryFields = [ordered]@{
            Path                   = [ordered]@{ Unit = 'path'; Meaning = 'Parent folder walked, repeated here so the telemetry block is readable on its own.' }
            Action                 = [ordered]@{ Unit = 'text'; Meaning = 'New, Check or Fix, repeated for the same reason.' }
            ComputerName           = [ordered]@{ Unit = 'text'; Meaning = 'Server the counters were measured on.' }
            WallClockMs            = [ordered]@{ Unit = 'milliseconds'; Meaning = 'Time spent inside the permission-setting stage on the server. Slightly less than the row Duration, which also covers the remote session setup and the return trip.' }

            AccountedMs            = [ordered]@{ Unit = 'milliseconds'; Meaning = 'The measured and estimated cost fields added together: AclReadMsEstimated, AclWriteMs, EnumerateMs, MatrixFolderReadMs and MatrixFolderWriteMs. How much of this row the counters below claim to explain.' }
            UnaccountedMs          = [ordered]@{ Unit = 'milliseconds'; Meaning = 'WallClockMs minus AccountedMs. Time inside the job that no counter attributes to anything. NOT clamped: a negative value means the counters overlap in time (see AccountedPct) rather than that something is wrong.' }
            AccountedPct           = [ordered]@{ Unit = 'percent'; Meaning = 'AccountedMs as a percentage of WallClockMs. READ THIS FIRST: it grades the measurement, not the run. Well below 100 means the counters do not explain this row and the cost fields should not be used to reason about it. Around 100 means the breakdown is trustworthy. Above 100 is normal, not an error: the millisecond totals sum concurrent work across the walker runspaces while WallClockMs is elapsed time.' }

            ItemsWalked            = [ordered]@{ Unit = 'count'; Meaning = 'Files plus folders whose ACL was examined during the inherited-permissions walk. THE volume number: compare it against Duration first.' }
            FoldersWalked          = [ordered]@{ Unit = 'count'; Meaning = 'Folders within ItemsWalked.' }
            FilesWalked            = [ordered]@{ Unit = 'count'; Meaning = 'Files within ItemsWalked.' }
            MatrixFolders          = [ordered]@{ Unit = 'count'; Meaning = 'Folders named in the Permissions worksheet that were processed with an explicit ACL. Grows when the matrix grows, not when the share does.' }
            MatrixFoldersCreated   = [ordered]@{ Unit = 'count'; Meaning = 'Matrix folders that did not exist and were created this run (Action New or Fix).' }

            SampleEvery            = [ordered]@{ Unit = 'count'; Meaning = 'After the warm-up, one item in this many is timed. A constant, recorded so the sampling scheme is legible from the file.' }
            SampleWarmup           = [ordered]@{ Unit = 'count'; Meaning = 'The first this-many items are timed unconditionally before SampleEvery takes over, so small paths still produce a usable mean.' }
            AclReadSamples         = [ordered]@{ Unit = 'count'; Meaning = 'How many ACL reads went into AclReadMsPerItem. The confidence figure: a few hundred is solid, single digits is noise.' }
            AclReadStrideSamples   = [ordered]@{ Unit = 'count'; Meaning = 'Samples taken by the 1-in-SampleEvery rule, spread evenly across the whole walk. These are the representative ones.' }
            AclReadWarmupSamples   = [ordered]@{ Unit = 'count'; Meaning = 'Samples taken from the first SampleWarmup items. Contiguous and therefore unrepresentative on a large tree (coldest items, tiny slice of the data), so they are only used when there are too few stride samples.' }
            AclReadBasis           = [ordered]@{ Unit = 'text'; Meaning = "Which pool produced AclReadMsPerItem. 'stride' is the trustworthy case. 'warmup+stride' means the subtree was too small for 30 stride samples, so the figure leans on the cold early items and should be read as indicative only. 'none' means nothing was timed." }
            AclReadMsPerItem       = [ordered]@{ Unit = 'milliseconds'; Meaning = 'Sampled mean cost of ONE ACL read, from the stride pool where possible (see AclReadBasis). The number that separates a slower disk from a bigger share, because it does not move when only the amount of data changes. Expect a few percent of run-to-run noise even on identical work; compare trends across several runs rather than reacting to one.' }
            AclReadMsEstimated     = [ordered]@{ Unit = 'milliseconds'; Meaning = 'AclReadMsPerItem multiplied by ItemsWalked. An extrapolation for apportioning the run time, not a measurement.' }

            AclWrites              = [ordered]@{ Unit = 'count'; Meaning = 'ACL writes during the walk, each one resetting an item to inherited-only. Zero for Action Check. On a settled tree this trends towards zero.' }
            AclWriteMs             = [ordered]@{ Unit = 'milliseconds'; Meaning = 'Total measured time in those writes. Timed in full rather than sampled, because writes are rare.' }
            AclWriteMsPerItem      = [ordered]@{ Unit = 'milliseconds'; Meaning = 'Mean cost of one ACL write, measured rather than sampled.' }

            EnumeratedDirs         = [ordered]@{ Unit = 'count'; Meaning = 'Directory listings opened. Normally larger than the folder count, because skipped children are enumerated before they are filtered out.' }
            EnumerateMs            = [ordered]@{ Unit = 'milliseconds'; Meaning = 'Time opening those listings. Separated from ACL time so directory-metadata slowness can be told apart from security-descriptor slowness.' }

            MatrixFolderReads      = [ordered]@{ Unit = 'count'; Meaning = 'ACL reads on the explicit matrix folders, on the orchestrating thread. Measured in full, not sampled.' }
            MatrixFolderReadMs     = [ordered]@{ Unit = 'milliseconds'; Meaning = 'Time in those reads.' }
            MatrixFolderWrites     = [ordered]@{ Unit = 'count'; Meaning = 'ACL writes on the explicit matrix folders. Zero for Action Check.' }
            MatrixFolderWriteMs    = [ordered]@{ Unit = 'milliseconds'; Meaning = 'Time in those writes.' }

            IncorrectItems         = [ordered]@{ Unit = 'count'; Meaning = 'Walked items whose ACL did not match what the matrix expects. For Check this is a finding; for Fix it is what was corrected. The same non-zero value every night means the tree is not converging.' }
            MatrixFoldersIncorrect = [ordered]@{ Unit = 'count'; Meaning = 'Explicit matrix folders whose ACL did not match. Same reading as IncorrectItems.' }
            AclReadDenied          = [ordered]@{ Unit = 'count'; Meaning = 'Reads that hit access-denied. Under Fix these trigger an ownership takeover, which is markedly slower than a normal read.' }
            AclReadFailed          = [ordered]@{ Unit = 'count'; Meaning = 'Reads that failed for a reason other than access-denied (corrupt descriptor, item locked). These items were neither checked nor corrected and need manual attention.' }
            AclWriteDenied         = [ordered]@{ Unit = 'count'; Meaning = 'Writes that hit access-denied and were retried after taking ownership.' }

            AceCountMean           = [ordered]@{ Unit = 'count'; Meaning = 'Mean number of access-control entries per item inspected. THE number to watch for ACL bloat: if permissions are appended instead of replaced, this climbs run over run and everything slows with it. Sampled during the walk, exact for matrix folders.' }
            AceCountMax            = [ordered]@{ Unit = 'count'; Meaning = 'Largest ACE count seen on any single item. Rises before the mean does when only a few folders are affected.' }
            AceCountItems          = [ordered]@{ Unit = 'count'; Meaning = 'How many items contributed to AceCountMean, so the mean can be judged the same way as AclReadSamples.' }
            Paths                  = [ordered]@{ Unit = 'array'; Meaning = 'Per-matrix-folder breakdown of everything above, sorted by measured cost descending. Because the parallel walker jobs partition the tree, these rows sum back to the totals. Flattened into Diagnostics.Paths.json; omitted from Diagnostics.json so that file stays CSV-convertible.' }
        }
    }
}

function Write-DiagnosticsFieldReferenceHC {
    <#
    .SYNOPSIS
        Writes 'Diagnostics.Fields.json' next to the run diagnostics roll-up.

    .DESCRIPTION
        Makes a run folder self-explanatory: whoever opens the logs, possibly
        months later and without the repository to hand, can read what every
        counter means and what the combinations imply.

    .NOTES
        Failures are swallowed, like the diagnostics writers themselves. A
        reference document that could not be written is not worth failing a run
        over, and not worth a line in the summary mail.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$LogFolder
    )

    try {
        Get-DiagnosticsFieldReferenceHC |
        ConvertTo-Json -Depth 10 |
        Out-File `
            -FilePath (Join-Path -Path $LogFolder -ChildPath 'Diagnostics.Fields.json') `
            -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Failed writing the diagnostics field reference: $_"
    }
}

function Out-LogFileHC {
    <#
    .SYNOPSIS
        Writes PowerShell objects to several file formats at once, from one
        shared base path.

    .DESCRIPTION
        Exports DataToExport to each requested extension ('.csv', '.json',
        '.txt', '.xlsx'), appending the extension to PartialPath. Extensions are
        de-duplicated and sorted before use. Returns the paths actually written.

        Format-specific handling:
        - JSON: ErrorRecord values are flattened to their message text so
          serialization cannot fail on depth. The caller's DataToExport is never
          mutated: a fresh object is built per item. With -Append the existing
          file is read and merged, rather than concatenating raw text.
        - XLSX: built through ImportExcel with frozen headers and auto-sized
          columns. Without -Append an existing file is deleted first, because
          Export-Excel is always called in append mode.

    .PARAMETER ExcelFile
        Formatting rules for '.xlsx'. Keys: 'SheetName', 'TableName',
        'CellStyle'.

    .NOTES
        A format that fails is reported with Write-Warning and its path is
        omitted from the return value; the remaining formats still export. A
        caller that ignores the returned paths will not notice a partial
        failure. An unsupported extension throws inside the loop and is caught
        by that same handler.

    .OUTPUTS
        System.String[]
        Paths of the log files successfully generated or updated.

    .EXAMPLE
        $exportedPaths = Out-LogFileHC `
            -DataToExport $data `
            -PartialPath 'C:\Logs\DailyReport' `
            -FileExtensions @('.csv', '.json', '.xlsx')
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [PSCustomObject[]]$DataToExport,
        [Parameter(Mandatory)] [String]$PartialPath,
        [Parameter(Mandatory)] [String[]]$FileExtensions,
        [hashtable]$ExcelFile = @{
            SheetName = 'Overview'
            TableName = 'Overview'
            CellStyle = $null
        },
        [Switch]$Append
    )

    $allPaths = @()

    foreach ($ext in ($FileExtensions | Sort-Object -Unique)) {

        $logFilePath = "$PartialPath$ext"

        try {
            switch ($ext) {

                '.csv' {
                    $DataToExport |
                    Export-Csv -LiteralPath $logFilePath -Delimiter ';' `
                        -Append:$Append -NoTypeInformation
                    break
                }

                '.json' {
                    # Build a new object per item so the caller's $DataToExport
                    # is never mutated. ErrorRecord values are rendered to their
                    # message text; all other values are copied as-is.
                    $converted = foreach ($item in $DataToExport) {
                        $clone = [ordered]@{}
                        foreach ($p in $item.PSObject.Properties) {
                            $clone[$p.Name] = if ($p.Value -is [System.Management.Automation.ErrorRecord]) {
                                $p.Value.Exception.Message
                            }
                            else {
                                $p.Value
                            }
                        }
                        [PSCustomObject]$clone
                    }

                    if ($Append -and (Test-Path $logFilePath)) {
                        $existing = Get-Content -LiteralPath $logFilePath -Raw | ConvertFrom-Json
                        $converted = @($existing) + @($converted)
                    }

                    $converted |
                    ConvertTo-Json -Depth 7 |
                    Out-File -LiteralPath $logFilePath -Encoding utf8 -Force
                    break
                }

                '.txt' {
                    $DataToExport |
                    Format-List * |
                    Out-File -LiteralPath $logFilePath -Append:$Append
                    break
                }

                '.xlsx' {
                    if (-not $Append -and (Test-Path $logFilePath)) {
                        Remove-Item -LiteralPath $logFilePath -Force
                    }

                    $params = @{
                        Path          = $logFilePath
                        Append        = $true
                        AutoNameRange = $true
                        AutoSize      = $true
                        FreezeTopRow  = $true
                        WorksheetName = $ExcelFile.SheetName
                        TableName     = $ExcelFile.TableName
                    }

                    if ($ExcelFile.CellStyle) {
                        $params.CellStyleSB = $ExcelFile.CellStyle
                    }

                    $DataToExport | Export-Excel @params
                    break
                }

                default {
                    throw "Unsupported file extension '$ext'."
                }
            }

            $allPaths += $logFilePath
        }
        catch {
            Write-Warning "Failed to export log '$logFilePath': $_"
        }
    }

    return $allPaths
}

function Remove-FileHC {
    <#
    .SYNOPSIS
        Deletes a file, downgrading locking/permission failures to a warning.

    .DESCRIPTION
        Removes FilePath when it exists. A missing file is a silent no-op.

    .NOTES
        Never throws. On failure the warning is routed to SystemErrors when that
        reference is supplied, and to the PowerShell warning stream when it is
        not.

    .EXAMPLE
        $sysErrors = [System.Collections.Generic.List[pscustomobject]]::new()
        Remove-FileHC `
            -FilePath 'C:\Temp\OldLog.txt' `
            -SystemErrors ([ref]$sysErrors)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$FilePath,
        [ref]$SystemErrors
    )

    try {
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { return }
        Remove-Item -LiteralPath $FilePath -Force -ErrorAction Stop
    }
    catch {
        if ($SystemErrors) {
            Add-ErrorHC `
                -Type 'Warning' `
                -Name 'Failed to remove file' `
                -Message "Failed to remove '$FilePath': $_" `
                -Category 'Logging' `
                -SystemErrors $SystemErrors
        }
        else {
            Write-Warning "Failed removing '$FilePath': $_"
        }
    }
}

function Write-EventLogSafeHC {
    <#
    .SYNOPSIS
        Formats and writes aggregated execution data and system errors to the
        Windows Event Log, without ever throwing.

    .DESCRIPTION
        Returns immediately unless 'SaveInEventLog.Save' is set and a LogName is
        configured. Otherwise it adds every entry in SystemErrors as its own
        Error event (EventID 2), appends a 'Script ended' Information event
        (EventID 199), then hands the batch to Write-EventsToEventLogHC.

    .NOTES
        - Messages longer than 31,000 characters are truncated with a marker.
          The Event Log API throws on oversized entries, which a large stack
          trace or data dump would otherwise trigger.
        - Failures (no permission to create the source or write the log) are
          appended to SystemErrors as a warning. Contrast
          Write-EventsToEventLogHC, which throws.

    .EXAMPLE
        Write-EventLogSafeHC `
            -EventLogData $eventData `
            -ScriptName 'Permission Matrix' `
            -Settings $Context.Config.Settings `
            -SystemErrors ([ref]$sysErrors)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$EventLogData,
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)][object]$Settings,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    $maxLen = 31000

    try {
        $logName = Get-StringValueHC $Settings.SaveInEventLog.LogName
        if (-not ($Settings.SaveInEventLog.Save -and $logName)) { return }

        # Append SystemErrors as individual error entries
        foreach ($err in $SystemErrors.Value) {
            $EventLogData.Add(
                [PSCustomObject]@{
                    Message   = $err.Message
                    DateTime  = $err.DateTime
                    EntryType = 'Error'
                    EventID   = '2'
                }
            )
        }

        # Add “script ended”
        $EventLogData.Add(
            [PSCustomObject]@{
                Message   = 'Script ended'
                DateTime  = Get-Date
                EntryType = 'Information'
                EventID   = '199'
            }
        )

        # Truncate too-long messages
        foreach ($item in $EventLogData) {
            if ($item.Message.Length -gt $maxLen) {
                $item.Message =
                $item.Message.Substring(0, $maxLen) +
                '... [TRUNCATED DUE TO EVENT LOG SIZE LIMITS]'
            }
        }

        Write-EventsToEventLogHC `
            -Source $ScriptName `
            -LogName $logName `
            -Events $EventLogData
    }
    catch {
        Add-ErrorHC `
            -Type 'Warning' `
            -Name 'Failed to write to event log' `
            -Message "Failed writing to event log: $_" `
            -Category 'Logging' `
            -SystemErrors $SystemErrors
    }
}

function Write-EventsToEventLogHC {
    <#
    .SYNOPSIS
        Writes an array of custom objects to the Windows Event Log, flattening
        their properties into the event message.

    .DESCRIPTION
        Registers the Event Source when missing, then writes one event per
        object. 'EntryType' and 'EventID' map to the matching Event Log fields;
        every other property is flattened into a bulleted message body. Missing
        values default to 'Information' and EventID 4.

    .NOTES
        - Creating a new Event Source requires Administrator privileges.
        - Unlike its wrapper Write-EventLogSafeHC, this function THROWS on
          failure and passes the exception back up the chain.

    .EXAMPLE
        $events = @(
            [pscustomobject]@{
                EntryType = 'Warning'
                EventID   = 99
                Action    = 'Cleanup'
                Status    = 'Folder locked by another process'
            }
        )

        Write-EventsToEventLogHC `
            -Source 'MyScript' `
            -LogName 'Application' `
            -Events $events
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$Source,
        [Parameter(Mandatory)][String]$LogName,
        [PSCustomObject[]]$Events
    )

    try {
        if (-not [System.Diagnostics.EventLog]::SourceExists($Source)) {
            New-EventLog -LogName $LogName -Source $Source -EA Stop
        }

        foreach ($eventItem in $Events) {

            $params = @{
                LogName     = $LogName
                Source      = $Source
                EntryType   = $eventItem.EntryType
                EventID     = $eventItem.EventID
                Message     = ''
                ErrorAction = 'Stop'
            }

            if (-not $params.EntryType) { $params.EntryType = 'Information' }
            if (-not $params.EventID) { $params.EventID = 4 }

            foreach ($prop in $eventItem.PSObject.Properties |
                Where-Object { $_.Name -notin 'EntryType', 'EventID' }) {

                $params.Message += "`n- $($prop.Name): $($prop.Value)"
            }

            Write-EventLog @params
        }
    }
    catch {
        throw "Failed writing events to Windows Event Log: $_"
    }
}

function Write-SystemErrorLogHC {
    <#
    .SYNOPSIS
        Exports system errors to a JSON log file and attaches it to the outgoing
        email parameters.

    .DESCRIPTION
        Serializes SystemErrors to 'SystemErrors.json' in the dated log folder,
        then adds that path to the 'Attachments' key of the referenced
        MailParams hashtable, creating the key when absent. Administrators then
        receive the raw error data alongside the HTML summary.

    .PARAMETER MailParams
        A [ref] to the SMTP splatting hashtable. MUTATED in place.

    .NOTES
        Returns without doing anything when there are no system errors or when
        the log folder does not exist.

    .EXAMPLE
        $mailSplat = @{ To = 'admin@domain.com'; Subject = 'Execution Report' }

        Write-SystemErrorLogHC `
            -SystemErrors $sysErrors `
            -LogFolder 'C:\MatrixLogs' `
            -MailParams ([ref]$mailSplat)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$SystemErrors,
        [Parameter(Mandatory)][string]$LogFolder,
        [Parameter(Mandatory)][ref]$MailParams,
        [datetime]$ScriptStartTime = (Get-Date),
        [string]$JsonFileName = 'MatrixConfig' 
    )

    if ($SystemErrors.Count -eq 0) { return }
    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return }

    $datedFolder = Get-DatedLogFolderPathHC `
        -LogFolder $LogFolder `
        -ScriptStartTime $ScriptStartTime `
        -JsonFileName $JsonFileName

    $partial = Join-Path $datedFolder 'SystemErrors'

    $attachments = Out-LogFileHC `
        -DataToExport $SystemErrors `
        -PartialPath $partial `
        -FileExtensions '.json' `
        -ErrorAction Ignore

    if ($attachments) {
        if (-not $MailParams.Value.ContainsKey('Attachments')) {
            $MailParams.Value['Attachments'] = @()
        }
        $MailParams.Value['Attachments'] += $attachments
    }
}