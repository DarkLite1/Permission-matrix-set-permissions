#Requires -Version 7

<#
    JOB DURATION CACHE

    Remembers how long each job took on the previous run so the next run can
    start the expensive jobs first.

    WHY:
        Jobs targeting one computer are drained from a shared queue by several
        workers. The queue is filled in matrix-file order, so whether the
        longest job starts first or last is pure luck, and that decides the wall
        clock: a run whose single longest job takes 25 minutes finishes in ~25
        minutes if it starts immediately and ~38 minutes if it starts last.

        Ordering longest-first (the standard LPT heuristic) removes that
        variance. Durations are not knowable in advance, but last night's
        duration is an excellent estimate of tonight's.

    NON-BLOCKING BY DESIGN:
        The cache is a performance hint and nothing more. Correctness never
        depends on it, so every failure mode - file missing, unreadable,
        corrupt, truncated, half-written, wrong schema, no permission, disk
        full - degrades silently to the previous behaviour (matrix-file order).

        Nothing here throws, and nothing here writes to SystemErrors: a stale
        cache is not something to wake anybody up about, and putting it in the
        summary mail would turn a hint into noise. Diagnostics go to
        Write-Verbose only.

    FILE FORMAT (JobDurations.json, at the root of the configured log folder):

        {
          "Version": 1,
          "Jobs": {
            "server01|\\\\server01\\share\\folder|fix": {
              "Seconds": 1465.32,
              "LastSeen": "2026-07-28T02:14:33.0000000Z"
            }
          }
        }

        Version exists so a future format change can be detected and the file
        ignored rather than misread.
#>

function Get-JobDurationCacheKeyHC {
    <#
    .SYNOPSIS
        Build the cache key that identifies one job across runs.

    .DESCRIPTION
        A job is not a matrix file: one file's Settings sheet can target several
        computers and paths, each taking its own amount of time. The identity of
        a unit of work is therefore the computer, the path and the action
        together.

        Windows paths are case-insensitive and cannot contain '|', so a
        lowercased pipe-joined triplet is both safe and stable. Trailing
        separators are trimmed so '\\server\share\folder' and
        '\\server\share\folder\' are recognised as the same job.
    #>
    [OutputType([string])]
    param (
        [string]$ComputerName,
        [string]$Path,
        [string]$Action
    )

    $safeComputerName = if ($ComputerName) {
        $ComputerName.Trim().ToLowerInvariant()
    }
    else { '' }

    $safePath = if ($Path) {
        $Path.Trim().TrimEnd('\', '/').ToLowerInvariant()
    }
    else { '' }

    $safeAction = if ($Action) {
        $Action.Trim().ToLowerInvariant()
    }
    else { '' }

    '{0}|{1}|{2}' -f $safeComputerName, $safePath, $safeAction
}

function Get-JobDurationCacheFilePathHC {
    <#
    .SYNOPSIS
        The full path of the cache file for a given log folder.

    .DESCRIPTION
        The file lives at the ROOT of the log folder, not inside the dated
        subfolders: it has to be found again on the next run, and a dated path
        would be a different path every night.

        It is rewritten on every run, so its timestamp always stays fresh and
        the DeleteLogsAfterDays cleanup will not age it out.
    #>
    [OutputType([string])]
    param (
        [Parameter(Mandatory)][string]$LogFolder
    )

    Join-Path -Path $LogFolder -ChildPath 'JobDurations.json'
}

function Get-JobDurationCacheHC {
    <#
    .SYNOPSIS
        Read the previous run's job durations.

    .DESCRIPTION
        Returns a hashtable of cache key -> duration in seconds.

        Returns an EMPTY hashtable for every failure mode rather than throwing:
        no log folder configured, file absent, unreadable, corrupt, truncated,
        or written by a future version of this script. The caller treats an
        empty result and a missing entry identically, so no error handling is
        needed at the call site.

        Individual malformed entries are skipped without discarding the rest of
        the file, so one bad record cannot cost the whole optimisation.

    .PARAMETER LogFolder
        Root of the configured log folder. An empty value returns an empty
        cache.

    .EXAMPLE
        $cache = Get-JobDurationCacheHC -LogFolder 'C:\Log\Permission matrix'
        $cache['server01|\\server01\share|fix']

        # Output: 1465.32
    #>
    [OutputType([hashtable])]
    param (
        [string]$LogFolder
    )

    $cache = @{}

    try {
        if ([string]::IsNullOrWhiteSpace($LogFolder)) {
            Write-Verbose 'No log folder available, job duration cache skipped'
            return $cache
        }

        $cacheFile = Get-JobDurationCacheFilePathHC -LogFolder $LogFolder

        if (-not (Test-Path -LiteralPath $cacheFile -PathType Leaf)) {
            Write-Verbose "Job duration cache '$cacheFile' not found, jobs run in matrix file order"
            return $cache
        }

        $content = Get-Content -LiteralPath $cacheFile -Raw -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($content)) {
            Write-Verbose "Job duration cache '$cacheFile' is empty"
            return $cache
        }

        $json = $content | ConvertFrom-Json -ErrorAction Stop

        # A file written by a newer version may mean something different by the
        # same field names, so it is ignored rather than guessed at.
        if ($json.Version -gt 1) {
            Write-Verbose "Job duration cache '$cacheFile' has unsupported version '$($json.Version)', ignored"
            return $cache
        }

        if (-not $json.Jobs) {
            Write-Verbose "Job duration cache '$cacheFile' holds no jobs"
            return $cache
        }

        foreach ($entry in $json.Jobs.PSObject.Properties) {
            # One malformed record must not discard the others.
            try {
                $seconds = [double]$entry.Value.Seconds

                if (
                    ($seconds -gt 0) -and
                    -not [double]::IsNaN($seconds) -and
                    -not [double]::IsInfinity($seconds)
                ) {
                    $cache[$entry.Name] = $seconds
                }
            }
            catch {
                Write-Verbose "Skipping malformed job duration entry '$($entry.Name)'"
            }
        }

        Write-Verbose "Read $($cache.Count) job duration(s) from '$cacheFile'"
    }
    catch {
        Write-Verbose "Job duration cache could not be read, jobs run in matrix file order: $_"
        return @{}
    }

    $cache
}

function Save-JobDurationCacheHC {
    <#
    .SYNOPSIS
        Record this run's job durations for the next run to use.

    .DESCRIPTION
        Merges the durations observed in this run into whatever the file already
        holds, prunes entries not seen for RetentionDays, and writes the result.

        Merging rather than replacing matters because a run does not necessarily
        touch every job: a matrix file that was disabled or absent tonight keeps
        its remembered duration for when it comes back.

        Only successful timings are recorded. A job that failed, was skipped, or
        has no JobTime is left alone rather than being written as zero, which
        would push a genuinely expensive job to the back of tomorrow's queue.

        The write is atomic - a temporary file in the same folder is renamed over
        the target - so an interrupted run cannot leave a half-written file for
        the next run to trip over.

        Never throws. Every failure is a silent no-op: the next run simply reads
        the older file, or none at all.

    .PARAMETER LogFolder
        Root of the configured log folder. An empty value makes this a no-op.

    .PARAMETER Matrices
        The flattened per-Settings-row matrix objects, normally
        $Context.AllMatrices. Each needs Setting.Formatted.ComputerName, .Path
        and .Action, plus JobTime.Duration.

    .PARAMETER RetentionDays
        Entries not seen for this many days are dropped, so paths that no longer
        exist do not accumulate forever. (Default: 60)

    .EXAMPLE
        Save-JobDurationCacheHC -LogFolder $logFolder -Matrices $Context.AllMatrices
    #>
    [CmdletBinding()]
    param (
        [string]$LogFolder,
        $Matrices,
        [ValidateRange(1, 3650)]
        [int]$RetentionDays = 60
    )

    try {
        if ([string]::IsNullOrWhiteSpace($LogFolder)) {
            Write-Verbose 'No log folder available, job durations not saved'
            return
        }

        if (-not $Matrices) {
            Write-Verbose 'No matrices to record job durations for'
            return
        }

        if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
            Write-Verbose "Log folder '$LogFolder' does not exist, job durations not saved"
            return
        }

        $cacheFile = Get-JobDurationCacheFilePathHC -LogFolder $LogFolder
        $now = (Get-Date).ToUniversalTime()
        $cutOff = $now.AddDays(-$RetentionDays)

        #region Start from the entries already on disk
        $merged = @{}

        try {
            if (Test-Path -LiteralPath $cacheFile -PathType Leaf) {
                $existing = Get-Content -LiteralPath $cacheFile -Raw -ErrorAction Stop |
                    ConvertFrom-Json -ErrorAction Stop

                if ($existing.Version -le 1) {
                    foreach ($entry in $existing.Jobs.PSObject.Properties) {
                        try {
                            $lastSeen = [datetime]::Parse(
                                $entry.Value.LastSeen,
                                [cultureinfo]::InvariantCulture,
                                [System.Globalization.DateTimeStyles]::RoundtripKind
                            )

                            if ($lastSeen -ge $cutOff) {
                                $merged[$entry.Name] = [PSCustomObject]@{
                                    Seconds  = [double]$entry.Value.Seconds
                                    LastSeen = $lastSeen
                                }
                            }
                        }
                        catch {
                            # Undated or unparseable entry: drop it. It will be
                            # re-learned the next time that job runs.
                            Write-Verbose "Dropping unreadable job duration entry '$($entry.Name)'"
                        }
                    }
                }
            }
        }
        catch {
            # An unreadable existing file is not worth preserving; start fresh
            # from this run's observations.
            Write-Verbose "Existing job duration cache could not be merged, writing a new one: $_"
            $merged = @{}
        }
        #endregion

        #region Overlay this run's observations
        $observed = 0

        foreach ($matrix in $Matrices) {
            $duration = $matrix.JobTime.Duration

            # Skipped and failed jobs have no usable duration. Writing zero
            # would demote an expensive job in tomorrow's ordering, so leave
            # whatever is already remembered in place.
            if (-not ($duration -is [timespan])) { continue }
            if ($duration.TotalSeconds -le 0) { continue }

            $key = Get-JobDurationCacheKeyHC `
                -ComputerName $matrix.Setting.Formatted.ComputerName `
                -Path $matrix.Setting.Formatted.Path `
                -Action $matrix.Setting.Formatted.Action

            $merged[$key] = [PSCustomObject]@{
                Seconds  = [math]::Round($duration.TotalSeconds, 2)
                LastSeen = $now
            }

            $observed++
        }

        if ($observed -eq 0) {
            Write-Verbose 'No completed jobs with a duration, cache left unchanged'
            return
        }
        #endregion

        #region Write atomically
        # Keys are written in sorted order so a diff between two runs shows what
        # actually changed rather than a reshuffle.
        $jobs = [ordered]@{}

        foreach ($key in ($merged.Keys | Sort-Object)) {
            $jobs[$key] = [ordered]@{
                Seconds  = $merged[$key].Seconds
                LastSeen = $merged[$key].LastSeen.ToString('o')
            }
        }

        $payload = [ordered]@{
            Version = 1
            Updated = $now.ToString('o')
            Jobs    = $jobs
        }

        # Same folder as the target, so the rename stays on one volume and is
        # therefore atomic.
        $tempFile = '{0}.{1}.tmp' -f $cacheFile, [guid]::NewGuid().ToString('N')

        try {
            $payload |
                ConvertTo-Json -Depth 5 |
                Set-Content -LiteralPath $tempFile -Encoding UTF8 -ErrorAction Stop

            Move-Item -LiteralPath $tempFile -Destination $cacheFile -Force -ErrorAction Stop

            Write-Verbose "Saved $observed job duration(s) of $($jobs.Count) remembered to '$cacheFile'"
        }
        catch {
            # Never leave the scratch file behind for the cleanup job to puzzle over.
            if (Test-Path -LiteralPath $tempFile -PathType Leaf) {
                Remove-Item -LiteralPath $tempFile -Force -ErrorAction Ignore
            }

            throw
        }
        #endregion
    }
    catch {
        Write-Verbose "Job durations could not be saved, next run falls back to matrix file order: $_"
    }
}

function Get-JobDurationEstimateHC {
    <#
    .SYNOPSIS
        The remembered duration of a job, or a value that sorts it first.

    .DESCRIPTION
        Used as the sort key when filling a computer's work queue.

        A job with no cached duration returns [double]::MaxValue so it sorts
        ahead of everything else. That is deliberate: an unknown job may be
        trivial or enormous, and the payoff is asymmetric. Starting a large
        unknown early saves the whole ordering benefit, while starting a small
        one early costs almost nothing.

    .PARAMETER Cache
        The hashtable returned by Get-JobDurationCacheHC.

    .EXAMPLE
        Get-JobDurationEstimateHC -Cache $cache -ComputerName 'server01' `
            -Path '\\server01\share' -Action 'Fix'

        # Output: 1465.32, or [double]::MaxValue when not cached
    #>
    [OutputType([double])]
    param (
        [hashtable]$Cache,
        [string]$ComputerName,
        [string]$Path,
        [string]$Action
    )

    if (-not $Cache -or $Cache.Count -eq 0) {
        return [double]::MaxValue
    }

    $key = Get-JobDurationCacheKeyHC `
        -ComputerName $ComputerName `
        -Path $Path `
        -Action $Action

    if ($Cache.ContainsKey($key)) {
        return [double]$Cache[$key]
    }

    [double]::MaxValue
}
