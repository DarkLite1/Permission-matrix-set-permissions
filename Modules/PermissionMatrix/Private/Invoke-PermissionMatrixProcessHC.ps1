function Invoke-PermissionMatrixProcessHC {
    <#
    .SYNOPSIS
        Executes the core remote processing stage of the Permission Matrix 
        pipeline.

    .DESCRIPTION
        This function serves as the 'PROCESS' stage of the orchestrator. It 
        filters out any matrices that suffered validation failures during the 
        'BEGIN' stage and executes the remaining jobs against their target 
        servers.

        Execution is broken into two highly optimized, multi-threaded phases:
        
        1. Requirements Validation: 
            Groups jobs by target 'ComputerName' and executes 'TestRequirements.
            ps1' in parallel. This verifies that the remote servers meet the 
            minimum PowerShell/.NET requirements and enforces baseline SMB 
            share settings.
        2. Permission Application: 
            Filters out any matrices that failed the requirements check, then 
            flattens the matrix data into safe Data Transfer Objects (DTOs). It 
            executes 'SetPermissions.ps1' in parallel, pushing the strict NTFS 
            permission arrays down to the target servers for evaluation and 
            enforcement.

        Architectural Note: By grouping tasks by 'ComputerName' and executing 
        via runspaces, the script drastically reduces WinRM connection overhead 
        and maximizes network throughput.

    .PARAMETER Context
        The global pipeline context object built during the 'BEGIN' stage. Must 
        contain the populated 'AllMatrices' array and configuration settings.

    .PARAMETER SystemErrors
        A reference variable ([ref]) containing a List[pscustomobject]. Used to 
        capture and bubble up terminating pipeline errors that occur during 
        remote execution routing.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
        Returns the updated `$Context` object, with the '.Check' lists of 
        individual matrices populated with the remote execution results (Errors/
        Warnings) and precise job duration timings.

    .EXAMPLE
        $sysErrors = [System.Collections.Generic.List[pscustomobject]]::new()
        
        $context = Invoke-PermissionMatrixProcessHC `
            -Context $context `
            -SystemErrors ([ref]$sysErrors)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Context,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    try {
        if (-not $Context.AllMatrices -or $Context.AllMatrices.Count -eq 0) {
            return $Context
        }

        #region Filter out matrices with fatal errors before processing
        $validMatrices = [System.Collections.Generic.List[pscustomobject]]::new()

        foreach ($file in $Context.FileResults) {
            # A fatal file-level error or a fatal Permissions sheet error blocks
            # permission application. FormData sheet errors do NOT, because they
            # only affect the ServiceNow export, not the NTFS permissions.
            if (
                Test-FileHasFatalErrorHC -File $file
            ) {
                continue
            }
            
            foreach ($matrixObj in $file.Matrices) {
                if (
                    -not (Test-ItemHasFatalErrorHC -CheckList $matrixObj.Check)
                ) {
                    $validMatrices.Add($matrixObj)
                }
            }
        }

        if ($validMatrices.Count -eq 0) {
            Write-Verbose 'No executable matrices found after initial validation.'
            return $Context
        }
        #endregion

        #region Set throttling and session configuration
        # CONCURRENCY
        #
        # Two independent caps, both meaning exactly what they say:
        #
        #   JobsPerComputer - the most jobs that may run at the same time on any
        #                     ONE remote computer. A job holds exactly one
        #                     PSSession for its lifetime, so this is also the
        #                     most WinRM shells this script will ever open on a
        #                     single machine. Enforced STRUCTURALLY: a computer
        #                     contributes exactly this many workers (see
        #                     $workerCount below), so no throttle value can
        #                     breach it.
        #
        #   JobsTotal       - the most jobs that may run at the same time across
        #                     ALL computers. Applied directly as -ThrottleLimit.
        #
        # WinRM's MaxShellsPerUser is enforced by the REMOTE server and limits
        # shells per user on that machine, so JobsPerComputer is the setting
        # that keeps it satisfied. Its default is commonly 30; confirm with
        # `winrm get winrm/config/winrs` on the busy hosts before raising
        # either number.
        $throttleJobsTotal = if (
            [string]::IsNullOrWhiteSpace($Context.Config.MaxConcurrent.JobsTotal)
        ) {
            10
        }
        else {
            [math]::Max(1, [int]$Context.Config.MaxConcurrent.JobsTotal)
        }

        $throttleJobsPerComputer = if (
            [string]::IsNullOrWhiteSpace($Context.Config.MaxConcurrent.JobsPerComputer)
        ) {
            1
        }
        else {
            [math]::Max(1, [int]$Context.Config.MaxConcurrent.JobsPerComputer)
        }

        $psSessionConfig = if (
            [string]::IsNullOrWhiteSpace($Context.Config.PSSessionConfiguration)
        ) {
            'PowerShell.7'
        }
        else {
            $Context.Config.PSSessionConfiguration
        }
        #endregion

        #region Test Requirements - Parallel by Computer
        $matrixGroups = $validMatrices | Group-Object -Property { 
            $_.Setting.Formatted.ComputerName
        }
        
        # DTO FLATTENING: Protects deep properties from runspace truncation 
        $safeReqGroups = foreach ($group in $matrixGroups) {
            [PSCustomObject]@{
                ComputerName = $group.Name
                PathsToCheck = @($group.Group.Setting.Formatted.Path)
            }
        }

        if ($safeReqGroups) {
            $reqResults = Invoke-WithOptionalParallelismHC `
                -InputObject $safeReqGroups `
                -ThrottleLimit $throttleJobsTotal `
                -ArgumentList $Context.ScriptPath, $psSessionConfig `
                -ScriptBlock {
                param($dto, $scriptPaths, $sessionConfig)

                # Raise the client-side deserialization cap. By default a
                # remote result larger than 200MB (MaximumReceivedObjectSize =
                # 209715200) makes Invoke-Command fail with 'The current
                # deserialized object size ... exceeded the allowed maximum
                # object size'. Large shares can return that much reporting
                # data, so lift the limit to the Int32 maximum (~2GB).
                $sessionOption = New-PSSessionOption `
                    -MaximumReceivedObjectSize ([Int32]::MaxValue)

                try {
                    $result = Invoke-Command `
                        -FilePath $scriptPaths.TestRequirements `
                        -ArgumentList $dto.PathsToCheck, $true `
                        -ConfigurationName $sessionConfig `
                        -ComputerName $dto.ComputerName `
                        -SessionOption $sessionOption `
                        -ErrorAction Stop
                
                    return [PSCustomObject]@{ 
                        ComputerName = $dto.ComputerName 
                        Result       = $result 
                    }
                }
                catch {
                    $errObj = [PSCustomObject]@{ 
                        DateTime    = Get-Date
                        Type        = 'FatalError'
                        Name        = 'Computer requirements'
                        Description = 'Failed checking computer requirements.' 
                        Value       = $_ 
                    } 
                    return [PSCustomObject]@{ 
                        ComputerName = $dto.ComputerName
                        Result       = $errObj 
                    }
                }
            }

            # Index matrices by ComputerName so each computer's requirement
            # result applies in O(1) instead of rescanning $validMatrices.
            $matricesByComputer = @{}
            foreach ($m in $validMatrices) {
                $cn = $m.Setting.Formatted.ComputerName
                if (-not $matricesByComputer.ContainsKey($cn)) {
                    $matricesByComputer[$cn] =
                    [System.Collections.Generic.List[object]]::new()
                }
                $matricesByComputer[$cn].Add($m)
            }

            # Main Thread Application: Add results back to the live objects
            foreach ($output in $reqResults) {
                if ($output.Result) {
                    $targetMatrices = $matricesByComputer[$output.ComputerName]

                    foreach ($m in $targetMatrices) {
                        $structured = @($output.Result | 
                            ConvertTo-StructuredObjectHC)

                        foreach ($entry in $structured) {
                            $m.Check.Add($entry)
                        }
                    }
                }
            }
        }
        #endregion

        #region Set Permissions - Parallel by Computer
        $matricesToExecute = $validMatrices.Where(
            { $_.Check.Type -notcontains 'FatalError' }
        )

        if ($matricesToExecute.Count -eq 0) { return $Context }

        $compGroupsForPerms = $matricesToExecute |
        Group-Object -Property { $_.Setting.Formatted.ComputerName }

        # DTO FLATTENING: Protects deep properties from runspace truncation.
        #
        # WORK QUEUE PER COMPUTER
        #
        # The obvious way to run several jobs per computer is to nest a second
        # parallel loop inside this one. That does not work here: a nested
        # ForEach-Object -Parallel cannot see module-private functions, and it
        # bypasses Invoke-WithOptionalParallelismHC's sequential fallback, which
        # is the seam the unit tests rely on to mock Invoke-Command.
        #
        # So the work is flattened to ONE level instead:
        #   - each computer owns a ConcurrentQueue holding all of its jobs
        #   - each computer contributes up to JobsPerComputer WORKERS, all
        #     draining that same queue
        #   - every worker, across every computer, is a single input item for a
        #     single Invoke-WithOptionalParallelismHC call
        #
        # Workers pull rather than being handed a fixed slice, so one very slow
        # matrix cannot strand its share of the queue while other workers idle.
        # ORDERING: longest job first, within each queue and across computers.
        #
        # Whether the single longest job starts at the beginning or the end of a
        # queue decides the wall clock, and matrix-file order decides it by
        # accident. The cache remembers how long each job took last night, which
        # is a good estimate of tonight.
        #
        # This is a hint only. A missing, unreadable or corrupt cache returns
        # empty and every job is then treated as unknown, which reproduces the
        # previous behaviour. Nothing below can throw on that account.
        $jobDurationCache = Get-JobDurationCacheHC `
            -LogFolder $Context.Config.Settings.SaveLogFiles.Where.Folder

        $safePermWorkers = foreach ($group in $compGroupsForPerms) {
            $jobQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()

            # Estimate each job's duration ONCE and reuse it for both the
            # ordering and the queue-cost total. Get-JobDurationEstimateHC is a
            # cache lookup, so computing it once per job (rather than once in
            # Sort-Object and again in the cost loop) avoids duplicate work.
            $estimatedJobs = foreach ($S in @($group.Group)) {
                [PSCustomObject]@{
                    Setting  = $S
                    Estimate = Get-JobDurationEstimateHC `
                        -Cache $jobDurationCache `
                        -ComputerName $S.Setting.Formatted.ComputerName `
                        -Path $S.Setting.Formatted.Path `
                        -Action $S.Setting.Formatted.Action
                }
            }

            # Workers drain the queue in order, so enqueueing expensive first is
            # what actually applies the ordering. Unknown jobs sort ahead of
            # everything (see Get-JobDurationEstimateHC).
            $orderedJobs = @($estimatedJobs) |
            Sort-Object -Property Estimate -Descending

            $queueCost = 0

            foreach ($orderedJob in $orderedJobs) {
                $S = $orderedJob.Setting

                $jobQueue.Enqueue(
                    [PSCustomObject]@{
                        ID           = $S.ID
                        ComputerName = $S.Setting.Formatted.ComputerName
                        Path         = $S.Setting.Formatted.Path
                        Action       = $S.Setting.Formatted.Action
                        MatrixJson   = (
                            $S.Matrix | 
                            ConvertTo-Json -Depth 10 -Compress
                        )
                    }
                )

                # MaxValue means 'unknown', not 'astronomically expensive': adding
                # it would overflow the total and make every queue holding one
                # unknown job look identical. Count it as zero and let the
                # QueueHasUnknown flag speak for it instead.
                if ($orderedJob.Estimate -ne [double]::MaxValue) {
                    $queueCost += $orderedJob.Estimate
                }
            }

            # Never create more workers than there is work for them to do.
            $workerCount = [math]::Min($throttleJobsPerComputer, $jobQueue.Count)

            for ($worker = 0; $worker -lt $workerCount; $worker++) {
                [PSCustomObject]@{
                    ComputerName = $group.Name
                    JobQueue     = $jobQueue
                    QueueCost    = $queueCost
                    QueueDepth   = $jobQueue.Count
                }
            }
        }

        # ForEach-Object -Parallel consumes the pipeline in order, so whichever
        # workers appear first are the ones that start first. Emitting them in
        # group order would let the busiest server begin last and become the
        # critical path for no reason.
        #
        # Longest-processing-time-first across computers: known cost is the best
        # signal, queue depth is the fallback when nothing is cached. Sorting on
        # both means a first run behaves exactly as before.
        $safePermWorkers = @($safePermWorkers) |
        Sort-Object -Property `
        @{ Expression = { $_.QueueCost }; Descending = $true },
        @{ Expression = { $_.QueueDepth }; Descending = $true }

        if ($safePermWorkers) {
            $permResults = Invoke-WithOptionalParallelismHC `
                -InputObject @($safePermWorkers) `
                -ThrottleLimit $throttleJobsTotal `
                -ArgumentList $Context.ScriptPath, $psSessionConfig, $Context.Config.MaxConcurrent, $Context.Config.Settings.SaveLogFiles.Detailed `
                -ScriptBlock {
                param(
                    $compDto, $scriptPaths, 
                    $sessionConfig, $maxConc, $detailedLog
                )

                # Raise the client-side deserialization cap. By default a
                # remote result larger than 200MB (MaximumReceivedObjectSize =
                # 209715200) makes Invoke-Command fail with 'The current
                # deserialized object size ... exceeded the allowed maximum
                # object size'. SetPermissions.ps1 can return very large
                # reporting lists (e.g. every incorrectly-permissioned path) on
                # big shares, so lift the limit to the Int32 maximum (~2GB).
                $sessionOption = New-PSSessionOption `
                    -MaximumReceivedObjectSize ([Int32]::MaxValue)

                $innerResults = [System.Collections.Generic.List[object]]::new()

                # Retry policy for TRANSIENT WinRM transport aborts only
                # (Win32 995 ERROR_OPERATION_ABORTED, "The I/O operation has
                # been aborted because of either a thread exit or an
                # application request"). These are client-side session
                # teardowns, not permission failures, and are safe to retry:
                # SetPermissions.ps1 enforces declarative desired-state ACLs,
                # so a re-run converges to the same result. Genuine business
                # errors are NOT retried (recorded on the first failure).
                #
                # EXPLICIT SESSIONS (not implicit -ComputerName): a client-side
                # abort does NOT guarantee the remote pipeline stopped -- WinRM
                # keeps the server-side shell alive until its IdleTimeout, so a
                # naive retry could start a SECOND SetPermissions run while the
                # first is still orphaned-but-running, racing on the same tree
                # (folder creation, SetAccessControl, ownership takeover). By
                # owning the PSSession we can Remove-PSSession in the `finally`,
                # which sends the WSMan terminate/delete that stops any orphaned
                # remote command BEFORE we open a fresh session for the retry.
                $maxAttempts = 3
                $retryDelaySeconds = 5

                # TryDequeue is atomic, so several workers can safely share one
                # computer's queue. A worker that finds it empty simply returns.
                $job = $null

                while ($compDto.JobQueue.TryDequeue([ref]$job)) {
                    $startTime = Get-Date
                    $attempt = 0
                    $jobResult = $null

                    while ($true) {
                        $attempt++
                        $session = $null
                        $needsRetry = $false

                        try {
                            # Wrap the whole if in @() so $restoredMatrix is
                            # ALWAYS a real array. Assigning the result of an if
                            # enumerates it: an empty branch collapses to $null
                            # and a single-element branch collapses to a bare
                            # scalar -- and $null cannot bind to the remote's
                            # mandatory [PSCustomObject[]]$Matrix parameter.
                            $restoredMatrix = @(
                                if (
                                    -not [string]::IsNullOrWhiteSpace($job.MatrixJson)
                                ) {
                                    $job.MatrixJson | ConvertFrom-Json
                                }
                            )

                            $session = New-PSSession `
                                -ComputerName $job.ComputerName `
                                -ConfigurationName $sessionConfig `
                                -SessionOption $sessionOption `
                                -ErrorAction Stop

                            $res = Invoke-Command `
                                -Session $session `
                                -FilePath $scriptPaths.SetPermissions `
                                -ArgumentList $job.Path, $job.Action, $restoredMatrix, $maxConc.FoldersPerMatrix, $detailedLog `
                                -ErrorAction Stop

                            $jobResult = [PSCustomObject]@{ 
                                ID     = $job.ID
                                Result = $res
                                Start  = $startTime 
                                End    = (Get-Date) 
                            }
                        }
                        catch {
                            $isTransientAbort = (
                                "$($_.Exception.Message)" -match 'I/O operation has been aborted' -or
                                ($_.Exception.HResult -band 0xFFFF) -eq 995
                            )

                            if ($isTransientAbort -and $attempt -lt $maxAttempts) {
                                $needsRetry = $true
                            }
                            else {
                                $errObj = [PSCustomObject]@{ 
                                    DateTime    = Get-Date
                                    Type        = 'FatalError' 
                                    Name        = 'Set permissions'
                                    Description = 'Failed applying action.' 
                                    Value       = $_ 
                                }
                                $jobResult = [PSCustomObject]@{
                                    ID     = $job.ID
                                    Result = $errObj
                                    Start  = $startTime
                                    End    = (Get-Date) 
                                }
                            }
                        }
                        finally {
                            # Always close the session. On a transient abort
                            # this forcibly terminates any still-running
                            # (orphaned) remote command server-side, so the
                            # retry never overlaps the previous run.
                            if ($session) {
                                Remove-PSSession `
                                    -Session $session `
                                    -ErrorAction SilentlyContinue
                            }
                        }

                        if ($needsRetry) {
                            # Settle delay: let the server finish tearing down
                            # the terminated command before we reconnect.
                            Start-Sleep -Seconds $retryDelaySeconds
                            continue
                        }

                        break
                    }

                    $innerResults.Add($jobResult)
                }
                return $innerResults
            }

            # Index matrices by ID so each result applies in O(1) instead of
            # rescanning $matricesToExecute for every returned job.
            $matrixById = @{}
            foreach ($m in $matricesToExecute) {
                $matrixById[$m.ID] = $m
            }

            # Main Thread Application: Add Job Times and Results back to Live Objects
            foreach ($resArray in $permResults) {
                foreach ($res in $resArray) {
                    $liveMatrix = $matrixById[$res.ID]
                    if ($liveMatrix) {
                        if ($res.Result) {
                            $structured = @($res.Result | ConvertTo-StructuredObjectHC)
                            
                            foreach ($entry in $structured) {
                                $liveMatrix.Check.Add($entry)
                            }
                        }
                        $liveMatrix.JobTime.Start = $res.Start
                        $liveMatrix.JobTime.End = $res.End
                        $liveMatrix.JobTime.Duration = New-TimeSpan -Start $res.Start -End $res.End 
                    }
                }
            }
        }
        #endregion

        return $Context

    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Category 'Runtime' `
            -Name 'PROCESS stage failure' `
            -Message "Unhandled exception occurred: $_" `
            -SystemErrors $SystemErrors 
        return $Context
    }
}