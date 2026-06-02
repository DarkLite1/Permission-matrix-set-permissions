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
            if (Test-ItemHasFatalErrorHC -CheckList $file.Check) {
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
        $throttleComputers = if (
            [string]::IsNullOrWhiteSpace($Context.Config.MaxConcurrent.Computers)
        ) {
            10
        }
        else {
            $Context.Config.MaxConcurrent.Computers
        }
        
        
        $psSessionConfig = if (
            [string]::IsNullOrWhiteSpace($Context.Config.Settings.PSSessionConfiguration)
        ) {
            'PowerShell.7'
        }
        else {
            $Context.Config.Settings.PSSessionConfiguration
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
                -ThrottleLimit $throttleComputers `
                -ArgumentList $Context.ScriptPath, $psSessionConfig `
                -ScriptBlock {
                param($dto, $scriptPaths, $sessionConfig)
                try {
                    $result = Invoke-Command `
                        -FilePath $scriptPaths.TestRequirements `
                        -ArgumentList $dto.PathsToCheck, $true `
                        -ConfigurationName $sessionConfig `
                        -ComputerName $dto.ComputerName `
                        -ErrorAction Stop
                
                    return [PSCustomObject]@{ 
                        ComputerName = $dto.ComputerName 
                        Result       = $result 
                    }
                }
                catch {
                    $errObj = [PSCustomObject]@{ 
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

            # Main Thread Application: Add results back to the live objects
            foreach ($output in $reqResults) {
                if ($output.Result) {
                    $targetMatrices = $validMatrices.Where(
                        { $_.Setting.Formatted.ComputerName -eq $output.ComputerName }
                    )

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

        # DTO FLATTENING: Protects deep properties from runspace truncation 
        $safePermGroups = foreach ($group in $compGroupsForPerms) {
            [PSCustomObject]@{
                ComputerName = $group.Name
                Matrices     = @(
                    foreach ($S in $group.Group) {
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
                    }
                )
            }
        }

        if ($safePermGroups) {
            $permResults = Invoke-WithOptionalParallelismHC `
                -InputObject $safePermGroups `
                -ThrottleLimit $throttleComputers `
                -ArgumentList $Context.ScriptPath, $psSessionConfig, $Context.Config.MaxConcurrent, $Context.Config.Settings.SaveLogFiles.Detailed `
                -ScriptBlock {
                param(
                    $compDto, $scriptPaths, 
                    $sessionConfig, $maxConc, $detailedLog
                )

                $innerResults = @()
            
                foreach ($job in $compDto.Matrices) {
                    $startTime = Get-Date
                    try {
                        $restoredMatrix = if (
                            -not [string]::IsNullOrWhiteSpace($job.MatrixJson)
                        ) {
                            @($job.MatrixJson | ConvertFrom-Json) 
                        }
                        else { @() } 
                    
                        $res = Invoke-Command `
                            -FilePath $scriptPaths.SetPermissions `
                            -ArgumentList $job.Path, $job.Action, $restoredMatrix, $maxConc.FoldersPerMatrix, $detailedLog `
                            -ConfigurationName $sessionConfig `
                            -ComputerName $job.ComputerName `
                            -ErrorAction Stop
                    
                        $innerResults += [PSCustomObject]@{ 
                            ID     = $job.ID
                            Result = $res
                            Start  = $startTime 
                            End    = (Get-Date) 
                        }
                    }
                    catch {
                        $errObj = [PSCustomObject]@{ 
                            Type        = 'FatalError' 
                            Name        = 'Set permissions'
                            Description = 'Failed applying action.' 
                            Value       = $_ 
                        }
                        $innerResults += [PSCustomObject]@{
                            ID     = $job.ID
                            Result = $errObj
                            Start  = $startTime
                            End    = (Get-Date) 
                        }
                    }
                }
                return $innerResults
            }

            # Main Thread Application: Add Job Times and Results back to Live Objects
            foreach ($resArray in $permResults) {
                foreach ($res in $resArray) {
                    $liveMatrix = $matricesToExecute.Where(
                        { $_.ID -eq $res.ID }, 'First'
                    ) | Select-Object -First 1
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