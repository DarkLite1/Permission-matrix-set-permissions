function Invoke-PermissionMatrixProcessHC {
    <#
    .SYNOPSIS
        PROCESS stage for the Permission Matrix pipeline.
    .DESCRIPTION
        1. Parallel (Grouped by Computer): Run 'Test requirements.ps1'
        2. Parallel (Grouped by Computer): Run 'Set permissions.ps1' on servers without FatalErrors.
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

        $executableSettings = [System.Collections.Generic.List[pscustomobject]]::new()

        foreach ($file in $Context.FileResults) {
            if (Test-ItemHasFatalErrorHC -CheckList $file.Check) {
                continue
            }
            
            foreach ($matrixObj in $file.Matrices) {
                if (
                    -not (Test-ItemHasFatalErrorHC `
                            -CheckList $matrixObj.Check)
                ) {
                    $executableSettings.Add($matrixObj)
                }
            }
        }

        if ($executableSettings.Count -eq 0) {
            Write-Verbose 'No executable matrices found after initial validation.'
            return $Context
        }

        $throttleComputers = $Context.Config.MaxConcurrent.Computers ?? 10
        $psSessionConfig = $Context.Config.Settings.PSSessionConfiguration ?? 'PowerShell.7'

        # =====================================================================
        # 1. PARALLEL: Test Requirements
        # =====================================================================
        $matrixGroups = $executableSettings | Group-Object -Property { $_.Import.ComputerName }
        
        # DTO FLATTENING: Protects deep properties from runspace truncation 
        $safeReqGroups = foreach ($group in $matrixGroups) {
            [PSCustomObject]@{
                ComputerName = $group.Name
                PathsToCheck = @($group.Group.Import.Path)
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
                        -FilePath $scriptPaths.TestRequirementsFile `
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
                    $targetSettings = $executableSettings.Where(
                        { $_.Import.ComputerName -eq $output.ComputerName }
                    )

                    foreach ($setting in $targetSettings) {
                        $setting.Check += $output.Result | ConvertTo-StructuredObjectHC 
                    }
                }
            }
        }

        # =====================================================================
        # 2. PARALLEL: Set Permissions
        # =====================================================================
        
        $validSettings = $executableSettings.Where(
            { $_.Check.Type -notcontains 'FatalError' }
        )

        if ($validSettings.Count -eq 0) { return $Context }

        $compGroupsForPerms = $validSettings |
        Group-Object -Property { $_.Import.ComputerName }

        # DTO FLATTENING: Build a shallow array and wrap the deep Matrix array in JSON 
        $safePermGroups = foreach ($group in $compGroupsForPerms) {
            [PSCustomObject]@{
                ComputerName = $group.Name
                Matrices     = @(
                    foreach ($S in $group.Group) {
                        [PSCustomObject]@{
                            ID           = $S.ID
                            ComputerName = $S.Import.ComputerName
                            Path         = $S.Import.Path
                            Action       = $S.Import.Action
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
                            -FilePath $scriptPaths.SetPermissionFile `
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
                    $liveSetting = $validSettings.Where(
                        { $_.ID -eq $res.ID }, 'First'
                    ) 
                    if ($liveSetting) {
                        if ($res.Result) {
                            $liveSetting.Check += $res.Result | ConvertTo-StructuredObjectHC 
                        }
                        $liveSetting.JobTime.Start = $res.Start
                        $liveSetting.JobTime.End = $res.End
                        $liveSetting.JobTime.Duration = New-TimeSpan -Start $res.Start -End $res.End 
                    }
                }
            }
        }

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