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
        if (-not $Context.Matrices -or $Context.Matrices.Count -eq 0) {
            return $Context
        }

        # Retrieve all specific Settings blocks across all matrices that don't have a FatalError yet
        # (Assuming Get-ExecutableMatrixHC essentially does this filtering)
        $executableSettings = @()
        foreach ($matrix in $Context.Matrices) {
            if ($matrix.Check.Type -notcontains 'FatalError') {
                $executableSettings += @($matrix.Settings | Where-Object { $_.Check.Type -notcontains 'FatalError' })
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
                    $targetSettings = $executableSettings.Where({ $_.Import.ComputerName -eq $output.ComputerName })
                    foreach ($setting in $targetSettings) {
                        $setting.Check += $output.Result | ConvertTo-StructuredObjectHC 
                    }
                }
            }
        }

        # =====================================================================
        # 2. PARALLEL: Set Permissions
        # =====================================================================
        
        # Re-filter matrices since some might have failed 'Test requirements.ps1'
        $validSettings = $executableSettings.Where({ $_.Check.Type -notcontains 'FatalError' })

        if ($validSettings.Count -eq 0) { return $Context }

        # Add default permissions just before execution
        if ($Context.Defaults.DefaultAcl.Count -gt 0) {
            foreach ($acl in $validSettings.Matrix.ACL.Where({ $_.Count -gt 0 })) {
                $Context.Defaults.DefaultAcl.GetEnumerator().Where({ -not $acl.ContainsKey($_.Key) }).ForEach({ $acl.Add($_.Key, $_.Value) }) 
            }
        }

        $compGroupsForPerms = $validSettings | Group-Object -Property { $_.Import.ComputerName }

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
                        $restoredMatrix = if (-not [string]::IsNullOrWhiteSpace($job.MatrixJson)) { @($job.MatrixJson | ConvertFrom-Json) } else { @() } 
                    
                        $res = Invoke-Command -FilePath $scriptPaths.SetPermissionFile `
                            -ArgumentList $job.Path, $job.Action, $restoredMatrix, $maxConc.FoldersPerMatrix, $detailedLog `
                            -ConfigurationName $sessionConfig `
                            -ComputerName $job.ComputerName `
                            -ErrorAction Stop
                    
                        $innerResults += [PSCustomObject]@{ ID = $job.ID; Result = $res; Start = $startTime; End = (Get-Date) }
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
                    $liveSetting = $validSettings.Where({ $_.ID -eq $res.ID }) 
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