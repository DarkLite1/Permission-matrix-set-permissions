function ConvertTo-WorkItemsHC {
    <#
        Converts the imported matrix objects into safe DTOs for parallel runspaces.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$ImportedMatrix
    )

    $workItems = @()

    foreach ($I in $ImportedMatrix) {
        $WI = [ordered]@{
            MatrixFileName = $I.File.Item.Name
            LogFolder      = $I.File.LogFolder
            FileMeta       = $I.File

            Permissions    = $I.Permissions.Import
            FormData       = $I.FormData.Import

            Settings       = @()
        }

        foreach ($S in $I.Settings) {
            $WI.Settings += [ordered]@{
                Import         = $S.Import
                Matrix         = $S.Matrix
                Check          = @()
                PreAdChecks    = @()
                AdIdentifiers  = @()
                AdChecks       = @()
                ExpandedChecks = @()
            }
        }

        $workItems += $WI
    }

    return $workItems
}

function Get-AllAdIdentifiersHC {
    <#
        Collects all AD identifiers across all matrices before AD lookup.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$ImportedMatrix
    )

    $set = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($I in $ImportedMatrix) {

        # Column headers
        $headers = $I.Permissions.Import | Select-Object -First 3
        foreach ($h in $headers) {
            if ($h.P2) { $null = $set.Add($h.P2) }
        }

        # Settings
        foreach ($S in $I.Settings) {
            if ($S.Import.GroupName) { $null = $set.Add($S.Import.GroupName) }
            if ($S.Import.SiteCode) { $null = $set.Add($S.Import.SiteCode) }
        }

        # MatrixResponsible
        if ($I.FormData.Import -and $I.FormData.Import.MatrixResponsible) {
            foreach ($n in $I.FormData.Import.MatrixResponsible.Split(',')) {
                if ($n.Trim()) { $null = $set.Add($n.Trim()) }
            }
        }
    }

    return $set.ToArray()
}

function Invoke-MatrixPhase1ParallelHC {
    <#
        Phase 1:
            - File Check
            - Permissions Check
            - FormData Check
            - Settings Pre-AD Check
            - Build AD Identifiers
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$WorkItems,

        [int]$Throttle = 8
    )

    $block = {
        param($WorkItem)

        # File
        $WorkItem.FileChecks = Test-MatrixFileHC -MatrixObject $WorkItem

        # Permissions
        $perm = Test-MatrixPermissionsHC -Permissions $WorkItem.Permissions
        if ($perm) { $WorkItem.PermissionChecks = $perm }

        # FormData
        $fd = Test-MatrixFormDataHC -FormData $WorkItem.FormData
        if ($fd) { $WorkItem.FormDataChecks = $fd }

        # Settings
        for ($i = 0; $i -lt $WorkItem.Settings.Count; $i++) {

            $S = $WorkItem.Settings[$i]
            $S.PreAdChecks = Test-MatrixSettingHC -Setting $S.Import

            if ($S.PreAdChecks -and $S.PreAdChecks.Type -contains 'FatalError') { continue }

            # Build AD identifiers using the original rules
            $params = @{
                Begin         = $S.Import.GroupName
                Middle        = $S.Import.SiteCode
                ColumnHeaders = $WorkItem.Permissions | Select-Object -First 3
            }
            $S.AdIdentifiers = ConvertTo-MatrixADNamesHC @params
        }

        return $WorkItem
    }

    return $WorkItems |
    ForEach-Object -Parallel $block -ThrottleLimit $Throttle
}

function Invoke-MatrixPhase2ParallelHC {
    <#
        Phase 2:
            - AD validation per settings row
            - Expanded Matrix validation
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$WorkItems,

        [Parameter(Mandatory)]
        $AdInfo,

        [Parameter(Mandatory)]
        $DefaultAcl,

        [Parameter(Mandatory)]
        [string[]]$ExcludedSamAccountName,

        [int]$Throttle = 8
    )

    $block = {
        param(
            $WorkItem,
            $AdInfo,
            $DefaultAcl,
            $ExcludedSamAccountName
        )

        for ($i = 0; $i -lt $WorkItem.Settings.Count; $i++) {

            $S = $WorkItem.Settings[$i]

            # Skip if pre-AD fatal
            if ($S.PreAdChecks -and $S.PreAdChecks.Type -contains 'FatalError') { continue }

            # AD Validation
            $ac = Test-AdObjectsHC -ADObjects $S.AdIdentifiers -AdInfo $AdInfo
            if ($ac) { $S.AdChecks = $ac }

            if ($S.AdChecks -and $S.AdChecks.Type -contains 'FatalError') { continue }

            # Expanded Matrix Check
            if ($S.Matrix) {
                $exp = Test-ExpandedMatrixHC `
                    -Matrix $S.Matrix `
                    -ADObject $AdInfo `
                    -DefaultAcl $DefaultAcl `
                    -ExcludedSamAccountName $ExcludedSamAccountName

                if ($exp) { $S.ExpandedChecks = $exp }
            }
        }

        return $WorkItem
    }

    return $WorkItems |
    ForEach-Object -Parallel $block `
        -ThrottleLimit $Throttle `
        -ArgumentList $AdInfo, $DefaultAcl, $ExcludedSamAccountName
}

function Merge-CheckResultsHC {
    <#
        Reintegrates parallel check results back into the original matrix structure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$ImportedMatrix,

        [Parameter(Mandatory)]
        [array]$Phase1,

        [Parameter(Mandatory)]
        [array]$Phase2
    )

    for ($i = 0; $i -lt $ImportedMatrix.Count; $i++) {

        $M = $ImportedMatrix[$i]
        $P1 = $Phase1[$i]
        $P2 = $Phase2[$i]

        # File
        $M.File.Check += $P1.FileChecks

        # Permissions
        $M.Permissions.Check += $P1.PermissionChecks

        # FormData
        $M.FormData.Check += $P1.FormDataChecks

        # Settings
        for ($j = 0; $j -lt $M.Settings.Count; $j++) {

            $S = $M.Settings[$j]
            $S1 = $P1.Settings[$j]
            $S2 = $P2.Settings[$j]

            if ($S1.PreAdChecks) { $S.Check += $S1.PreAdChecks }
            if ($S2.AdChecks) { $S.Check += $S2.AdChecks }
            if ($S2.ExpandedChecks) { $S.Check += $S2.ExpandedChecks }
        }
    }

    return $ImportedMatrix
}

function Invoke-MatrixChecksHC {
    <#
        High-level orchestrator function for all check logic.
        Steps:
            1. Convert to safe DTOs
            2. Phase 1 parallel (no AD)
            3. Single AD lookup
            4. Phase 2 parallel (with AD)
            5. Merge results
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$ImportedMatrix,

        [Parameter(Mandatory)]
        $DefaultAcl,

        [Parameter(Mandatory)]
        [string[]]$ExcludedSamAccountName,

        [int]$Throttle = 8
    )

    # 1. Convert DTOs
    $WorkItems = ConvertTo-WorkItemsHC -ImportedMatrix $ImportedMatrix

    # 2. Phase 1 parallel
    $Phase1 = Invoke-MatrixPhase1ParallelHC `
        -WorkItems $WorkItems `
        -Throttle $Throttle

    # 3. Single AD lookup
    $AllIDs = Get-AllAdIdentifiersHC -ImportedMatrix $ImportedMatrix
    $AdInfo = Get-ADObjectDetailHC -ADObjectName $AllIDs -Type SamAccountName

    # 4. Phase 2 parallel
    $Phase2 = Invoke-MatrixPhase2ParallelHC `
        -WorkItems $Phase1 `
        -AdInfo $AdInfo `
        -DefaultAcl $DefaultAcl `
        -ExcludedSamAccountName $ExcludedSamAccountName `
        -Throttle $Throttle

    # 5. Merge
    return Merge-CheckResultsHC `
        -ImportedMatrix $ImportedMatrix `
        -Phase1 $Phase1 `
        -Phase2 $Phase2
}


