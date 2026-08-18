#Requires -Version 7
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Scans an NTFS folder structure to create, check, or fix file system
    permissions.

.DESCRIPTION
    This script iterates through a specified directory tree and strictly
    enforces NTFS permissions based on a provided matrix of Access Control
    Lists (ACLs). It handles explicit permissions for matrix-defined folders
    and enforces strict inheritance rules for unlisted subfolders and files.

    It leverages a custom C# `TokenManipulator` class to temporarily grant the
    PowerShell process SeRestorePrivilege, SeBackupPrivilege, and
    SeTakeOwnershipPrivilege. This allows the script to forcefully correct ACLs
    and reclaim ownership even on folders where the Administrator currently
    receives an "Access Denied" error.

.PARAMETER Path
    The absolute path to the parent folder (local to the machine executing the
    script) where the folder tree begins.

.PARAMETER Action
    The execution mode.
    - New
        Creates the missing folder structure and applies the correct explicit
        permissions.
    - Check
        Audits the current permissions and reports discrepancies without
        modifying them.
    - Fix
        Audits the permissions, automatically corrects discrepancies, and
        forces ownership if access is denied.

.PARAMETER Matrix
    An array of PSCustomObjects containing the structured folder paths,
    inheritance flags, and their corresponding ACL hashtables.

.PARAMETER JobThrottleLimit
    The maximum number of concurrent runspaces to use when processing inherited
    folder and file permissions in parallel.

.PARAMETER DetailedLog
    If $true, captures the exact SDDL (Security Descriptor Definition Language)
    strings for both the old/incorrect permissions and the new/expected
    permissions, along with the matrix column headers.
    Note: Enabling this increases memory usage and reduces overall performance.
    Use primarily for troubleshooting.
#>

[OutputType([PSCustomObject[]])]
[CmdLetBinding()]
param (
    [Parameter(Mandatory)]
    [String]$Path,
    [Parameter(Mandatory)]
    [ValidateSet('New', 'Check', 'Fix')]
    [String]$Action,
    [Parameter(Mandatory)]
    [PSCustomObject[]]$Matrix,
    [Parameter(Mandatory)]
    [Int]$JobThrottleLimit,
    [Boolean]$DetailedLog,
    [Boolean]$CollectTestedPaths = $false
)

begin {
    #region Function New-AceHC
    function New-AceHC {
        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [ValidateSet('L', 'R', 'W', 'F', 'M')]
            [String]$Access,

            [Parameter(Mandatory)]
            [String]$Name,

            [Parameter(Mandatory)]
            [ValidateSet('Folder', 'InheritedFile', 'InheritedFolder')]
            [String]$Type
        )

        # Accept either a SID string ("S-1-5-...") or a bare SamAccountName.
        # SIDs come from the orchestrator's pre-resolved AD lookup and are
        # domain-portable; bare names get the legacy local-domain prefix for
        # backwards compatibility with any caller still passing SamAccountNames.
        if ($Name -match '^S-\d-\d+(-\d+)+$') {
            $identity = [System.Security.Principal.SecurityIdentifier]::new($Name)
        }
        else {
            $identity = [System.Security.Principal.NTAccount]::new("$env:USERDOMAIN\$Name")
        }

        $allow = [System.Security.AccessControl.AccessControlType]::Allow
        $rules = [System.Collections.Generic.List[System.Security.AccessControl.FileSystemAccessRule]]::new()

        $createRule = {
            param($rights, $inheritance, $propagation)
            $rules.Add([System.Security.AccessControl.FileSystemAccessRule]::new($identity, $rights, $inheritance, $propagation, $allow))
        }

        switch ($Access) {
            'L' {
                if ($Type -in 'Folder', 'InheritedFolder') {
                    &$createRule 'ReadAndExecute' 'ContainerInherit' 'None'
                }
            }
            'W' {
                if ($Type -eq 'Folder') {
                    &$createRule 'CreateFiles, AppendData, DeleteSubdirectoriesAndFiles, ReadAndExecute, Synchronize' 'None' 'None'
                    &$createRule 'DeleteSubdirectoriesAndFiles, Modify, Synchronize' 'ContainerInherit, ObjectInherit' 'InheritOnly'
                }
                elseif ($Type -eq 'InheritedFolder') {
                    &$createRule 'DeleteSubdirectoriesAndFiles, Modify, Synchronize' 'ContainerInherit, ObjectInherit' 'InheritOnly'
                }
                elseif ($Type -eq 'InheritedFile') {
                    &$createRule 'DeleteSubdirectoriesAndFiles, Modify, Synchronize' 'None' 'None'
                }
            }
            default {
                $rights = switch ($Access) {
                    'R' { 'ReadAndExecute' }
                    'F' { 'FullControl' }
                    'M' { 'Modify' }
                }

                if ($Type -in 'Folder', 'InheritedFolder') {
                    &$createRule $rights 'ContainerInherit, ObjectInherit' 'None'
                }
                elseif ($Type -eq 'InheritedFile') {
                    &$createRule $rights 'None' 'None'
                }
            }
        }
        return $rules.ToArray()
    }
    #endregion

    #region Function ConvertTo-HashtableHC (Main Thread)
    function ConvertTo-HashtableHC {
        # Rebuild a Deserialized.PSCustomObject (from a remoting/serialization
        # boundary) into a real hashtable so .Keys/.Count work; pass through a
        # $null or an existing dictionary unchanged.
        param($InputObject)

        if (($null -eq $InputObject) -or ($InputObject -is [System.Collections.IDictionary])) {
            return $InputObject
        }

        $hash = @{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            if ($prop.MemberType -match 'NoteProperty') { $hash[$prop.Name] = $prop.Value }
        }
        $hash
    }
    #endregion

    #region Function ConvertTo-MatrixAdObjectHC (Main Thread)
    function ConvertTo-MatrixAdObjectHC {
        <#
        .SYNOPSIS
            Build the human-readable 'MatrixFileAcl' array for the detail JSON.

        .DESCRIPTION
            Combines the matrix AD objects with their requested permission so the
            detail report shows both who was granted access and what was
            requested, e.g. 'GROUPHC\Group 1  List'. The identity is the display
            form of the SID (DOMAIN\name when it translates, the raw SID when it
            does not) so it lines up with the 'OldAcl'/'NewAcl' entries. The permission
            character (L/R/W/F/M) is mapped to its friendly word. Entries are
            sorted for stable output.

        .PARAMETER Names
            Hashtable keyed by SID (or SamAccountName) whose values are the matrix
            author labels. Its keys drive which objects are reported.

        .PARAMETER Permissions
            Hashtable keyed by the same SIDs whose values are the requested
            permission characters (L/R/W/F/M). Optional; when absent only the
            identity is emitted.
        #>
        param(
            [Parameter(Mandatory)]
            $Names,
            $Permissions
        )

        $permWord = @{
            'L' = 'List'
            'R' = 'Read'
            'W' = 'Write'
            'F' = 'FullControl'
            'M' = 'Modify'
        }

        $items = foreach ($sid in $Names.Keys) {
            $displayKey = try {
                ([System.Security.Principal.SecurityIdentifier]::new($sid)).
                Translate([System.Security.Principal.NTAccount]).Value
            }
            catch { $sid }

            $char = if ($Permissions) { "$($Permissions[$sid])".Trim().ToUpper() } else { '' }
            $type = if ($permWord.ContainsKey($char)) { $permWord[$char] } else { $char }

            if ($type) { "$displayKey  $type" } else { $displayKey }
        }

        , @($items | Sort-Object)
    }
    #endregion

    #region Function New-UnreadableAclEntryHC (Main Thread)
    function New-UnreadableAclEntryHC {
        # Build the DetailedLog entry for a path whose ACL could not be read.
        # 'OldAcl' carries the failure reason; MatrixFileAcl is added only when
        # the matrix labels are known so the user can map it to the Excel columns.
        param(
            [Parameter(Mandatory)]
            [String]$Reason,
            $AdNames,
            $AdPermissions
        )

        $entry = [ordered]@{
            'OldAcl' = @("ACL could not be read: $Reason")
        }
        if ($AdNames -and $AdNames.Count -gt 0) {
            $entry['MatrixFileAcl'] = ConvertTo-MatrixAdObjectHC -Names $AdNames -Permissions $AdPermissions
        }
        $entry
    }
    #endregion

    #region Function Get-DirectoryAclSafeHC (Main Thread)
    function Get-DirectoryAclSafeHC {
        # Read a directory's ACL via the fast .NET API, falling back to Get-Acl,
        # and report the outcome as flags so the caller owns the logging and
        # collection routing. NOT used by the hot child-walker (kept inline there
        # to avoid per-item call overhead on large trees).
        param(
            [Parameter(Mandatory)]
            [System.IO.DirectoryInfo]$DirectoryInfo
        )

        $result = [PSCustomObject]@{
            Acl              = $null
            AccessDenied     = $false
            Removed          = $false
            UnreadableReason = $null
        }

        try {
            $result.Acl = [System.IO.FileSystemAclExtensions]::GetAccessControl($DirectoryInfo)
        }
        catch [System.UnauthorizedAccessException] {
            $result.AccessDenied = $true
        }
        catch {
            try {
                $result.Acl = Get-Acl -LiteralPath $DirectoryInfo.FullName -ErrorAction Stop
            }
            catch [System.UnauthorizedAccessException] {
                $result.AccessDenied = $true
            }
            catch {
                if (-not (Test-Path -LiteralPath $DirectoryInfo.FullName)) {
                    $result.Removed = $true
                    $Error.RemoveAt(0)
                }
                else {
                    $result.UnreadableReason = $_.Exception.Message
                }
            }
        }

        $result
    }
    #endregion

    #region Function Test-AclEqualHC (Main Thread)
    function Test-AclEqualHC {
        [OutputType([Boolean])]
        param (
            [Parameter(Mandatory)]
            [AllowNull()]
            [AllowEmptyCollection()]
            [System.Object[]]$ReferenceAce,

            [Parameter(Mandatory)]
            [AllowNull()]
            [AllowEmptyCollection()]
            [System.Object[]]$DifferenceAce
        )

        try {
            # Build deduplicated fingerprint sets for both sides and compare
            # with SetEquals. This mirrors the parallel-thread implementation and
            # is robust to duplicate ACEs (which Windows can merge on-disk). It
            # avoids two defects of a raw '.Count' guard + one-directional
            # Contains() check:
            #   - the false "not equal" when the reference collapses two
            #     fingerprint-identical ACEs into one but the on-disk ACL has
            #     both (or vice-versa), and
            #   - the false "equal" when both sides have matching counts yet the
            #     difference side is missing a reference ACE but repeats another.
            # The fingerprint stays propagation-blind on purpose: inherited ACEs
            # land with PropagationFlags=None while the matrix models them as
            # InheritOnly, and that difference must compare as equal.
            $refSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($R in $ReferenceAce) {
                # [int] casts bypass slow string evaluations
                [void]$refSet.Add("$([int]$R.FileSystemRights)|$([int]$R.AccessControlType)|$($R.IdentityReference.ToString())|$([int]$R.InheritanceFlags)")
            }

            $diffSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($D in $DifferenceAce) {
                [void]$diffSet.Add("$([int]$D.FileSystemRights)|$([int]$D.AccessControlType)|$($D.IdentityReference.ToString())|$([int]$D.InheritanceFlags)")
            }

            return $refSet.SetEquals($diffSet)
        }
        catch {
            throw "Failed testing the ACL for equality: $_"
        }
    }
    #endregion

    #region ScriptBlock InheritedPermissionsScriptBlock
    $inheritedPermissionsScriptBlock = {
        [OutputType([PSCustomObject[]])]
        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            [ValidateSet('Check', 'Fix')]
            [String]$Action,

            [Parameter(Mandatory)]
            [AllowNull()]
            [AllowEmptyCollection()]
            [System.Object[]]$FolderAclAccessList = @(),

            [Parameter(Mandatory)]
            [AllowNull()]
            [AllowEmptyCollection()]
            [System.Object[]]$FileAclAccessList = @(),

            [Parameter(Mandatory)]
            [HashTable]$IgnoredFolderPaths,
            [Parameter(Mandatory)]
            [String]$TokenPrivileges,

            [Parameter()]
            [PSObject]$AdNames,

            [Parameter()]
            [PSObject]$AdPermissions,

            [Boolean]$CheckSeedPath,
            [Boolean]$CheckInheritedOnly,
            [Boolean]$DetailedLog,
            [Boolean]$CollectTestedPaths = $false
        )

        $ErrorActionPreference = 'Stop'

        <#
         WHEN THIS JOB RAN, not just what it cost.

         Everything else in this record is a cost total summed across
         concurrent workers, which deliberately says nothing about elapsed
         time. That leaves the one question a partitioned parallel walk always
         raises unanswerable: a Settings row cannot finish before its SLOWEST
         job, so if one matrix folder holds most of the tree, the row is bound
         by that folder and the totals will not show it.

         Stopwatch::GetTimestamp() reads a process-wide monotonic counter and
         ForEach-Object -Parallel runs its runspaces in THIS process, so raw
         timestamps taken here are directly comparable between jobs and with
         the main thread's own start. They are therefore emitted raw and
         differenced on the main thread, which avoids passing the Settings-row
         start into every job just to subtract it here.

         Cost: two clock reads per JOB. Not per item.
        #>
        $jobStartTicks = [System.Diagnostics.Stopwatch]::GetTimestamp()

        #region Telemetry counters (Parallel Thread)
        <#
         Volume and cost counters for this runspace's slice of the walk.
         Merged by the main thread into one per-setting telemetry record.

         THREE RULES, ALL OF THEM MEASURED (see Tests/Operations/
         TelemetryOverhead.Tests.ps1 for the benchmark that produced them).

         1. HOIST THE LOOKUP.
            Reading '$telemetry' from a nested function is a DYNAMIC SCOPE
            lookup and costs ~1us EVERY TIME — far more than the increment
            itself, and it dwarfs the container type (a [long[]] measured no
            better than a hashtable). Every function and scriptblock below
            therefore copies the reference into a local '$t' once, then indexes
            the local. Measured: 2,949ns/item unhoisted vs 777ns/item hoisted.
            Do not "simplify" '$t' away.

         2. SAMPLE THE EXPENSIVE SIGNALS.
            Stopwatch::GetTimestamp() costs more through the PowerShell
            interpreter than the counters do, so timing every item would make
            the timing the dominant cost. Timing and the ACE census are taken
            on 1 item in $sampleEvery instead, which drops the per-item cost
            by roughly half.

            WARM-UP: the first $sampleWarmup items are sampled unconditionally
            before the 1-in-N rule takes over. Without it a small path yields
            too few samples for the mean to mean anything — 200 items at 1-in-64
            is 3 samples, and a 3-sample mean measured at 1.44x the true value
            in testing. The warm-up guarantees a usable floor on small trees and
            is reported as its own pool so it cannot skew large-tree means.
            Always check 'AclReadSamples' and 'AclReadBasis' before trusting
            'AclReadMsPerItem'.

         3. COUNT RARE EVENTS UNCONDITIONALLY.
            Denials, read failures, writes and incorrect items are exceptional
            on a converged tree, so they are counted on every occurrence — no
            sampling, no estimation. If they ever stop being rare, that is
            itself the finding.

         WHAT IS DELIBERATELY NOT COLLECTED
         No per-item strings, no per-item log lines, no per-item paths. That is
         what $CollectTestedPaths does, and it is the expensive pattern: a
         dictionary holding every path walked, marshalled back across the
         runspace boundary. Everything below is a scalar, so the whole record
         is a few hundred bytes no matter how large the tree is.
        #>

        # Power of two so the sampling test is a bitmask, not a modulo.
        # $sampleCounter is script-scoped to this runspace so the sample points
        # keep advancing across the recursive Get-FolderContentHC calls instead
        # of restarting (and re-sampling the first item) in every directory.
        $sampleEvery = 64
        $sampleMask = $sampleEvery - 1
        $sampleWarmup = 300
        $script:sampleCounter = 0

        <#
         IDENTITY CENSUS: A SEPARATE, MUCH RARER SAMPLE.

         Translating a security identifier into an account name is an LSA
         lookup. Windows caches it per machine, so the cost is driven by how
         many DISTINCT identities the tree uses - a few groups repeated across
         millions of items is nearly free, thousands of them thrash the cache -
         and by how many identities do not resolve at all. An entry left behind
         by a deleted account cannot be cached as a name and is the classic
         cause of an overnight regression appearing with no code change.

         Neither is visible in AceCountMean: twenty entries naming the same
         twenty groups everywhere and twenty entries naming a different twenty
         per folder look identical there.

         WHY 1-IN-1024 AND NOT 1-IN-64. This is the only telemetry that walks
         the ACE list, so it is the only telemetry whose cost scales with ACE
         count rather than item count. Measured on a 20-entry list, interleaved
         minimum-of-7 at 200k iterations: at the read timer's rate it costs 588
         ns per item (1.55x the previous shape) - more than every other counter
         put together. At 1-in-1024 it costs 113 ns (1.10x), which is 1.4
         seconds across the whole 2026_08_17 run.

         Cardinality does not need many samples. It needs to distinguish
         'twenty groups' from 'thousands', and the sparse rate still recovered
         all 20 distinct names from 495 sampled items in the same test. So it
         gets its own gate. 1024 is a multiple of 64, so every identity sample
         is also a read sample and the two stay in phase.

         CAPPED. Beyond $identityCap distinct names the answer is already
         'a lot'; counting further only grows what crosses the runspace
         boundary. The cap being hit is itself reported.
        #>
        $identityEvery = 1024
        $identityMask = $identityEvery - 1
        $identityCap = 512

        $telemetry = @{
            # The subtree these counters describe. Each parallel job walks ONE
            # matrix folder and skips any child that is itself a matrix folder,
            # so the jobs form a non-overlapping partition of the tree and this
            # label is unambiguous.
            #
            # Non-numeric, so the main thread's merge must not try to add it —
            # see the merge loop, which only sums keys it already holds.
            WalkedPath            = $Path
            SeedOnly              = $CheckInheritedOnly
            # Raw timestamps, not counters. Like WalkedPath these must not be
            # summed by the main thread's merge — see the merge loop, which
            # excludes them explicitly.
            JobStartTicks         = $jobStartTicks
            JobEndTicks           = 0L
            # The distinct identity names seen by this job, capped. A set, not
            # a counter: the main thread UNIONS these across jobs, because two
            # jobs each seeing twenty names may be seeing the same twenty or a
            # different twenty, and only the union can tell the difference.
            #
            # This is the one exception to 'no strings leave the walk', and it
            # is bounded by $identityCap rather than by the size of the tree.
            IdentitySet           = [System.Collections.Generic.HashSet[String]]::new()
            IdentityObservations  = 0L
            IdentityTruncated     = 0L
            AceUnresolvedSids     = 0L
            AceUnresolvedItems    = 0L
            FoldersWalked         = 0L
            FilesWalked           = 0L
            # Warm-up and stride samples are kept APART on purpose. The warm-up
            # covers the first N items CONTIGUOUSLY, which is the worst possible
            # basis for a mean on a large tree: those items are the coldest
            # (cache, JIT, first directory opens) and they represent a fraction
            # of a percent of the data. Mixed into one pool they dominated the
            # estimate and measured up to 1.9x the true value in testing.
            # Reported separately, the stride pool gives an unbiased mean for
            # large trees and the warm-up pool rescues small ones.
            AclReadWarmupSamples  = 0L
            AclReadWarmupTicks    = 0L
            AclReadStrideSamples  = 0L
            AclReadStrideTicks    = 0L
            # The two stages that sit BETWEEN the read and the verdict, and
            # that scale with the number of ACEs on an item rather than with
            # the number of items:
            #
            #   Project - materialising $acl.Access. This is not a property
            #             read. It builds the rule collection and asks Windows
            #             to translate every ACE's SID into an NTAccount, which
            #             is an LSA call that can leave the machine.
            #   Compare - turning those rules into fingerprints and testing
            #             them against the reference set.
            #
            # Both were previously outside every timer, so an item carrying 30
            # ACEs and an item carrying 3 cost the same in the telemetry while
            # differing by an order of magnitude on the clock. Same sample
            # decision and the same warm-up/stride split as the read timer, for
            # the reasons given above.
            AclProjectWarmupSamples = 0L
            AclProjectWarmupTicks   = 0L
            AclProjectStrideSamples = 0L
            AclProjectStrideTicks   = 0L
            AclCompareWarmupSamples = 0L
            AclCompareWarmupTicks   = 0L
            AclCompareStrideSamples = 0L
            AclCompareStrideTicks   = 0L
            AclReadDenied      = 0L
            AclReadFailed      = 0L
            AclWrites          = 0L
            AclWriteTicks      = 0L
            AclWriteDenied     = 0L
            IncorrectItems     = 0L
            AceCountTotal      = 0L
            AceCountItems      = 0L
            AceCountMax        = 0L
            EnumeratedDirs     = 0L
            EnumerateTicks     = 0L
            SampleEvery        = [long]$sampleEvery
            SampleWarmup       = [long]$sampleWarmup
        }
        #endregion

        #region Function ConvertTo-HashtableHC (Parallel Thread)
        # Duplicated from the main-thread definition because this scriptblock is
        # rehydrated in a fresh runspace that cannot see the parent's functions.
        function ConvertTo-HashtableHC {
            param($InputObject)

            if (($null -eq $InputObject) -or ($InputObject -is [System.Collections.IDictionary])) {
                return $InputObject
            }

            $hash = @{}
            foreach ($prop in $InputObject.PSObject.Properties) {
                if ($prop.MemberType -match 'NoteProperty') { $hash[$prop.Name] = $prop.Value }
            }
            $hash
        }
        #endregion

        #region Normalize AdNames/AdPermissions to real hashtables
        # Defensive: when this scriptblock is invoked with data that crossed a
        # remoting/serialization boundary, a nested hashtable arrives as a
        # Deserialized.PSCustomObject. Rebuild it so .Keys/.Count work below.
        $AdNames = ConvertTo-HashtableHC -InputObject $AdNames
        $AdPermissions = ConvertTo-HashtableHC -InputObject $AdPermissions
        #endregion

        #region Function ConvertTo-MatrixAdObjectHC (Parallel Thread)
        # Duplicated from the main-thread definition because this scriptblock is
        # rehydrated in a fresh runspace that cannot see the parent's functions.
        function ConvertTo-MatrixAdObjectHC {
            param(
                [Parameter(Mandatory)]
                $Names,
                $Permissions
            )

            $permWord = @{
                'L' = 'List'
                'R' = 'Read'
                'W' = 'Write'
                'F' = 'FullControl'
                'M' = 'Modify'
            }

            $items = foreach ($sid in $Names.Keys) {
                $displayKey = try {
                    ([System.Security.Principal.SecurityIdentifier]::new($sid)).
                    Translate([System.Security.Principal.NTAccount]).Value
                }
                catch { $sid }

                $char = if ($Permissions) { "$($Permissions[$sid])".Trim().ToUpper() } else { '' }
                $type = if ($permWord.ContainsKey($char)) { $permWord[$char] } else { $char }

                if ($type) { "$displayKey  $type" } else { $displayKey }
            }

            , @($items | Sort-Object)
        }
        #endregion

        #region Function New-UnreadableAclEntryHC (Parallel Thread)
        # Duplicated from the main-thread definition because this scriptblock is
        # rehydrated in a fresh runspace that cannot see the parent's functions.
        function New-UnreadableAclEntryHC {
            param(
                [Parameter(Mandatory)]
                [String]$Reason,
                $AdNames,
                $AdPermissions
            )

            $entry = [ordered]@{
                'OldAcl' = @("ACL could not be read: $Reason")
            }
            if ($AdNames -and $AdNames.Count -gt 0) {
                $entry['MatrixFileAcl'] = ConvertTo-MatrixAdObjectHC -Names $AdNames -Permissions $AdPermissions
            }
            $entry
        }
        #endregion

        #region Function Get-DirectoryAclSafeHC (Parallel Thread)
        # Duplicated from the main-thread definition because this scriptblock is
        # rehydrated in a fresh runspace that cannot see the parent's functions.
        # Used only by the seed-folder check (cold); the hot child-walker keeps
        # its cascade inline to avoid per-item call overhead on large trees.
        function Get-DirectoryAclSafeHC {
            param(
                [Parameter(Mandatory)]
                [System.IO.DirectoryInfo]$DirectoryInfo
            )

            $result = [PSCustomObject]@{
                Acl              = $null
                AccessDenied     = $false
                Removed          = $false
                UnreadableReason = $null
            }

            try {
                $result.Acl = [System.IO.FileSystemAclExtensions]::GetAccessControl($DirectoryInfo)
            }
            catch [System.UnauthorizedAccessException] {
                $result.AccessDenied = $true
            }
            catch {
                try {
                    $result.Acl = Get-Acl -LiteralPath $DirectoryInfo.FullName -ErrorAction Stop
                }
                catch [System.UnauthorizedAccessException] {
                    $result.AccessDenied = $true
                }
                catch {
                    if (-not (Test-Path -LiteralPath $DirectoryInfo.FullName)) {
                        $result.Removed = $true
                        $Error.RemoveAt(0)
                    }
                    else {
                        $result.UnreadableReason = $_.Exception.Message
                    }
                }
            }

            $result
        }
        #endregion

        try { Import-Module -Name 'Microsoft.PowerShell.Security' } catch { throw "Failed loading .NET library: $_" }

        # OPTIMIZATION: Setup HashSets ONCE per runspace to avoid repeating work for every file
        $folderRulesSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        if ($FolderAclAccessList) { foreach ($r in $FolderAclAccessList) { [void]$folderRulesSet.Add($r) } }

        $fileRulesSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        if ($FileAclAccessList) { foreach ($r in $FileAclAccessList) { [void]$fileRulesSet.Add($r) } }

        #region Function Test-AclEqualHC (Parallel Thread)
        function Test-AclEqualHC {
            [OutputType([Boolean])]
            param (
                [Parameter(Mandatory)]
                [System.Collections.Generic.HashSet[string]]$ReferenceSet,

                [Parameter(Mandatory)]
                [AllowNull()]
                [AllowEmptyCollection()]
                [System.Object[]]$DifferenceAce
            )

            try {
                if ($ReferenceSet.Count -eq 0) {
                    return ($DifferenceAce.Count -eq 0)
                }

                # Compare each on-disk fingerprint against the prebuilt
                # reference set and keep a deduplicated seen set for the final
                # SetEquals check. This preserves the propagation-blind duplicate
                # behavior while avoiding a second full reference set per item.
                $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($D in $DifferenceAce) {
                    # [int] casts bypass slow string evaluations
                    $fingerprint = "$([int]$D.FileSystemRights)|$([int]$D.AccessControlType)|$($D.IdentityReference.ToString())|$([int]$D.InheritanceFlags)"
                    if (-not $ReferenceSet.Contains($fingerprint)) {
                        return $false
                    }
                    [void]$seen.Add($fingerprint)
                }

                return $ReferenceSet.SetEquals($seen)
            }
            catch {
                throw "Failed testing the ACL for equality: $_"
            }
        }
        #endregion

        #region Function Test-AclInheritedOnlyHC
        function Test-AclInheritedOnlyHC {
            [OutputType([Boolean])]
            param (
                [Parameter(Mandatory)]
                [System.Security.AccessControl.FileSystemSecurity]$Acl
            )

            if ($Acl.AreAccessRulesProtected) {
                return $false
            }

            foreach ($rule in $Acl.Access) {
                if (-not $rule.IsInherited) {
                    return $false
                }
            }

            return $true
        }
        #endregion

        #region Function Get-FolderContentHC
        function Get-FolderContentHC {
            param (
                # The folder to process, as a DirectoryInfo. The top-level call
                # builds one from the parent path; recursive calls hand us the
                # already-enumerated child from the parent's
                # EnumerateFileSystemInfos, whose cached .Exists/.Attributes skip
                # a redundant metadata syscall per folder. Recursion only runs
                # for containers, so this is always a directory (never a file);
                # a non-directory path is still handled by the .Exists guard.
                [Parameter(Mandatory)]
                [System.IO.DirectoryInfo]$DirectoryInfo
            )

            $fullName = $DirectoryInfo.FullName

            # Hoist the parent-scope telemetry lookup into a local ONCE per
            # directory. See rule 1 in the counter notes: reading $telemetry
            # directly from inside this function is a dynamic-scope lookup
            # costing ~1us per access, which is more than the work it measures.
            # $t is a reference to the same hashtable, so mutating it mutates
            # the parent's object.
            $t = $telemetry

            if ($null -eq $t) {
                <#
                 DIAGNOSTICS MUST NEVER BE LOAD-BEARING.

                 This function is also dot-sourced standalone — the Pester suite
                 lifts it out of this script by AST and supplies only the
                 closure variables it needs — so $telemetry is not guaranteed to
                 exist. Without this guard, '$t['EnumerateTicks'] += ...' throws
                 'Cannot index into a null array', and because that sits inside
                 the enumeration try/catch it resurfaces as the misleading
                 'Failed retrieving the folder content of ...'.

                 Counting into a throwaway hashtable keeps the walk correct and
                 simply discards the numbers. A missing counter is a lost
                 measurement; a thrown counter is a failed permission run, and
                 those are not remotely the same cost.
                #>
                $t = @{}
                $sampleWarmup = 0
                $sampleMask = 0
            }

            try {
                # Perf: Write-Verbose "Get content of folder '$fullName'" removed
                # — it fired once per folder (millions of calls on large trees)
                # and the string interpolation + call overhead adds up even when
                # $VerbosePreference is SilentlyContinue.

                # Skip anything that is not a real, enumerable directory. The
                # matrix can list a name that exists on disk as a file, or as a
                # DFS link / reparse point / junction. DirectoryInfo.Exists is
                # $false for a file (or a missing path), and enumerating a DFS
                # link or reparse point throws 'The parameter is incorrect'.
                # Both cases have no inheritable children to process here, so
                # return gracefully instead of aborting the entire run with a
                # FatalError. This mirrors the child-level skip below.
                if (-not $DirectoryInfo.Exists) {
                    Write-Verbose "Skip '$fullName': not a directory (file or missing)"
                    return
                }
                if ($DirectoryInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                    Write-Verbose "Skip '$fullName': reparse point or DFS link"
                    return
                }

                # Telemetry: time the enumerator handle, not the iteration.
                # EnumerateFileSystemInfos is lazy, so this measures the
                # directory open only; the per-child cost lands in the ACL
                # read timer below. Splitting the two is the whole point —
                # it separates 'the tree got bigger / the disk got slower'
                # from 'the ACL operations got more expensive'.
                # Timed unsampled: this fires once per DIRECTORY, not per
                # item, so its cost is amortized over every child below it.
                $tsEnum = [System.Diagnostics.Stopwatch]::GetTimestamp()
                $enumerator = $DirectoryInfo.EnumerateFileSystemInfos()
                $t['EnumerateTicks'] += (
                    [System.Diagnostics.Stopwatch]::GetTimestamp() - $tsEnum
                )
                $t['EnumeratedDirs']++
            }
            catch {
                throw "Failed retrieving the folder content of '$fullName': $_"
            }

            foreach ($child in $enumerator) {
                # Skip DFS links, Reparse Points and System directories
                if (
                    ($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -or
                    ($child.Attributes -band [System.IO.FileAttributes]::System)
                    # ($child.Attributes -band [System.IO.FileAttributes]::Hidden)
                ) {
                    continue
                }

                if ($IgnoredFolderPaths.ContainsKey($child.FullName)) {
                    continue
                }

                $isContainer = $child -is [System.IO.DirectoryInfo]

                # Telemetry: the ONLY unconditional per-item cost. This is
                # what turns 'this path got slower' into 'this path got slower
                # AND grew by 40k files' — or did not grow at all, which is the
                # more interesting answer.
                if ($isContainer) { $t['FoldersWalked']++ }
                else { $t['FilesWalked']++ }

                # Sample decision, taken once and reused by the read timer and
                # the ACE census below so the branch is only evaluated once.
                # The warm-up term is first so that on small trees the -or
                # short-circuits before the bitmask is ever evaluated.
                $sampleCounter = ++$script:sampleCounter
                $isWarmupSample = ($sampleCounter -le $sampleWarmup)
                $isSample = (
                    $isWarmupSample -or
                    (($sampleCounter -band $sampleMask) -eq 0)
                )

                $accessDenied = $false
                $acl = $null
                $tsRead = if ($isSample) {
                    [System.Diagnostics.Stopwatch]::GetTimestamp()
                }
                else { 0L }
                try {
                    # FAST .NET API Call bypassing PowerShell provider overhead
                    if ($isContainer) {
                        $acl = [System.IO.FileSystemAclExtensions]::GetAccessControl([System.IO.DirectoryInfo]$child)
                    }
                    else {
                        $acl = [System.IO.FileSystemAclExtensions]::GetAccessControl([System.IO.FileInfo]$child)
                    }
                }
                catch [System.UnauthorizedAccessException] {
                    $accessDenied = $true
                    $t['AclReadDenied']++
                }
                catch {
                    # FALLBACK: Use classic Get-Acl if .NET method fails
                    try {
                        $acl = Get-Acl -LiteralPath $child.FullName -ErrorAction Stop
                    }
                    catch [System.UnauthorizedAccessException] {
                        $accessDenied = $true
                        $t['AclReadDenied']++
                    }
                    catch {
                        $t['AclReadFailed']++

                        if (-not (Test-Path -LiteralPath $child.FullName)) {
                            Write-Verbose "Item '$($child.FullName)' removed"
                            $Error.RemoveAt(0)
                        }
                        else {
                            $errorMessage = "Failed retrieving the ACL of '$child': $_"

                            Write-Warning $errorMessage

                            # The ACL could not be read at all (not access-denied,
                            # not removed), so the item was neither checked nor
                            # corrected. Report it under its own 'ACL could not be
                            # read' warning instead of the inherited-incorrect
                            # list, so the reason survives and the user knows the
                            # path needs manual attention.
                            if ($DetailedLog) {
                                $unreadableAcl[$child.FullName] = New-UnreadableAclEntryHC -Reason $_.Exception.Message -AdNames $AdNames -AdPermissions $AdPermissions
                            }
                            else {
                                $unreadableAcl.Add($child.FullName)
                            }
                        }

                        # Closed here too: a 'continue' must not skip the
                        # timer, or the slowest reads (the failing ones) would
                        # be the only ones missing from the mean.
                        if ($isSample) {
                            $readTicks = [System.Diagnostics.Stopwatch]::GetTimestamp() - $tsRead
                            if ($isWarmupSample) {
                                $t['AclReadWarmupSamples']++
                                $t['AclReadWarmupTicks'] += $readTicks
                            }
                            else {
                                $t['AclReadStrideSamples']++
                                $t['AclReadStrideTicks'] += $readTicks
                            }
                        }
                        continue
                    }
                }

                if ($isSample) {
                    # Kept in a variable rather than discarded: this same
                    # instant closes the read window and opens the projection
                    # window below, so the next stage costs no extra clock read.
                    $tsAfterRead = [System.Diagnostics.Stopwatch]::GetTimestamp()
                    $readTicks = $tsAfterRead - $tsRead
                    if ($isWarmupSample) {
                        $t['AclReadWarmupSamples']++
                        $t['AclReadWarmupTicks'] += $readTicks
                    }
                    else {
                        $t['AclReadStrideSamples']++
                        $t['AclReadStrideTicks'] += $readTicks
                    }
                }

                if ($CollectTestedPaths) {
                    $testedInheritedFilesAndFolders[$child.FullName] = $true
                }

                <#
                 STAGE 2: PROJECTION.

                 '$acl.Access' is not a field access. It is
                 GetAccessRules($true, $true, [NTAccount]), which materialises
                 the whole rule collection AND translates every ACE's security
                 identifier into an account name. That translation is an LSA
                 lookup: cached per machine when it hits, a round trip to a
                 domain controller when it does not.

                 The cost therefore scales with the ACE count and with how many
                 DISTINCT identities the tree uses, neither of which the item
                 counters can see. Timing it separately is what turns 'this row
                 is slow and we do not know why' into a number.
                #>
                $diffAce = if (-not $accessDenied -and $acl) { @($acl.Access) } else { @() }

                <#
                 ONE SAMPLE BRANCH, THREE JOBS. Closing the projection window,
                 taking the ACE census and opening the comparison window are
                 all gated on the same $isSample, so they share a single test
                 instead of three.

                 Measured, interleaved minimum-of-7 at 200k iterations, census
                 present in both shapes: three separate gates cost 455 ns per
                 item over the previous shape; this merged form costs 97 ns
                 (1.09x). Across the whole 2026_08_17 BNL run that is 1.2
                 seconds for 12.2M items — 0.03% of the largest row.

                 The census is taken HERE, between the two windows, so it is
                 charged to neither. It is telemetry overhead; letting it land
                 inside the comparison window would make the instrumentation
                 inflate the field it exists to report.
                #>
                $tsCompare = 0L

                if ($isSample) {
                    $tsAfterProject = [System.Diagnostics.Stopwatch]::GetTimestamp()
                    $projectTicks = $tsAfterProject - $tsAfterRead

                    if ($isWarmupSample) {
                        $t['AclProjectWarmupSamples']++
                        $t['AclProjectWarmupTicks'] += $projectTicks
                    }
                    else {
                        $t['AclProjectStrideSamples']++
                        $t['AclProjectStrideTicks'] += $projectTicks
                    }

                    # Telemetry: ACE census. Reading $diffAce.Count is free — it
                    # is already materialized above for the comparison — but the
                    # three counter writes are not, so the census rides the same
                    # 1-in-N sample as the timers. Growth in the mean or max
                    # here, run over run, is what an ACL that is being appended
                    # to rather than replaced looks like.
                    if ($diffAce.Count) {
                        $t['AceCountTotal'] += $diffAce.Count
                        $t['AceCountItems']++
                        if ($diffAce.Count -gt $t['AceCountMax']) {
                            $t['AceCountMax'] = $diffAce.Count
                        }

                        <#
                         Identity census, on its own much rarer gate. Nested
                         inside the sample branch rather than tested separately
                         so non-sampled items never evaluate it at all, and
                         placed here, between the two timing windows, so the
                         walk over the ACE list is charged to neither stage.

                         An unresolved entry is detected by TYPE, not by
                         string-matching 'S-1-'. When Windows cannot resolve a
                         security identifier, .NET leaves the IdentityReference
                         as a SecurityIdentifier instead of an NTAccount, so the
                         type is the answer and no parsing is needed.
                        #>
                        if ($isWarmupSample -or (($sampleCounter -band $identityMask) -eq 0)) {
                            $identitySet = $t['IdentitySet']
                            $itemHasUnresolved = $false

                            foreach ($ace in $diffAce) {
                                $identity = $ace.IdentityReference

                                if ($identity -is [System.Security.Principal.SecurityIdentifier]) {
                                    $t['AceUnresolvedSids']++
                                    $itemHasUnresolved = $true
                                }

                                if ($identitySet.Count -lt $identityCap) {
                                    [void]$identitySet.Add($identity.Value)
                                }
                                else {
                                    $t['IdentityTruncated'] = 1L
                                }
                            }

                            $t['IdentityObservations'] += $diffAce.Count

                            if ($itemHasUnresolved) {
                                $t['AceUnresolvedItems']++
                            }
                        }
                    }

                    $tsCompare = [System.Diagnostics.Stopwatch]::GetTimestamp()
                }

                <#
                 STAGE 3: COMPARISON. Window opened at the end of the sample
                 branch above, after the census, so the census is charged to
                 neither stage.

                 Accumulated inside each branch rather than once after the
                 if/else, because the container branch RECURSES into
                 Get-FolderContentHC before it ends. A single accumulator after
                 the branch would fold the entire subtree walk into this item's
                 comparison time and report nonsense.
                #>
                if ($isContainer) {
                    $isIncorrect = if ($CheckInheritedOnly) {
                        $accessDenied -or (-not $acl) -or (-not (Test-AclInheritedOnlyHC -Acl $acl))
                    }
                    else {
                        $accessDenied -or (-not (Test-AclEqualHC -ReferenceSet $folderRulesSet -DifferenceAce $diffAce))
                    }

                    if ($isSample) {
                        $compareTicks = [System.Diagnostics.Stopwatch]::GetTimestamp() - $tsCompare
                        if ($isWarmupSample) {
                            $t['AclCompareWarmupSamples']++
                            $t['AclCompareWarmupTicks'] += $compareTicks
                        }
                        else {
                            $t['AclCompareStrideSamples']++
                            $t['AclCompareStrideTicks'] += $compareTicks
                        }
                    }

                    if ($isIncorrect) {
                        & $incorrectAclInheritedOnly
                    }

                    if ((-not $accessDenied) -or ($Action -eq 'Fix')) {
                        # Pass only the already-enumerated DirectoryInfo; the
                        # function derives the path from it and reuses its cached
                        # .Exists/.Attributes instead of re-stating the path.
                        Get-FolderContentHC -DirectoryInfo $child
                    }
                }
                else {
                    $isIncorrect = if ($CheckInheritedOnly) {
                        $accessDenied -or (-not $acl) -or (-not (Test-AclInheritedOnlyHC -Acl $acl))
                    }
                    else {
                        $accessDenied -or (-not (Test-AclEqualHC -ReferenceSet $fileRulesSet -DifferenceAce $diffAce))
                    }

                    if ($isSample) {
                        $compareTicks = [System.Diagnostics.Stopwatch]::GetTimestamp() - $tsCompare
                        if ($isWarmupSample) {
                            $t['AclCompareWarmupSamples']++
                            $t['AclCompareWarmupTicks'] += $compareTicks
                        }
                        else {
                            $t['AclCompareStrideSamples']++
                            $t['AclCompareStrideTicks'] += $compareTicks
                        }
                    }

                    if ($isIncorrect) {
                        & $incorrectAclInheritedOnly
                    }
                }
            }
        }
        #endregion

        #region ScriptBlock IncorrectAclInheritedOnly
        $incorrectAclInheritedOnly = {
            Write-Warning "Incorrect ACL '$($child.FullName)'"

            # Hoisted for the same reason as in Get-FolderContentHC, and guarded
            # for the same reason: the Pester suite replaces this scriptblock
            # with an empty one in some contexts and dot-sources it standalone in
            # others, so $telemetry cannot be assumed to exist here either.
            $t = $telemetry
            if ($null -eq $t) { $t = @{} }
            $t['IncorrectItems']++

            if ($DetailedLog) {
                # One array element per ACE keeps the detail JSON readable
                # instead of a single string with embedded '\n' escapes. Sort the
                # ACE lines so 'OldAcl' has a stable order (aligns with the
                # non-inherited 'OldAcl'/'NewAcl' warning).
                $aclText = @(if ($accessDenied) { 'Access Denied' } else { $acl.AccessToString -split '\r?\n' | Where-Object { $_ } | Sort-Object })

                if ($AdNames -and $AdNames.Count -gt 0) {
                    # Key name matches the non-inherited warning: 'OldAcl' is the
                    # current ACL found on disk. Inherited-only items have no
                    # target ACL (goal is pure inheritance) so there is no 'NewAcl'.
                    # Use [ordered] so the detail JSON always emits the keys in
                    # the same order (OldAcl, NewAcl, MatrixFileAcl).
                    $entry = [ordered]@{
                        'OldAcl'        = $aclText
                        'MatrixFileAcl' = ConvertTo-MatrixAdObjectHC -Names $AdNames -Permissions $AdPermissions
                    }
                    $incorrectInheritedAcl[$child.FullName] = $entry
                }
                else {
                    $incorrectInheritedAcl[$child.FullName] = $aclText
                }
            }
            else {
                $incorrectInheritedAcl.Add($child.FullName)
            }

            if ($Action -eq 'Fix') {
                Write-Verbose "Set ACL to inherited only '$($child.FullName)'"

                # IMPORTANT: build a FRESH inherited-only ACL object for every
                # item. A DirectorySecurity/FileSecurity remembers which sections
                # were changed (owner/DACL) and clears those "modified" flags
                # after the first successful Persist (SetAccessControl). Reusing
                # a single template object across the walk therefore only wrote
                # the very first folder and the very first file; every later
                # item computed persistRules = None and was silently left
                # untouched (still reported as fixed). Creating a new object per
                # item guarantees the DACL protection + owner are re-marked as
                # modified so each item is genuinely reset to inherited-only.
                # Telemetry: writes are counted and timed separately from
                # reads. Writes are rare relative to reads on a converged
                # tree, so this timer is cheap by construction — and if it
                # ever stops being cheap, that IS the finding: a tree that
                # rewrites the same items every night is not converging.
                $tsWrite = [System.Diagnostics.Stopwatch]::GetTimestamp()

                if ($isContainer) {
                    $dirInfo = [System.IO.DirectoryInfo]::new($child.FullName)

                    $inheritedDirAcl = New-Object System.Security.AccessControl.DirectorySecurity
                    $inheritedDirAcl.SetOwner($builtinAdmin)
                    $inheritedDirAcl.SetAccessRuleProtection($false, $false)

                    if ($accessDenied) {
                        [TokenManipulator]::SetOwner($child.FullName, 'BUILTIN\Administrators')
                    }

                    try {
                        [System.IO.FileSystemAclExtensions]::SetAccessControl($dirInfo, $inheritedDirAcl)
                    }
                    catch [System.UnauthorizedAccessException] {
                        $t['AclWriteDenied']++
                        [TokenManipulator]::SetOwner($child.FullName, 'BUILTIN\Administrators')
                        [System.IO.FileSystemAclExtensions]::SetAccessControl($dirInfo, $inheritedDirAcl)
                    }
                }
                else {
                    $fileInfo = [System.IO.FileInfo]::new($child.FullName)

                    $inheritedFileAcl = New-Object System.Security.AccessControl.FileSecurity
                    $inheritedFileAcl.SetOwner($builtinAdmin)
                    $inheritedFileAcl.SetAccessRuleProtection($false, $false)

                    if ($accessDenied) {
                        [TokenManipulator]::SetOwner($child.FullName, 'BUILTIN\Administrators')
                    }

                    try {
                        [System.IO.FileSystemAclExtensions]::SetAccessControl($fileInfo, $inheritedFileAcl)
                    }
                    catch [System.UnauthorizedAccessException] {
                        $t['AclWriteDenied']++
                        [TokenManipulator]::SetOwner($child.FullName, 'BUILTIN\Administrators')
                        [System.IO.FileSystemAclExtensions]::SetAccessControl($fileInfo, $inheritedFileAcl)
                    }
                }

                # Unsampled: writes are exceptional on a converged tree, so
                # the timer runs on every one. If that ever becomes expensive,
                # the expense IS the finding — a tree that rewrites the same
                # items every night is not converging.
                $t['AclWrites']++
                $t['AclWriteTicks'] += (
                    [System.Diagnostics.Stopwatch]::GetTimestamp() - $tsWrite
                )
            }
        }
        #endregion

        try {
            #region Logging Setup
            if ($CollectTestedPaths) {
                $testedInheritedFilesAndFolders = @{ }
            }

            if ($DetailedLog) {
                $incorrectInheritedAcl = @{ }
                $unreadableAcl = @{ }
            }
            else {
                $incorrectInheritedAcl = [System.Collections.Generic.List[String]]::New()
                $unreadableAcl = [System.Collections.Generic.List[String]]::New()
            }
            #endregion

            #region Get super powers
            try {
                Write-Verbose 'Get super powers'

                if (-not ('TokenManipulator' -as [type])) {
                    try {
                        Add-Type $tokenPrivileges -ErrorAction Stop
                    }
                    catch {
                        if ($_.Exception.Message -notmatch 'already exists') {
                            throw $_
                        }
                    }
                }

                [void][TokenManipulator]::AddPrivilege('SeRestorePrivilege')
                [void][TokenManipulator]::AddPrivilege('SeBackupPrivilege')
                [void][TokenManipulator]::AddPrivilege('SeTakeOwnershipPrivilege')
            }
            catch { throw "Failed getting super powers: $_" }
            #endregion

            #region Create inherited folder and file acl
            # Only the owner principal is prepared here. The actual inherited-only
            # ACL objects are (re)created per item inside $incorrectAclInheritedOnly
            # because a DirectorySecurity/FileSecurity clears its modified-section
            # flags after its first Persist, so a shared template would only ever
            # rewrite the first folder and the first file (see comment there).
            Write-Verbose 'Inherited permissions'
            $builtinAdmin = [System.Security.Principal.NTAccount]'BUILTIN\Administrators'
            #endregion

            #region Check or fix the seed folder itself when it is inherit-only
            if ($CheckSeedPath) {
                $child = [System.IO.DirectoryInfo]::new($Path)
                $isContainer = $true
                $accessDenied = $false
                $acl = $null
                $unreadable = $false

                $aclRead = Get-DirectoryAclSafeHC -DirectoryInfo $child
                $acl = $aclRead.Acl
                $accessDenied = $aclRead.AccessDenied

                if ($aclRead.Removed) {
                    Write-Verbose "Seed folder '$($child.FullName)' removed"
                    $unreadable = $true
                }
                elseif ($aclRead.UnreadableReason) {
                    Write-Warning "Failed retrieving the ACL of '$($child.FullName)': $($aclRead.UnreadableReason)"

                    # ACL unreadable (not access-denied): do not check or
                    # reset it. Report under 'ACL could not be read'.
                    if ($DetailedLog) {
                        $unreadableAcl[$child.FullName] = New-UnreadableAclEntryHC -Reason $aclRead.UnreadableReason -AdNames $AdNames -AdPermissions $AdPermissions
                    }
                    else {
                        $unreadableAcl.Add($child.FullName)
                    }
                    $unreadable = $true
                }

                if (-not $unreadable) {
                    if ($CollectTestedPaths) {
                        $testedInheritedFilesAndFolders[$child.FullName] = $true
                    }

                    if ($accessDenied -or (-not $acl) -or (-not (Test-AclInheritedOnlyHC -Acl $acl))) {
                        & $incorrectAclInheritedOnly
                    }
                }
            }
            #endregion

            #region Check or fix folder and file permissions
            try { Get-FolderContentHC -DirectoryInfo ([System.IO.DirectoryInfo]::new($Path)) } catch { throw "Failed checking or setting the inheritance in folder '$Path': $_" }
            #endregion
        }
        catch { throw "Failed setting permissions for '$Path': $_" }
        finally {
            # 'Telemetry' rides along on the object that already crosses the
            # runspace boundary, so it costs nothing extra to transport: a
            # fixed ~15 integers per job regardless of tree size. Emitted from
            # 'finally' so a job that throws still reports the volume it
            # managed to get through before failing.
            #
            # Closing the job window here for the same reason: a job that threw
            # still consumed wall clock, and a straggler that fails late is
            # exactly the case worth seeing.
            $telemetry['JobEndTicks'] = [System.Diagnostics.Stopwatch]::GetTimestamp()

            $result = [PSCustomObject]@{
                IncorrectInheritedAcl = $incorrectInheritedAcl
                UnreadableAcl         = $unreadableAcl
                Telemetry             = $telemetry
            }
            if ($CollectTestedPaths) {
                $result | Add-Member -NotePropertyName 'TestedInheritedFilesAndFolders' -NotePropertyValue $testedInheritedFilesAndFolders
            }
            $result
        }
    }
    #endregion

    #region TokenManipulator C# Class
    $tokenPrivileges = @'
using System;
using System.Runtime.InteropServices;
using System.Security.Principal;

public class TokenManipulator
{
    [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
    internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall, ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);

    [DllImport("kernel32.dll", ExactSpelling = true)]
    internal static extern IntPtr GetCurrentProcess();

    [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
    internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);

    [DllImport("advapi32.dll", SetLastError = true)]
    internal static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);

    [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    internal static extern uint SetNamedSecurityInfo(string pObjectName, int objectType, uint securityInfo, byte[] psidOwner, byte[] psidGroup, IntPtr pDacl, IntPtr pSacl);

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    internal struct TokPriv1Luid
    {
        public int Count;
        public long Luid;
        public int Attr;
    }

    internal const int SE_PRIVILEGE_DISABLED = 0x00000000;
    internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
    internal const int TOKEN_QUERY = 0x00000008;
    internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;

    internal const uint OWNER_SECURITY_INFORMATION = 0x00000001;
    internal const int SE_FILE_OBJECT = 1;
    internal const int ERROR_NOT_ALL_ASSIGNED = 1300;

    public static bool AddPrivilege(string privilege)
    {
        TokPriv1Luid tp;
        IntPtr hproc = GetCurrentProcess();
        IntPtr htok = IntPtr.Zero;
        if (!OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok))
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        tp.Count = 1;
        tp.Luid = 0;
        tp.Attr = SE_PRIVILEGE_ENABLED;
        if (!LookupPrivilegeValue(null, privilege, ref tp.Luid))
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        bool retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
        int lastError = Marshal.GetLastWin32Error();
        // AdjustTokenPrivileges can return true even when the privilege was not
        // held by the token; ERROR_NOT_ALL_ASSIGNED signals that partial failure.
        if (!retVal || lastError == ERROR_NOT_ALL_ASSIGNED)
            throw new System.ComponentModel.Win32Exception(lastError == 0 ? ERROR_NOT_ALL_ASSIGNED : lastError);
        return retVal;
    }

    public static bool RemovePrivilege(string privilege)
    {
        TokPriv1Luid tp;
        IntPtr hproc = GetCurrentProcess();
        IntPtr htok = IntPtr.Zero;
        if (!OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok))
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        tp.Count = 1;
        tp.Luid = 0;
        tp.Attr = SE_PRIVILEGE_DISABLED;
        if (!LookupPrivilegeValue(null, privilege, ref tp.Luid))
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        bool retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
        int lastError = Marshal.GetLastWin32Error();
        if (!retVal || lastError == ERROR_NOT_ALL_ASSIGNED)
            throw new System.ComponentModel.Win32Exception(lastError == 0 ? ERROR_NOT_ALL_ASSIGNED : lastError);
        return retVal;
    }

    public static void SetOwner(string path, string accountName)
    {
        NTAccount account = new NTAccount(accountName);
        SecurityIdentifier sid = (SecurityIdentifier)account.Translate(typeof(SecurityIdentifier));
        byte[] sidBytes = new byte[sid.BinaryLength];
        sid.GetBinaryForm(sidBytes, 0);

        uint result = SetNamedSecurityInfo(path, SE_FILE_OBJECT, OWNER_SECURITY_INFORMATION, sidBytes, null, IntPtr.Zero, IntPtr.Zero);
        if (result != 0)
        {
            throw new System.ComponentModel.Win32Exception((int)result);
        }
    }
}
'@
    #endregion
}

process {
    try {
        $ErrorActionPreference = 'Stop'

        #region Pre-process the Matrix properties
        $missingFolders = [System.Collections.Generic.List[String]]::New()

        if ($Matrix) {
            foreach ($M in $Matrix) {
                if (-not $M.PSObject.Properties.Match('Parent').Count) {
                    $M | Add-Member -NotePropertyName 'Parent' -NotePropertyValue $false
                }
                if (-not $M.PSObject.Properties.Match('Ignore').Count) {
                    $M | Add-Member -NotePropertyName 'Ignore' -NotePropertyValue $false
                }

                if ($M.PSObject.Properties.Match('ACL').Count) {
                    $M.ACL = ConvertTo-HashtableHC -InputObject $M.ACL
                }

                # PSRemoting deserializes nested hashtables to PSCustomObject.
                # Rebuild 'AdNames' (added by the BEGIN stage via Add-Member)
                # just like 'ACL' above, so it binds to the strictly typed
                # [hashtable]$AdNames parameter of the inherited permissions
                # scriptblock. Guard on the property existing so we never add it.
                if ($M.PSObject.Properties.Match('AdNames').Count) {
                    $M.AdNames = ConvertTo-HashtableHC -InputObject $M.AdNames
                }
            }
        }
        #endregion

        #region Logging Setup
        if ($CollectTestedPaths) {
            $testedInheritedFilesAndFolders = @{ }
        }

        if ($DetailedLog) {
            $incorrectAclNonInheritedFolders = @{ }
            $incorrectInheritedAcl = @{ }
            $unreadableAcl = @{ }
        }
        else {
            $incorrectAclNonInheritedFolders = [System.Collections.Generic.List[String]]::New()
            $incorrectInheritedAcl = [System.Collections.Generic.List[String]]::New()
            $unreadableAcl = [System.Collections.Generic.List[String]]::New()
        }
        #endregion

        #region Telemetry accumulator (Main Thread)
        <#
         The run-wide totals for this one Settings row. Two sources feed it:

         - the matrix-folder loop below, which runs on this thread and touches
           only the folders named in the 'Permissions' worksheet (hundreds of
           items — cold path, cost irrelevant)
         - the parallel walker, which touches every file and folder underneath
           them (millions of items — hot path, see the counter notes in the
           scriptblock)

         'MatrixFolders*' therefore describes the explicit ACLs, and
         'Folders/FilesWalked' describes the inherited ones. Keeping them
         apart matters: a matrix that grows by ten rows and a share that grows
         by 200k files are different problems with the same symptom.

         Stopwatch.Frequency is captured once and applied at the end, so no
         tick-to-millisecond division happens per item.
        #>
        $telemetryStart = [System.Diagnostics.Stopwatch]::GetTimestamp()

        <#
         Per-path breakdown, so a regression can be localised to the folder that
         caused it instead of only to the Settings row.

         Two sources, keyed on the same folder path:
         - this thread's explicit-ACL loop (one entry per matrix folder)
         - each parallel walker job (one job per matrix folder)

         Because the walker jobs partition the tree, the per-path rows sum back
         to the Settings-row totals. That is worth preserving: it means the
         breakdown can be trusted as a decomposition rather than a sample.
        #>
        $pathTelemetry = @{}

        $newPathRow = {
            param($RowPath)

            if (-not $pathTelemetry.ContainsKey($RowPath)) {
                $pathTelemetry[$RowPath] = @{
                    Path                   = $RowPath
                    # Job window for this subtree. MIN of starts and MAX of
                    # ends, so a path served by more than one job still reports
                    # the span that path occupied rather than a meaningless sum
                    # of absolute timestamps.
                    JobStartTicks          = 0L
                    JobEndTicks            = 0L
                    JobCount               = 0L
                    IdentitySet            = [System.Collections.Generic.HashSet[String]]::new()
                    IdentityObservations   = 0L
                    IdentityTruncated      = 0L
                    AceUnresolvedSids      = 0L
                    AceUnresolvedItems     = 0L
                    MatrixFolderReads      = 0L
                    MatrixFolderReadTicks  = 0L
                    MatrixFolderWrites     = 0L
                    MatrixFolderWriteTicks = 0L
                    MatrixFoldersIncorrect = 0L
                    MatrixFoldersCreated   = 0L
                    FoldersWalked          = 0L
                    FilesWalked            = 0L
                    AclReadWarmupSamples   = 0L
                    AclReadWarmupTicks     = 0L
                    AclReadStrideSamples   = 0L
                    AclReadStrideTicks     = 0L
                    AclProjectWarmupSamples = 0L
                    AclProjectWarmupTicks   = 0L
                    AclProjectStrideSamples = 0L
                    AclProjectStrideTicks   = 0L
                    AclCompareWarmupSamples = 0L
                    AclCompareWarmupTicks   = 0L
                    AclCompareStrideSamples = 0L
                    AclCompareStrideTicks   = 0L
                    AclReadDenied          = 0L
                    AclReadFailed          = 0L
                    AclWrites              = 0L
                    AclWriteTicks          = 0L
                    AclWriteDenied         = 0L
                    IncorrectItems         = 0L
                    AceCountTotal          = 0L
                    AceCountItems          = 0L
                    AceCountMax            = 0L
                    EnumeratedDirs         = 0L
                    EnumerateTicks         = 0L
                    Walked                 = $false
                }
            }

            return $pathTelemetry[$RowPath]
        }

        $telemetry = @{
            MatrixFolders          = 0L
            MatrixFolderReads      = 0L
            MatrixFolderReadTicks  = 0L
            MatrixFolderWrites     = 0L
            MatrixFolderWriteTicks = 0L
            MatrixFoldersIncorrect = 0L
            MatrixFoldersCreated   = 0L
            FoldersWalked          = 0L
            FilesWalked            = 0L
            IdentitySet            = [System.Collections.Generic.HashSet[String]]::new()
            IdentityObservations   = 0L
            IdentityTruncated      = 0L
            AceUnresolvedSids      = 0L
            AceUnresolvedItems     = 0L
            AclReadWarmupSamples   = 0L
            AclReadWarmupTicks     = 0L
            AclReadStrideSamples   = 0L
            AclReadStrideTicks     = 0L
            AclProjectWarmupSamples = 0L
            AclProjectWarmupTicks   = 0L
            AclProjectStrideSamples = 0L
            AclProjectStrideTicks   = 0L
            AclCompareWarmupSamples = 0L
            AclCompareWarmupTicks   = 0L
            AclCompareStrideSamples = 0L
            AclCompareStrideTicks   = 0L
            AclReadDenied          = 0L
            AclReadFailed          = 0L
            AclWrites              = 0L
            AclWriteTicks          = 0L
            AclWriteDenied         = 0L
            IncorrectItems         = 0L
            AceCountTotal          = 0L
            AceCountItems          = 0L
            AceCountMax            = 0L
            EnumeratedDirs         = 0L
            EnumerateTicks         = 0L
            SampleEvery            = 0L
            SampleWarmup           = 0L
        }
        #endregion

        #region Get super powers
        try {
            Write-Verbose 'Get super powers'

            if (-not ('TokenManipulator' -as [type])) {
                try {
                    Add-Type $tokenPrivileges -ErrorAction Stop
                }
                catch {
                    if ($_.Exception.Message -notmatch 'already exists') {
                        throw $_
                    }
                }
            }

            [void][TokenManipulator]::AddPrivilege('SeRestorePrivilege')
            [void][TokenManipulator]::AddPrivilege('SeBackupPrivilege')
            [void][TokenManipulator]::AddPrivilege('SeTakeOwnershipPrivilege')
        }
        catch { throw "Failed getting super powers: $_" }
        #endregion

        #region Import library for .NET calls
        try { Import-Module -Name 'Microsoft.PowerShell.Security' -Force } catch { throw "Failed loading .NET library: $_" }
        #endregion

        #region Create the parent folder when action is New
        try {
            # A file at the parent path can never be a valid parent folder.
            # Report it clearly for every action instead of the misleading
            # 'exists already' (New) / 'missing' (Check/Fix) messages, which
            # would otherwise bounce the user between the two actions.
            if (Test-Path -LiteralPath $Path -PathType Leaf) {
                return [PSCustomObject]@{
                    DateTime    = Get-Date
                    Type        = 'FatalError'
                    Name        = 'Parent folder path occupied by a file'
                    Description = "The path defined as 'Path' in the worksheet 'Settings' already exists as a file on the remote machine. Please remove or rename the file, or correct the path."
                    Value       = $Path
                }
            }

            if ($Action -eq 'New') {
                # Only a pre-existing path is a 'exists already' case; any other
                # New-Item failure (access denied, invalid path, missing drive/share,
                # a path segment that is a file) must surface its real cause.
                if (Test-Path -LiteralPath $Path) {
                    return [PSCustomObject]@{
                        DateTime    = Get-Date
                        Type        = 'FatalError'
                        Name        = 'Parent folder exists already'
                        Description = "The folder defined as 'Path' in the worksheet 'Settings' cannot be present on the remote machine when 'Action=New' is used. Please use 'Action' with value 'Check' or 'Fix' instead."
                        Value       = $Path
                    }
                }

                try { $missingFolders.Add((New-Item -Path $Path -ItemType Directory -EA Stop).FullName) }
                catch {
                    $errorMessage = $_.Exception.Message
                    $Error.RemoveAt(0)
                    return [PSCustomObject]@{
                        DateTime    = Get-Date
                        Type        = 'FatalError'
                        Name        = 'Parent folder creation failed'
                        Description = "The folder defined as 'Path' in the worksheet 'Settings' could not be created: $errorMessage"
                        Value       = $Path
                    }
                }
            }
            elseif (-not (Test-Path -LiteralPath $Path -PathType Container)) {
                return [PSCustomObject]@{
                    DateTime    = Get-Date
                    Type        = 'FatalError'
                    Name        = 'Parent folder missing'
                    Description = "The folder defined as 'Path' in the worksheet 'Settings' needs to be available on the remote machine. In case the folder structure needs to be created, please use 'Action=New' instead."
                    Value       = $Path
                }
            }

            Write-Verbose "Parent folder '$Path'"
        }
        catch { throw "Failed checking the existence of the parent folder: $_" }
        #endregion

        #region Add the FullName for each path
        foreach ($M in $Matrix) {
            $tmpPath = if ($M.Parent) { $Path } else { Join-Path -Path $Path -ChildPath $M.Path }
            $M.Path = $tmpPath
        }
        #endregion

        #region Remove ignored folders from the matrix
        $ignoredFolders, $Matrix = $Matrix.Where( { $_.Ignore }, 'Split')
        $ignoredFolderPaths = @{}

        if ($ignoredFolders) {
            $IgnoredFolders.Path.ForEach({
                    Write-Verbose "Ignored folder '$_'"
                    $ignoredFolderPaths[$_] = $true
                })

            [PSCustomObject]@{
                DateTime    = Get-Date
                Type        = 'Information'
                Name        = 'Ignored folder'
                Description = "All rows in the worksheet 'Permissions' that have the character 'i' defined are ignored. These folders are not checked for incorrect permissions."
                Value       = $IgnoredFolders.Path
            }
        }
        #endregion

        #region Create file and folder ACL for each path in the matrix
        try {
            Write-Verbose "Create ACE 'BUILTIN\Administrators' : 'FullControl'"
            $builtinAdmin = [System.Security.Principal.NTAccount]'BUILTIN\Administrators'

            $adminFullControlAce = @{
                Folder = New-Object System.Security.AccessControl.FileSystemAccessRule($builtinAdmin, [System.Security.AccessControl.FileSystemRights]::FullControl, [System.Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit', [System.Security.AccessControl.PropagationFlags]::None, [System.Security.AccessControl.AccessControlType]::Allow)
                File   = New-Object System.Security.AccessControl.FileSystemAccessRule($builtinAdmin, [System.Security.AccessControl.FileSystemRights]::FullControl, [System.Security.AccessControl.AccessControlType]::Allow)
            }

            foreach ($M in $Matrix) {
                $M | Add-Member -NotePropertyMembers @{ FolderAcl = $null; InheritedFileAcl = $null; InheritedFolderAcl = $null }
            }

            $Matrix.Where( { $_.ACL.Count -eq 0 }).ForEach( { $_.ACL = $null })

            $aceCache = @{ }

            foreach ($M in $Matrix.Where( { $_.ACL })) {
                Write-Verbose "Create ACL for path '$($M.Path)'"

                $acl = @{
                    Folder          = New-Object System.Security.AccessControl.DirectorySecurity
                    InheritedFolder = New-Object System.Security.AccessControl.DirectorySecurity
                    InheritedFile   = New-Object System.Security.AccessControl.FileSecurity
                }

                $acl.Folder.SetAccessRuleProtection($true, $false)
                $acl.Folder.SetOwner($builtinAdmin)

                $acl.InheritedFolder.SetAccessRuleProtection($false, $false)
                $acl.InheritedFolder.SetOwner($builtinAdmin)

                $acl.InheritedFile.SetAccessRuleProtection($false, $false)
                $acl.InheritedFile.SetOwner($builtinAdmin)

                $M.ACL.GetEnumerator().Foreach({
                        try {
                            $ID = "$($_.Key)@$($_.Value)"

                            if (-not $aceCache.ContainsKey($ID)) {
                                $param = @{ Access = $_.Value; Name = $_.Key }
                                $aceCache[$ID] = @{
                                    Folder          = @( New-AceHC @param -Type 'Folder' )
                                    InheritedFolder = @( New-AceHC @param -Type 'InheritedFolder' )
                                    InheritedFile   = @( New-AceHC @param -Type 'InheritedFile' )
                                }
                            }

                            $aceCache[$ID]['Folder'].ForEach({ $acl.Folder.AddAccessRule($_) })
                            $aceCache[$ID]['InheritedFolder'].ForEach({ $acl.InheritedFolder.AddAccessRule($_) })
                            $aceCache[$ID]['InheritedFile'].ForEach({ $acl.InheritedFile.AddAccessRule($_) })
                        }
                        catch { throw "AD object '$($ID.split('@')[0])' with permission character '$($ID.split('@')[1])' probably doesn't exist in AD: $_" }
                    })

                $acl.Folder.AddAccessRule($adminFullControlAce.Folder)
                $acl.InheritedFolder.AddAccessRule($adminFullControlAce.Folder)
                $acl.InheritedFile.AddAccessRule($adminFullControlAce.File)

                $M.FolderAcl = $acl.Folder
                $M.inheritedFolderAcl = $acl.InheritedFolder
                $M.inheritedFileAcl = $acl.InheritedFile
            }
        }
        catch { throw "Failed creating the AccessControlList: $_" }
        #endregion

        #region Create Missing Folders (Check/Fix Matrix)
        try {
            $pathsToCreate = [System.Collections.Generic.List[String]]::New()
            foreach ($M in $Matrix) {
                if (($M.Parent -eq $false) -and (-not (Test-Path -LiteralPath $M.Path -PathType Container))) {
                    $pathsToCreate.Add($M.Path)
                }
            }

            foreach ($nonExistingPath in $pathsToCreate) {
                # A file occupying the folder's path blocks New-Item -Directory
                # (it throws instead of overwriting), so report it and move on.
                if (Test-Path -LiteralPath $nonExistingPath -PathType Leaf) {
                    Write-Verbose "Folder path occupied by a file '$nonExistingPath'"
                    [PSCustomObject]@{
                        DateTime    = Get-Date
                        Type        = 'FatalError'
                        Name        = 'Folder path occupied by a file'
                        Description = "A folder defined in the worksheet 'Permissions' cannot be created because a file with the same name already exists on the remote machine. Please remove or rename the file, or correct the matrix."
                        Value       = $nonExistingPath
                    }
                    continue
                }

                if ($Action -eq 'Check') {
                    Write-Verbose "Missing folder '$nonExistingPath'"
                    $missingFolders.Add($nonExistingPath)
                }
                else {
                    Write-Verbose "Create missing folder '$nonExistingPath'"
                    $missingFolders.Add((New-Item -Path $nonExistingPath -ItemType Directory -Force -EA Stop).FullName)
                    $telemetry['MatrixFoldersCreated']++
                }
            }

            if ($Action -eq 'Check' -and $missingFolders.Count -gt 0) {
                $Matrix = $Matrix.Where({ $_.Path -notin $missingFolders })
            }

            if ($missingFolders.Count -ne 0) {
                $Obj = [PSCustomObject]@{
                    DateTime    = Get-Date
                    Type        = 'Warning'
                    Name        = $null
                    Description = $null
                    Value       = $missingFolders.ToArray()
                }

                switch ($Action) {
                    'New' { $Obj.Name = 'Child folder created'; $Obj.Description = "All folders defined in the worksheet 'Permissions' have been created with the correct permissions underneath the parent folder defined in the worksheet 'Settings'."; break }
                    'Fix' { $Obj.Name = 'Child folder created'; $Obj.Description = 'The missing folders underneath the parent folder have been created.'; break }
                    'Check' { $Obj.Name = 'Child folder missing'; $Obj.Description = "Not all folders defined in the worksheet 'Permissions' were found underneath the parent folder."; break }
                    default { throw "Action '$_' is not supported." }
                }

                $Obj
            }
            else { Write-Verbose 'All folders present, no missing folders' }
        }
        catch { throw "Failed checking/creating the missing child folders: $_" }
        #endregion

        #region Non-Inherited folder permissions check and apply
        if ($CollectTestedPaths) {
            $testedNonInheritedFolders = @{}
        }
        Write-Verbose 'Folders with ACL in the matrix that are not ignored'

        [array]$foldersWithAcl = $Matrix.Where({ ($_.FolderAcl) -and (-not $_.ignore) }) | Sort-Object -Property 'Path'
        [array]$foldersWithInheritedOnlyAcl = $Matrix.Where({ (-not $_.FolderAcl) -and (-not $_.ignore) }) | Sort-Object -Property 'Path'

        foreach ($folder in $foldersWithInheritedOnlyAcl) {
            $ignoredFolderPaths[$folder.Path] = $true
        }

        foreach ($folder in $foldersWithAcl) {
            try {
                $ignoredFolderPaths[$folder.Path] = $true
                Write-Verbose "Matrix ACL folder '$($folder.Path)'"

                $dirInfo = [System.IO.DirectoryInfo]::new($folder.Path)
                if ($CollectTestedPaths) {
                    $testedNonInheritedFolders[$folder.Path] = $folder
                }

                $accessDenied = $false
                $acl = $null

                $telemetry['MatrixFolders']++
                $pathRow = & $newPathRow $folder.Path

                $tsRead = [System.Diagnostics.Stopwatch]::GetTimestamp()
                $aclRead = Get-DirectoryAclSafeHC -DirectoryInfo $dirInfo
                $readTicks = [System.Diagnostics.Stopwatch]::GetTimestamp() - $tsRead

                $telemetry['MatrixFolderReads']++
                $telemetry['MatrixFolderReadTicks'] += $readTicks
                $pathRow['MatrixFolderReads']++
                $pathRow['MatrixFolderReadTicks'] += $readTicks

                $acl = $aclRead.Acl
                $accessDenied = $aclRead.AccessDenied

                if ($aclRead.Removed) {
                    Write-Verbose "Matrix folder '$($folder.Path)' removed"
                    continue
                }
                if ($aclRead.UnreadableReason) {
                    Write-Warning "Failed retrieving the ACL of '$($folder.Path)': $($aclRead.UnreadableReason)"

                    # The ACL could not be read (not access-denied), so this
                    # matrix folder was neither checked nor corrected. Report it
                    # under the dedicated 'ACL could not be read' warning instead
                    # of aborting the run.
                    if ($DetailedLog) {
                        $folderAdNames = ConvertTo-HashtableHC -InputObject $folder.AdNames
                        $unreadableAcl[$folder.Path] = New-UnreadableAclEntryHC -Reason $aclRead.UnreadableReason -AdNames $folderAdNames -AdPermissions $folder.ACL
                    }
                    else {
                        $unreadableAcl.Add($folder.Path)
                    }
                    continue
                }

                $diffAce = if (-not $accessDenied -and $acl) { @($acl.Access) } else { @() }

                # Unsampled here: this loop covers only the folders named in
                # the 'Permissions' worksheet — hundreds of items, not
                # millions — so the cold-path cost is irrelevant and the census
                # is exact.
                if ($diffAce.Count) {
                    $telemetry['AceCountTotal'] += $diffAce.Count
                    $telemetry['AceCountItems']++
                    if ($diffAce.Count -gt $telemetry['AceCountMax']) {
                        $telemetry['AceCountMax'] = $diffAce.Count
                    }
                }

                if ($accessDenied -or (-not $acl) -or (-not $acl.AreAccessRulesProtected) -or (-not (Test-AclEqualHC -ReferenceAce ($folder.FolderAcl).Access -DifferenceAce $diffAce))) {
                    Write-Warning "Incorrect folder ACL '$($folder.Path)'"

                    $telemetry['MatrixFoldersIncorrect']++
                    $pathRow['MatrixFoldersIncorrect']++

                    #region Log Incorrect ACL
                    if ($Action -ne 'New') {
                        if ($DetailedLog) {
                            # Split the multi-line AccessToString into one array
                            # element per ACE so the detail JSON stays readable
                            # instead of a single string with embedded '\n'.
                            # Sort the ACE lines so 'OldAcl' and 'NewAcl' have a
                            # stable, comparable order (mirrors 'MatrixFileAcl').
                            # Use [ordered] so the detail JSON always emits the
                            # keys in the same order (OldAcl, NewAcl, MatrixFileAcl).
                            $entry = [ordered]@{
                                'OldAcl' = @(if ($accessDenied) { 'Access Denied' } else { $acl.AccessToString -split '\r?\n' | Where-Object { $_ } | Sort-Object })
                                'NewAcl' = @(($folder.FolderAcl).AccessToString -split '\r?\n' | Where-Object { $_ } | Sort-Object)
                            }

                            # Surface the matrix-author labels so users can map
                            # ACL entries back to their Excel column headers.
                            # Each entry is the display form (DOMAIN\name when the
                            # SID translates, raw SID when it doesn't) followed by
                            # the requested permission, which matches what
                            # AccessToString puts in OldAcl/NewAcl.
                            # Defensive: rebuild AdNames if it crossed a
                            # serialization boundary and arrived as a
                            # Deserialized.PSCustomObject (no .Keys/.Count).
                            $folderAdNames = ConvertTo-HashtableHC -InputObject $folder.AdNames

                            if ($folderAdNames -and $folderAdNames.Count -gt 0) {
                                $entry['MatrixFileAcl'] = ConvertTo-MatrixAdObjectHC -Names $folderAdNames -Permissions $folder.ACL
                            }

                            $incorrectAclNonInheritedFolders[$folder.Path] = $entry
                        }
                        else {
                            $incorrectAclNonInheritedFolders.Add($folder.Path)
                        }
                    }
                    #endregion

                    #region Set corrected ACL
                    if ($Action -ne 'Check') {
                        Write-Verbose 'Set correct ACL'

                        if ($accessDenied) { [TokenManipulator]::SetOwner($folder.Path, 'BUILTIN\Administrators') }

                        $newAcl = [System.Security.AccessControl.DirectorySecurity]::new()
                        $newAcl.SetOwner($builtinAdmin)
                        $newAcl.SetAccessRuleProtection($true, $false)
                        foreach ($rule in $folder.FolderAcl.Access) { $newAcl.AddAccessRule($rule) }

                        $tsWrite = [System.Diagnostics.Stopwatch]::GetTimestamp()

                        try {
                            [System.IO.FileSystemAclExtensions]::SetAccessControl($dirInfo, $newAcl)
                        }
                        catch [System.UnauthorizedAccessException] {
                            [TokenManipulator]::SetOwner($folder.Path, 'BUILTIN\Administrators')
                            [System.IO.FileSystemAclExtensions]::SetAccessControl($dirInfo, $newAcl)
                        }

                        $writeTicks = [System.Diagnostics.Stopwatch]::GetTimestamp() - $tsWrite

                        $telemetry['MatrixFolderWrites']++
                        $telemetry['MatrixFolderWriteTicks'] += $writeTicks
                        $pathRow['MatrixFolderWrites']++
                        $pathRow['MatrixFolderWriteTicks'] += $writeTicks

                        Write-Verbose 'ACL corrected'
                    }
                    #endregion
                }
            }
            catch { throw "Failed checking/setting the permissions on non inherited folder '$($folder.Path)': $_" }
        }

        if ($incorrectAclNonInheritedFolders.Count -ne 0) {
            [PSCustomObject]@{
                DateTime    = Get-Date
                Type        = 'Warning'
                Name        = 'Non inherited folder incorrect permissions'
                Description = "The folders that have permissions defined in the worksheet 'Permissions' are not matching with the permissions found on the folders of the remote machine."
                Value       = if ($DetailedLog) { $incorrectAclNonInheritedFolders } else { $incorrectAclNonInheritedFolders.ToArray() }
            }
        }
        #endregion

        #region Inherited folder and file permissions check and apply
        try {
            Write-Verbose 'Inherited permissions'
            if ($Action -ne 'New') {

                $ErrorActionPreference = 'Continue'
                $scriptBlockString = $inheritedPermissionsScriptBlock.ToString()

                $extractRules = {
                    param($acl)
                    if (-not $acl) { return @() }
                    $arr = [System.Collections.Generic.List[String]]::New()
                    foreach ($r in $acl.Access) {
                        # OPTIMIZATION: Extract to primitive string before sending into the runspace!
                        $arr.Add("$([int]$r.FileSystemRights)|$([int]$r.AccessControlType)|$($r.IdentityReference.ToString())|$([int]$r.InheritanceFlags)")
                    }
                    return $arr
                }

                # Fields shared by every DTO; merged with the per-folder fields
                # below so the two loops don't repeat the common half.
                $sharedDto = @{
                    Action             = $Action
                    IgnoredFolderPaths = $ignoredFolderPaths
                    TokenPrivileges    = $tokenPrivileges
                    DetailedLog        = $DetailedLog
                    CollectTestedPaths = $CollectTestedPaths
                    ScriptString       = $scriptBlockString
                }

                $safeFolders = @(
                    foreach ($folder in $foldersWithAcl) {
                        [PSCustomObject]($sharedDto + @{
                                Path               = $folder.Path
                                FolderRules        = &$extractRules $folder.InheritedFolderAcl
                                FileRules          = &$extractRules $folder.InheritedFileAcl
                                CheckSeedPath      = $false
                                CheckInheritedOnly = $false
                                AdNames            = $folder.AdNames
                                AdPermissions      = $folder.ACL
                            })
                    }

                    foreach ($folder in $foldersWithInheritedOnlyAcl) {
                        [PSCustomObject]($sharedDto + @{
                                Path               = $folder.Path
                                FolderRules        = @()
                                FileRules          = @()
                                CheckSeedPath      = $true
                                CheckInheritedOnly = $true
                                AdNames            = $null
                                AdPermissions      = $null
                            })
                    }
                )

                $jobResults = $safeFolders | ForEach-Object -Parallel {
                    $folderDto = $_

                    $params = @{
                        Path                = $folderDto.Path
                        Action              = $folderDto.Action
                        FolderAclAccessList = $folderDto.FolderRules
                        FileAclAccessList   = $folderDto.FileRules
                        IgnoredFolderPaths  = $folderDto.IgnoredFolderPaths
                        TokenPrivileges     = $folderDto.TokenPrivileges
                        AdNames             = $folderDto.AdNames
                        AdPermissions       = $folderDto.AdPermissions
                        CheckSeedPath       = $folderDto.CheckSeedPath
                        CheckInheritedOnly  = $folderDto.CheckInheritedOnly
                        DetailedLog         = $folderDto.DetailedLog
                        CollectTestedPaths  = $folderDto.CollectTestedPaths
                    }

                    $rehydratedBlock = [scriptblock]::Create($folderDto.ScriptString)
                    & $rehydratedBlock @params

                } -ThrottleLimit $JobThrottleLimit

                foreach ($jobResult in $jobResults) {
                    #region Merge telemetry from this worker
                    # Counts and ticks are additive across workers. 'AceCountMax'
                    # is the one exception: it is a maximum, not a sum.
                    #
                    # Note that the tick totals are the sum of CONCURRENT work,
                    # so they exceed the job's wall clock by roughly the
                    # throttle limit. That is intentional — it measures cost,
                    # not elapsed time, and cost is what is comparable between
                    # runs when the concurrency setting is unchanged.
                    if ($jobResult.Telemetry) {
                        # Per-path breakdown: fold this job's counters into the
                        # row for the subtree it walked, before summing them into
                        # the Settings-row totals below.
                        $walkedPath = $jobResult.Telemetry['WalkedPath']

                        if ($walkedPath) {
                            $jobRow = & $newPathRow $walkedPath
                            $jobRow['Walked'] = $true

                            if ($jobResult.Telemetry['SeedOnly']) {
                                $jobRow['SeedOnly'] = $true
                            }

                            $jobRow['JobCount']++

                            foreach ($j in $jobResult.Telemetry.GetEnumerator()) {
                                if (-not $jobRow.ContainsKey($j.Key)) { continue }
                                if ($j.Key -eq 'AceCountMax') {
                                    if ($j.Value -gt $jobRow['AceCountMax']) {
                                        $jobRow['AceCountMax'] = $j.Value
                                    }
                                }
                                elseif ($j.Key -eq 'IdentitySet') {
                                    # Union, not sum. Falling through to the
                                    # generic branch would silently DROP this
                                    # (a HashSet is not [long]) and report zero
                                    # distinct identities for every path.
                                    $jobRow['IdentitySet'].UnionWith($j.Value)
                                }
                                elseif ($j.Key -eq 'IdentityTruncated') {
                                    if ($j.Value -gt $jobRow['IdentityTruncated']) {
                                        $jobRow['IdentityTruncated'] = $j.Value
                                    }
                                }
                                elseif ($j.Key -eq 'JobStartTicks') {
                                    # Absolute instants, so MIN/MAX rather than
                                    # sum. Adding two timestamps produces a
                                    # number with no meaning at all, and the
                                    # generic branch below would happily do it.
                                    if (($jobRow['JobStartTicks'] -eq 0) -or ($j.Value -lt $jobRow['JobStartTicks'])) {
                                        $jobRow['JobStartTicks'] = $j.Value
                                    }
                                }
                                elseif ($j.Key -eq 'JobEndTicks') {
                                    if ($j.Value -gt $jobRow['JobEndTicks']) {
                                        $jobRow['JobEndTicks'] = $j.Value
                                    }
                                }
                                elseif ($jobRow[$j.Key] -is [long]) {
                                    $jobRow[$j.Key] += $j.Value
                                }
                            }
                        }

                        foreach ($t in $jobResult.Telemetry.GetEnumerator()) {
                            if ($t.Key -eq 'AceCountMax') {
                                if ($t.Value -gt $telemetry['AceCountMax']) {
                                    $telemetry['AceCountMax'] = $t.Value
                                }
                            }
                            elseif ($t.Key -eq 'IdentitySet') {
                                # Union across every job, so the distinct count
                                # is the tree's, not one worker's. '+=' on a
                                # HashSet would build an array of sets.
                                $telemetry['IdentitySet'].UnionWith($t.Value)
                            }
                            elseif ($t.Key -eq 'IdentityTruncated') {
                                # A flag: set if ANY job hit the cap.
                                if ($t.Value -gt $telemetry['IdentityTruncated']) {
                                    $telemetry['IdentityTruncated'] = $t.Value
                                }
                            }
                            elseif ($t.Key -in 'WalkedPath', 'SeedOnly', 'JobStartTicks', 'JobEndTicks') {
                                # Labels and absolute instants, not counters.
                                # Consumed by the per-path breakdown above;
                                # adding them here would concatenate strings
                                # into the totals and sum raw clock readings
                                # into a number with no meaning.
                                continue
                            }
                            elseif ($t.Key -in 'SampleEvery', 'SampleWarmup') {
                                # Constants, identical in every worker. Summing
                                # them would report 'every 256th item' for a run
                                # with four workers.
                                $telemetry[$t.Key] = $t.Value
                            }
                            elseif ($telemetry.ContainsKey($t.Key)) {
                                $telemetry[$t.Key] += $t.Value
                            }
                        }
                    }
                    #endregion

                    if ($CollectTestedPaths) {
                        foreach ($j in $jobResult.TestedInheritedFilesAndFolders) {
                            foreach ($i in $j.GetEnumerator()) { $testedInheritedFilesAndFolders[$i.Key] = $i.Value }
                        }
                    }
                    foreach ($j in $jobResult.IncorrectInheritedAcl) {
                        if ($DetailedLog) {
                            foreach ($i in $j.GetEnumerator()) { $IncorrectInheritedAcl[$i.Key] = $i.Value }
                        }
                        else { $IncorrectInheritedAcl.Add($j) }
                    }
                    foreach ($j in $jobResult.UnreadableAcl) {
                        if ($DetailedLog) {
                            foreach ($i in $j.GetEnumerator()) { $unreadableAcl[$i.Key] = $i.Value }
                        }
                        else { $unreadableAcl.Add($j) }
                    }
                }

                if ($IncorrectInheritedAcl.Count -ne 0) {
                    [PSCustomObject]@{
                        DateTime    = Get-Date
                        Type        = 'Warning'
                        Name        = 'Inherited permissions incorrect'
                        Description = "All folders that don't have permissions assigned to them in the worksheet 'Permissions' are supposed to inherit their permissions from the parent folder. Files can only inherit permissions from the parent folder and are not allowed to have explicit permissions."
                        Value       = if ($DetailedLog) { $IncorrectInheritedAcl } else { $IncorrectInheritedAcl.ToArray() }
                    }
                }
            }
        }
        catch { throw "Failed checking/setting the inheritance on folders and files: $_" }
        #endregion

        #region Report folders or files whose ACL could not be read
        # Populated by the non-inherited matrix loop (any action) and the
        # inherited walker (Check/Fix), so emit outside the Action gate above.
        if ($unreadableAcl.Count -ne 0) {
            [PSCustomObject]@{
                DateTime    = Get-Date
                Type        = 'Warning'
                Name        = 'ACL could not be read'
                Description = "The permissions of these folders or files could not be read on the remote machine (for example the security descriptor is corrupt or the item is locked by another process). They were not checked or corrected and need manual attention."
                Value       = if ($DetailedLog) { $unreadableAcl } else { $unreadableAcl.ToArray() }
            }
        }
        #endregion

        #region Emit execution telemetry
        <#
         Type 'Telemetry' is NOT a check. Invoke-PermissionMatrixProcessHC
         splits it out of the result stream and parks it on the matrix
         object's 'Telemetry' property, so it never reaches $matrix.Check and
         never renders as a card. Anything that walks 'Check' (the pass/fail
         tally, the summary mail, the issue report) is therefore unaffected by
         its presence.

         Ticks are converted to milliseconds here, once, using the frequency
         of the machine that produced them — the remote file server. Doing it
         on the orchestrator would be wrong on any host with a different
         Stopwatch.Frequency.
        #>
        $tickToMs = 1000.0 / [System.Diagnostics.Stopwatch]::Frequency

        $itemsWalked = $telemetry['FoldersWalked'] + $telemetry['FilesWalked']

        $round = { param($v) [math]::Round($v, 2) }

        <#
         Mean cost of one ACL read, in milliseconds.

         PREFERS THE STRIDE POOL. Stride samples are spread evenly across the
         whole walk, so their mean is representative. Warm-up samples are the
         first N items in order, which on a large tree are both unrepresentative
         (coldest) and a tiny slice of the data — including them measured up to
         1.9x the true per-item cost in testing.

         Falls back to the combined pool only when there are too few stride
         samples to mean anything, which is exactly the small-tree case the
         warm-up was added for. 30 is the conventional floor for treating a
         sample mean as usable.

         Returns the mean and the basis, so the JSON can say which pool it used
         rather than leaving the reader to guess.
        #>
        $strideFloor = 30

        <#
         Generalised over the stage name so the read, projection and comparison
         timers all get the same stride-preferred rule from one implementation.
         Three copies of this logic would be three places for the fallback
         threshold to drift apart.
        #>
        $poolMean = {
            param($Stage, $Counters)

            $strideN = $Counters["Acl${Stage}StrideSamples"]
            $warmN = $Counters["Acl${Stage}WarmupSamples"]

            if ($strideN -ge $strideFloor) {
                return @{
                    Ms      = ($Counters["Acl${Stage}StrideTicks"] * $tickToMs) / $strideN
                    Basis   = 'stride'
                    Samples = $strideN
                }
            }

            $totalN = $strideN + $warmN

            if ($totalN -gt 0) {
                return @{
                    Ms      = (
                        ($Counters["Acl${Stage}StrideTicks"] + $Counters["Acl${Stage}WarmupTicks"]) * $tickToMs
                    ) / $totalN
                    Basis   = 'warmup+stride'
                    Samples = $totalN
                }
            }

            return @{ Ms = 0; Basis = 'none'; Samples = 0 }
        }

        $aclRead = & $poolMean 'Read' $telemetry
        $aclProject = & $poolMean 'Project' $telemetry
        $aclCompare = & $poolMean 'Compare' $telemetry

        <#
         SELF-AUDIT OF THE TELEMETRY ITSELF.

         Every cost counter above measures one specific operation. Nothing
         measures whether those operations add up to the time the job actually
         took, and without that the reader has no way to know whether a
         breakdown is a breakdown or a rounding error. On a row where the
         counters cover the work, AclReadMsPerItem is the answer to 'why was
         this slow'. On a row where they cover 15% of it, the same field is
         a distraction, and looks identical.

         So the figure is emitted rather than left for the reader to derive.
         AccountedPct is a QUALITY INDICATOR FOR THE MEASUREMENT, not a
         physical breakdown of the wall clock:

           - Well below 100 means the counters do not explain this row. Read
             it as 'the telemetry is not measuring the expensive thing here',
             and do not draw conclusions from the cost fields until it is.
           - Around 100 means the counters cover the work and the breakdown
             can be trusted.
           - ABOVE 100 IS NORMAL AND NOT AN ERROR. The millisecond totals sum
             concurrent work across the walker runspaces while WallClockMs is
             elapsed time, so overlapping work is counted more than once. See
             the matching entry in the field reference.

         UnaccountedMs is therefore allowed to be negative, and is deliberately
         NOT clamped: a clamp would hide the concurrency case behind a zero and
         make the two very different situations look the same.

         Cost: three arithmetic operations, once per Settings row, on values
         that already exist. Nothing is added to the walk.
        #>
        $wallClockMs = (
            [System.Diagnostics.Stopwatch]::GetTimestamp() - $telemetryStart
        ) * $tickToMs

        <#
         STRAGGLER ANALYSIS.

         A Settings row cannot finish before its slowest job. If one matrix
         folder holds most of the tree, the row is bound by that one folder and
         adding throttle does nothing — the fix is to split the folder, which
         is a completely different action to "the storage is slow".

         Two figures separate those cases:

           JobStragglerPct     the longest single job as a share of the row.
                               Near 100 means the row IS that job.
           JobConcurrencyMean  total job time divided by elapsed time, i.e. how
                               many workers were busy on average. Compare it to
                               JobThrottleLimit: far below means the throttle is
                               not the constraint and raising it will not help.

         Derived from the per-path rows that already exist, once per Settings
         row. Nothing is added to the walk.
        #>
        $jobRows = @($pathTelemetry.Values | Where-Object { $_['JobEndTicks'] -gt 0 })
        $jobSpansMs = @($jobRows | ForEach-Object {
                ($_['JobEndTicks'] - $_['JobStartTicks']) * $tickToMs
            })

        $jobLongestMs = 0
        $jobLongestPath = ''

        foreach ($jr in $jobRows) {
            $spanMs = ($jr['JobEndTicks'] - $jr['JobStartTicks']) * $tickToMs
            if ($spanMs -gt $jobLongestMs) {
                $jobLongestMs = $spanMs
                $jobLongestPath = $jr['Path']
            }
        }

        $jobSpanSumMs = ($jobSpansMs | Measure-Object -Sum).Sum
        if ($null -eq $jobSpanSumMs) { $jobSpanSumMs = 0 }

        $accountedMs = (
            ($aclRead.Ms * $itemsWalked) +
            ($aclProject.Ms * $itemsWalked) +
            ($aclCompare.Ms * $itemsWalked) +
            ($telemetry['AclWriteTicks'] * $tickToMs) +
            ($telemetry['EnumerateTicks'] * $tickToMs) +
            ($telemetry['MatrixFolderReadTicks'] * $tickToMs) +
            ($telemetry['MatrixFolderWriteTicks'] * $tickToMs)
        )

        [PSCustomObject]@{
            DateTime    = Get-Date
            Type        = 'Telemetry'
            Name        = 'Execution telemetry'
            Description = 'Volume and cost counters for this Settings row. Compare the same path between runs: a duration that grew while ItemsWalked stayed flat points at the storage or the ACLs, not at the amount of data.'
            Value       = [ordered]@{
                Path                   = $Path
                Action                 = $Action
                ComputerName           = $env:COMPUTERNAME
                WallClockMs            = & $round $wallClockMs

                # --- Does the breakdown below explain the wall clock? ---
                # Read AccountedPct FIRST. It says whether the cost fields
                # further down are worth reading at all on this row.
                AccountedMs            = & $round $accountedMs
                UnaccountedMs          = & $round ($wallClockMs - $accountedMs)
                # --- Was this row bound by one slow job? ---
                JobCount               = $jobRows.Count
                JobWallClockMsMax      = & $round $jobLongestMs
                JobWallClockMsSum      = & $round $jobSpanSumMs
                JobLongestPath         = $jobLongestPath
                JobStragglerPct        = & $round (
                    $(if ($wallClockMs -gt 0) { ($jobLongestMs / $wallClockMs) * 100 }
                        else { 0 })
                )
                JobConcurrencyMean     = & $round (
                    $(if ($wallClockMs -gt 0) { $jobSpanSumMs / $wallClockMs }
                        else { 0 })
                )

                AccountedPct           = & $round (
                    $(if ($wallClockMs -gt 0) { ($accountedMs / $wallClockMs) * 100 }
                        else { 0 })
                )

                # --- Volume: how much was there? ---
                ItemsWalked            = $itemsWalked
                FoldersWalked          = $telemetry['FoldersWalked']
                FilesWalked            = $telemetry['FilesWalked']
                MatrixFolders          = $telemetry['MatrixFolders']
                MatrixFoldersCreated   = $telemetry['MatrixFoldersCreated']

                # --- Cost: what did touching it take? ---
                # AclReadMsPerItem is a SAMPLED mean (1 item in SampleEvery),
                # not a total. It is the number to compare between runs: it
                # isolates the per-operation cost of the storage from the
                # amount of data, which a duration alone cannot do.
                # AclReadMsEstimated scales it back up to a whole-job figure —
                # useful for apportioning the wall clock, but it is an
                # estimate, and it is named so nobody mistakes it for measured.
                SampleEvery            = $telemetry['SampleEvery']
                SampleWarmup           = $telemetry['SampleWarmup']
                AclReadSamples         = $aclRead.Samples
                AclReadStrideSamples   = $telemetry['AclReadStrideSamples']
                AclReadWarmupSamples   = $telemetry['AclReadWarmupSamples']
                AclReadBasis           = $aclRead.Basis
                AclReadMsPerItem       = & $round $aclRead.Ms
                AclReadMsEstimated     = & $round ($aclRead.Ms * $itemsWalked)

                # The two stages that scale with ACE count rather than item
                # count. On a row where these dominate, the answer is the ACLs
                # (how many entries, how many distinct identities) and not the
                # disk — which is the opposite conclusion to the one
                # AclReadMsPerItem alone would suggest.
                AclProjectSamples      = $aclProject.Samples
                AclProjectBasis        = $aclProject.Basis
                AclProjectMsPerItem    = & $round $aclProject.Ms
                AclProjectMsEstimated  = & $round ($aclProject.Ms * $itemsWalked)
                AclCompareSamples      = $aclCompare.Samples
                AclCompareBasis        = $aclCompare.Basis
                AclCompareMsPerItem    = & $round $aclCompare.Ms
                AclCompareMsEstimated  = & $round ($aclCompare.Ms * $itemsWalked)
                AclWrites              = $telemetry['AclWrites']
                AclWriteMs             = & $round ($telemetry['AclWriteTicks'] * $tickToMs)
                AclWriteMsPerItem      = & $round (
                    $(if ($telemetry['AclWrites']) {
                            ($telemetry['AclWriteTicks'] * $tickToMs) / $telemetry['AclWrites']
                        }
                        else { 0 })
                )
                EnumeratedDirs         = $telemetry['EnumeratedDirs']
                EnumerateMs            = & $round ($telemetry['EnumerateTicks'] * $tickToMs)

                MatrixFolderReads      = $telemetry['MatrixFolderReads']
                MatrixFolderReadMs     = & $round ($telemetry['MatrixFolderReadTicks'] * $tickToMs)
                MatrixFolderWrites     = $telemetry['MatrixFolderWrites']
                MatrixFolderWriteMs    = & $round ($telemetry['MatrixFolderWriteTicks'] * $tickToMs)

                # --- Convergence: is the run idempotent? ---
                # On a settled tree these trend to zero. A path that reports
                # the same non-zero IncorrectItems every night is being
                # rewritten every night, and is worth investigating before
                # blaming the storage.
                IncorrectItems         = $telemetry['IncorrectItems']
                MatrixFoldersIncorrect = $telemetry['MatrixFoldersIncorrect']
                AclReadDenied          = $telemetry['AclReadDenied']
                AclReadFailed          = $telemetry['AclReadFailed']
                AclWriteDenied         = $telemetry['AclWriteDenied']

                # --- Identity census: WHY is the ACL expensive to interpret? ---
                # AceCountMean says how many entries. These say how many
                # DISTINCT accounts those entries name, and how many of them
                # Windows could not resolve at all.
                IdentityDistinct       = $telemetry['IdentitySet'].Count
                IdentityObservations   = $telemetry['IdentityObservations']
                IdentityTruncated      = [Boolean]$telemetry['IdentityTruncated']
                AceUnresolvedSids      = $telemetry['AceUnresolvedSids']
                AceUnresolvedItems     = $telemetry['AceUnresolvedItems']
                AceUnresolvedPct       = & $round (
                    $(if ($telemetry['IdentityObservations'] -gt 0) {
                            ($telemetry['AceUnresolvedSids'] / $telemetry['IdentityObservations']) * 100
                        }
                        else { 0 })
                )

                # --- ACE census: are the ACLs themselves growing? ---
                # Sampled in the walker, exact for the matrix folders. A rising
                # AceCountMean on one tree while another stays flat is what
                # non-idempotent ACL application looks like.
                AceCountMean           = & $round (
                    $(if ($telemetry['AceCountItems']) {
                            $telemetry['AceCountTotal'] / $telemetry['AceCountItems']
                        }
                        else { 0 })
                )
                AceCountMax            = $telemetry['AceCountMax']
                AceCountItems          = $telemetry['AceCountItems']

                # --- Per-path breakdown ---
                # One entry per matrix folder, so a regression can be localised
                # to the folder that caused it rather than only to this Settings
                # row. Because the walker jobs partition the tree, these sum
                # back to the totals above.
                #
                # Sorted by measured cost descending: the folder that got slower
                # is the reason anyone opens this file, so it should be the first
                # thing they read rather than something they have to sort for.
                Paths                  = @(
                    $pathTelemetry.Values |
                    Sort-Object -Property @{
                        Expression = {
                            $_['AclReadWarmupTicks'] + $_['AclReadStrideTicks'] +
                            $_['AclWriteTicks'] +
                            $_['EnumerateTicks'] + $_['MatrixFolderReadTicks'] +
                            $_['MatrixFolderWriteTicks']
                        }
                        Descending = $true
                    } |
                    ForEach-Object {
                        $row = $_
                        $rowItems = $row['FoldersWalked'] + $row['FilesWalked']

                        # Same stride-preferred rule as the Settings-row total,
                        # from the same helper. Per-path pools are smaller, so
                        # the fallback fires more often here — hence reporting
                        # the basis per row rather than once for the whole
                        # setting.
                        $rowRead = & $poolMean 'Read' $row
                        $rowProject = & $poolMean 'Project' $row
                        $rowCompare = & $poolMean 'Compare' $row

                        [ordered]@{
                            Path                   = $row['Path']
                            ItemsWalked            = $rowItems
                            FoldersWalked          = $row['FoldersWalked']
                            FilesWalked            = $row['FilesWalked']
                            EnumeratedDirs         = $row['EnumeratedDirs']

                            # Cost attributable to this folder's subtree. The
                            # read figure is sampled, so it carries its own
                            # sample count for the same reason as the total.
                            # How long this subtree's job actually took, and
                            # when it started relative to the Settings row. A
                            # job that starts late was queued behind the
                            # throttle; a job that starts at zero and runs to
                            # the end IS the row.
                            IdentityDistinct       = $row['IdentitySet'].Count
                            AceUnresolvedSids      = $row['AceUnresolvedSids']
                            JobWallClockMs         = & $round (
                                ($row['JobEndTicks'] - $row['JobStartTicks']) * $tickToMs
                            )
                            JobStartOffsetMs       = & $round (
                                $(if ($row['JobStartTicks'] -gt 0) {
                                        ($row['JobStartTicks'] - $telemetryStart) * $tickToMs
                                    }
                                    else { 0 })
                            )
                            AclReadSamples         = $rowRead.Samples
                            AclReadBasis           = $rowRead.Basis
                            AclReadMsPerItem       = & $round $rowRead.Ms
                            AclProjectBasis        = $rowProject.Basis
                            AclProjectMsPerItem    = & $round $rowProject.Ms
                            AclCompareBasis        = $rowCompare.Basis
                            AclCompareMsPerItem    = & $round $rowCompare.Ms
                            AclWrites              = $row['AclWrites']
                            AclWriteMs             = & $round ($row['AclWriteTicks'] * $tickToMs)
                            EnumerateMs            = & $round ($row['EnumerateTicks'] * $tickToMs)
                            MatrixFolderReadMs     = & $round ($row['MatrixFolderReadTicks'] * $tickToMs)
                            MatrixFolderWriteMs    = & $round ($row['MatrixFolderWriteTicks'] * $tickToMs)

                            IncorrectItems         = $row['IncorrectItems']
                            MatrixFoldersIncorrect = $row['MatrixFoldersIncorrect']
                            AclReadDenied          = $row['AclReadDenied']
                            AclReadFailed          = $row['AclReadFailed']
                            AclWriteDenied         = $row['AclWriteDenied']

                            AceCountMean           = & $round (
                                $(if ($row['AceCountItems']) {
                                        $row['AceCountTotal'] / $row['AceCountItems']
                                    }
                                    else { 0 })
                            )
                            AceCountMax            = $row['AceCountMax']

                            # False means this folder's ACL was checked but its
                            # subtree was never walked — it is ignored, or every
                            # child belongs to another matrix folder. Explains a
                            # row with cost but zero ItemsWalked.
                            Walked                 = [bool]$row['Walked']
                        }
                    }
                )
            }
        }
        #endregion
    }
    catch { throw "Failed setting the permissions: $_" }
}