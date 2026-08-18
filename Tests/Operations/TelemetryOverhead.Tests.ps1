#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

<#
    The measurements behind the three counter rules documented in
    SetPermissions.ps1, kept executable so the rules can be re-checked on a new
    PowerShell version instead of trusted on faith.

    WHY THIS EXISTS
    The telemetry sits in a loop that runs millions of times per night. The
    original implementation used a parent-scope hashtable and timed every item,
    which measured at ~2,900ns per item — enough to add minutes to a large
    tree, which would have made the diagnostics themselves a cause of the
    slowdown they were added to investigate.

    Hoisting the lookup into a local and sampling the timers brought that to
    ~280ns. These tests assert the shape stays that way.

    NOTE ON THRESHOLDS
    Absolute nanosecond budgets would be flaky on shared CI. Every assertion
    below is therefore a RATIO between two shapes measured back-to-back on the
    same host, which is stable even when the host is not.
#>

Describe 'Telemetry instrumentation overhead' -Tag 'Performance' {
    BeforeAll {
        # Parsed once and shared: the structural rules below are assertions
        # about the shipped source, so they read it directly rather than
        # measuring a stand-in shape.
        $script:RepoRoot = $null
        $candidate = [System.IO.DirectoryInfo]::new($PSScriptRoot)
        while ($candidate) {
            if (Test-Path -LiteralPath (Join-Path $candidate.FullName 'Scripts\Operations\SetPermissions.ps1')) {
                $script:RepoRoot = $candidate.FullName; break
            }
            $candidate = $candidate.Parent
        }
        if (-not $script:RepoRoot) {
            throw "Could not locate the repository root above '$PSScriptRoot'."
        }
        $script:SetPermissionsAst = [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path $script:RepoRoot 'Scripts\Operations\SetPermissions.ps1'), [ref]$null, [ref]$null
        )

        # Per round, per shape. Five interleaved rounds keeps total work similar
        # to the old single 200k pass while spreading it across the window.
        $script:Iterations = 50000
        $script:Rounds = 5

        <#
         MEASURES ALL SHAPES TOGETHER, ROUND-ROBIN, AND KEEPS THE MINIMUM.

         An earlier version measured each shape once, sequentially, and compared
         the results. That assumes host load is constant across the two
         measurements, which on a shared machine is simply false — and it fails
         arbitrarily badly, not marginally. A real run reported the SAMPLED
         shape at 6,171 ns/item against the unsampled shape at 631, i.e.
         sampling appearing 10x slower than not sampling, which is impossible
         since it does strictly less work. Everything else in that run was
         inflated 5-8x too. The machine was busy, and whichever shape happened
         to land in the busy window lost.

         Two corrections, both standard practice for benchmarking on hardware
         you do not control:

         - INTERLEAVE. Every round measures every shape, and the order reverses
           on alternate rounds, so no shape can systematically own the quiet
           part of the window.
         - TAKE THE MINIMUM, not the mean. Contention can only ADD time, never
           remove it, so the fastest observed run is the closest estimate of the
           true cost. A mean is dragged around by whatever else the host was
           doing; a minimum is not.

         Verified to preserve correct ordering with six competing CPU-bound
         processes running, where the sequential version inverted.
        #>
        function Measure-ShapesMinNs {
            param(
                [Parameter(Mandatory)]
                [System.Collections.Specialized.OrderedDictionary]$Shapes,
                [int]$Iterations = 50000,
                [int]$Rounds = 5
            )

            $best = @{}

            foreach ($key in $Shapes.Keys) {
                $best[$key] = [double]::MaxValue
                # Warm up so JIT and the PowerShell compiler are not timed.
                $null = & $Shapes[$key] 2000
            }

            foreach ($round in 1..$Rounds) {
                $keys = @($Shapes.Keys)
                if ($round % 2 -eq 0) { [array]::Reverse($keys) }

                foreach ($key in $keys) {
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    $null = & $Shapes[$key] $Iterations
                    $sw.Stop()

                    $ns = $sw.Elapsed.TotalMilliseconds / $Iterations * 1000000
                    if ($ns -lt $best[$key]) { $best[$key] = $ns }
                }
            }

            return $best
        }

        # Parent-scope counter read from inside a nested function — the shape
        # the instrumentation must NOT use.
        $script:UnhoistedShape = {
            param($n)
            $telemetry = @{ A = 0L; B = 0L; C = 0L }
            function Inner {
                param($n)
                for ($i = 0; $i -lt $n; $i++) {
                    $telemetry['A']++; $telemetry['B']++; $telemetry['C']++
                }
            }
            Inner -n $n
            $telemetry
        }

        # Same counters, reference hoisted into a local first.
        $script:HoistedShape = {
            param($n)
            $telemetry = @{ A = 0L; B = 0L; C = 0L }
            function Inner {
                param($n)
                $t = $telemetry
                for ($i = 0; $i -lt $n; $i++) {
                    $t['A']++; $t['B']++; $t['C']++
                }
            }
            Inner -n $n
            $telemetry
        }

        # Timing every item.
        $script:UnsampledShape = {
            param($n)
            $telemetry = @{ Items = 0L; Samples = 0L; Ticks = 0L }
            function Inner {
                param($n)
                $t = $telemetry
                for ($i = 0; $i -lt $n; $i++) {
                    $t['Items']++
                    $ts = [System.Diagnostics.Stopwatch]::GetTimestamp()
                    $t['Samples']++
                    $t['Ticks'] += ([System.Diagnostics.Stopwatch]::GetTimestamp() - $ts)
                }
            }
            Inner -n $n
            $telemetry
        }

        # Timing 1 item in 64 — the shipped shape.
        $script:SampledShape = {
            param($n)
            $telemetry = @{ Items = 0L; Samples = 0L; Ticks = 0L }
            $sampleMask = 63
            $script:counter = 0
            function Inner {
                param($n, $sampleMask)
                $t = $telemetry
                for ($i = 0; $i -lt $n; $i++) {
                    $t['Items']++
                    if ((++$script:counter -band $sampleMask) -eq 0) {
                        $ts = [System.Diagnostics.Stopwatch]::GetTimestamp()
                        $t['Samples']++
                        $t['Ticks'] += ([System.Diagnostics.Stopwatch]::GetTimestamp() - $ts)
                    }
                }
            }
            Inner -n $n -sampleMask $sampleMask
            $telemetry
        }
    }

    Context 'Rule 1: the counters are accessed through a hoisted local' {
        <#
         ASSERTED FROM THE SOURCE, NOT FROM A STOPWATCH.

         This rule was previously guarded by timing two shapes and asserting the
         hoisted one was at least 2x cheaper. That assertion failed
         intermittently on a shared host and could not be stabilised: across
         runs of IDENTICAL code it produced ratios of 4.69x, 2.81x, 2.70x,
         1.70x and 1.17x. Three separate methodology fixes (interleaving,
         minimum-of-N, a relaxed threshold) removed the gross inversions but
         never made the ratio itself reliable.

         The benefit of hoisting was established once, by measurement, and is
         recorded in the notes in SetPermissions.ps1. It does not need
         re-establishing on every CI run, and a machine under load cannot
         re-establish it anyway. What genuinely needs guarding is that the
         shipped code still FOLLOWS the rule — and that is a property of the
         source, so it can be asserted exactly, in milliseconds, on any host.

         The measured comparison is still printed by the informational test at
         the end of this file, for anyone who wants the numbers.
        #>
        BeforeAll {
            $script:WalkerAst = $script:SetPermissionsAst.FindAll({ param($n)
                    ($n -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and
                    ($n.Name -eq 'Get-FolderContentHC')
                }, $true) | Select-Object -First 1
        }

        It 'finds the walker function' {
            # Guards the guard: without this the assertions below would pass
            # vacuously against a null AST.
            $script:WalkerAst | Should -Not -BeNullOrEmpty
        }

        It 'references the parent-scope counter exactly once, to hoist it' {
            # Every additional bare '$telemetry' inside this function is a
            # dynamic-scope lookup in a loop that runs millions of times per
            # night. One reference is the hoist; more than one is the
            # regression this rule exists to prevent.
            $bare = $script:WalkerAst.FindAll({ param($n)
                    ($n -is [System.Management.Automation.Language.VariableExpressionAst]) -and
                    ($n.VariablePath.UserPath -eq 'telemetry')
                }, $true)

            $bare.Count | Should -Be 1 -Because 'the only bare $telemetry should be the hoist "$t = $telemetry"'
        }

        It 'hoists into a local before the enumeration begins' {
            $assignments = $script:WalkerAst.FindAll({ param($n)
                    ($n -is [System.Management.Automation.Language.AssignmentStatementAst]) -and
                    ($n.Left.Extent.Text -eq '$t') -and
                    ($n.Right.Extent.Text -eq '$telemetry')
                }, $true)

            $assignments.Count | Should -Be 1
        }

        It 'indexes only the hoisted local for counters' {
            $counterTargets = $script:WalkerAst.FindAll({ param($n)
                    $n -is [System.Management.Automation.Language.IndexExpressionAst]
                }, $true) |
            Where-Object { $_.Index.Extent.Text -match "Walked|Acl|Ace|Enumerate|Incorrect|Sample" } |
            ForEach-Object { $_.Target.Extent.Text } |
            Sort-Object -Unique

            $counterTargets | Should -Be @('$t')
        }

        It 'still mutates the parent object' {
            # The behavioural half of the rule: a hoisted reference only works
            # because a hashtable is a reference type. If this ever fails, every
            # counter silently reports zero.
            $result = & $HoistedShape 1000
            $result.A | Should -Be 1000
        }
    }

    Context 'Rule 2: the expensive signals are sampled' {
        <#
         Structural for the same reason as Rule 1. The measured win was 2.0x to
         6.6x across three hosts; the assertion here is that the sampling gate
         still exists and still wraps the timers, which is what would actually
         break if someone "simplified" the code.
        #>
        It 'computes a sample decision from the warm-up and the stride mask' {
            $decision = $script:WalkerAst.FindAll({ param($n)
                    ($n -is [System.Management.Automation.Language.AssignmentStatementAst]) -and
                    ($n.Left.Extent.Text -eq '$isSample')
                }, $true) | Select-Object -First 1

            $decision | Should -Not -BeNullOrEmpty
            $decision.Right.Extent.Text | Should -Match 'sampleWarmup|isWarmupSample'
            $decision.Right.Extent.Text | Should -Match 'sampleMask'
        }

        It 'gates every sample accumulator behind that decision' {
            <#
             Checks each ACCUMULATOR individually rather than counting gates.

             Counting was the obvious approach and it does not work: an earlier
             version asserted 'more than two if(isSample) blocks', and a
             mutation that ungated one timer left three, so the test passed
             while the code was broken. Verified by deliberately breaking the
             source and confirming the test now fails.
            #>
            $accumulators = $script:WalkerAst.FindAll({ param($n)
                    ($n -is [System.Management.Automation.Language.AssignmentStatementAst]) -and
                    ($n.Left.Extent.Text -match 'AclRead(Warmup|Stride)(Ticks|Samples)')
                }, $true)

            $accumulators.Count | Should -BeGreaterThan 0 -Because 'the read timer must accumulate somewhere'

            foreach ($acc in $accumulators) {
                $gated = $false
                $node = $acc.Parent

                while ($node -and $node -ne $script:WalkerAst) {
                    if (
                        ($node -is [System.Management.Automation.Language.IfStatementAst]) -and
                        ($node.Clauses[0].Item1.Extent.Text -match 'isSample')
                    ) { $gated = $true; break }
                    $node = $node.Parent
                }

                $gated | Should -BeTrue -Because "'$($acc.Left.Extent.Text)' must sit inside an if(`$isSample) gate, or the timer runs on every item"
            }
        }

        It 'routes warm-up and stride samples into separate pools' {
            <#
             The bias fix: pooling the contiguous warm-up with the stride
             samples measured up to 1.9x the true per-item cost.

             Asserts the BRANCH, not just the presence of both names. A rename
             or a partial merge leaves the names scattered around the file, so a
             text search passes on broken code — again confirmed by mutation.
            #>
            $branches = $script:WalkerAst.FindAll({ param($n)
                    ($n -is [System.Management.Automation.Language.IfStatementAst]) -and
                    ($n.Clauses[0].Item1.Extent.Text -match 'isWarmupSample') -and
                    ($null -ne $n.ElseClause)
                }, $true)

            $branches.Count | Should -BeGreaterThan 0 -Because 'the two pools must be chosen by an if/else on $isWarmupSample'

            foreach ($branch in $branches) {
                $warmupBody = $branch.Clauses[0].Item2.Extent.Text
                $strideBody = $branch.ElseClause.Extent.Text

                # Each branch must touch its OWN pool and NOT the other one.
                # Asserting only 'the warm-up branch mentions AclReadWarmup' is
                # not enough: each branch writes both a Samples and a Ticks
                # counter, so moving just the Samples line across still left the
                # Ticks line matching and the test passed on broken code.
                # Mutation testing caught this; the negative half is the fix.
                $warmupBody | Should -Match 'AclReadWarmup'
                $warmupBody | Should -Not -Match 'AclReadStride' -Because 'the warm-up branch must not write the stride pool'

                $strideBody | Should -Match 'AclReadStride'
                $strideBody | Should -Not -Match 'AclReadWarmup' -Because 'the stride branch must not write the warm-up pool'
            }
        }

        It 'samples at exactly the configured rate' {
            $expectedSamples = [math]::Floor($Iterations / 64)

            $result = & $SampledShape $Iterations

            $result.Items | Should -Be $Iterations
            $result.Samples | Should -Be $expectedSamples
            $result.Samples | Should -BeLessThan ($Iterations / 10) -Because 'sampling must be a real reduction, not near-1:1'
        }
    }

    Context 'Measured overhead (informational)' {
        <#
         Prints the numbers, asserts nothing about them.

         The ratios are genuinely useful to see — they are how the design rules
         were arrived at — but on a shared host they are not a pass/fail signal.
         Observed across runs of identical code: 4.69x, 2.81x, 2.70x, 1.70x,
         1.17x for Rule 1. Anything asserted on that range is either so loose it
         catches nothing or so tight it fails on a busy machine.

         So: measure, print, move on. The rules themselves are enforced
         structurally above, where the answer does not depend on what else the
         machine is doing.
        #>
        It 'reports the current cost of each shape' {
            $r1 = Measure-ShapesMinNs -Iterations $Iterations -Rounds $Rounds -Shapes ([ordered]@{
                    unhoisted = $UnhoistedShape
                    hoisted   = $HoistedShape
                })
            $r2 = Measure-ShapesMinNs -Iterations $Iterations -Rounds $Rounds -Shapes ([ordered]@{
                    unsampled = $UnsampledShape
                    sampled   = $SampledShape
                })

            Write-Host ("      hoisting : {0,6:N0} -> {1,6:N0} ns/item ({2:N2}x)" -f `
                    $r1.unhoisted, $r1.hoisted, ($r1.unhoisted / $r1.hoisted))
            Write-Host ("      sampling : {0,6:N0} -> {1,6:N0} ns/item ({2:N2}x)" -f `
                    $r2.unsampled, $r2.sampled, ($r2.unsampled / $r2.sampled))
            Write-Host  '      (informational only - a loaded host compresses these ratios)'

            # The only thing worth asserting: the harness produced numbers at
            # all. A zero or a NaN would mean the measurement itself broke.
            $r1.hoisted | Should -BeGreaterThan 0
            $r2.sampled | Should -BeGreaterThan 0
        }
    }

    Context 'Sampling logic' {
        <#
         WHY THERE IS NO TIGHT ASSERTION AGAINST WALL-CLOCK TIMING HERE.

         Two earlier versions of this context tried and both flaked:

         - A flat +/-10% band. Picked from five passing runs on one machine
           rather than from the estimator's precision. Failed elsewhere at 1.112.
         - A band computed from the sample's own standard error. Better
           reasoned, still wrong: the gap between the sampled mean and the full
           mean is driven by RARE EXPENSIVE OUTLIERS in the full population that
           the sample may miss entirely. When a sample happens to be
           homogeneous its SE collapses, the band narrows to +/-1.5%, and the
           test fails at a ratio of 0.976 — a perfectly good estimate. 2 in 10
           runs failed that way.

         The lesson is that a ~300-sample mean of a heavy-tailed timing
         distribution on a shared host cannot support a tight assertion, and no
         cleverness about the tolerance changes that.

         So the sampling LOGIC is tested exactly, with synthetic costs and no
         timing at all, where the arithmetic is fully determined. Real timing is
         asserted only loosely, to prove the plumbing measures something real.
         Exact where exactness is possible, loose where it is not.
        #>

        BeforeAll {
            # Replays the shipped sampling rule over a known cost sequence and
            # returns the pools, so a test can assert against an independently
            # computed expected value.
            function Invoke-SamplingRule {
                param(
                    # AllowEmptyCollection is required, not decorative: a
                    # Mandatory [double[]] rejects @() outright, and the
                    # zero-sample case is one of the behaviours under test.
                    [Parameter(Mandatory)]
                    [AllowEmptyCollection()]
                    [double[]]$Costs,
                    [int]$SampleMask = 63,
                    [int]$SampleWarmup = 300
                )

                $counter = 0
                $stride = [System.Collections.Generic.List[double]]::new()
                $warmup = [System.Collections.Generic.List[double]]::new()

                foreach ($cost in $Costs) {
                    $counter++
                    $isWarmupSample = ($counter -le $SampleWarmup)
                    $isSample = (
                        $isWarmupSample -or (($counter -band $SampleMask) -eq 0)
                    )
                    if (-not $isSample) { continue }

                    if ($isWarmupSample) { $warmup.Add($cost) }
                    else { $stride.Add($cost) }
                }

                return @{ Stride = $stride; Warmup = $warmup }
            }

            # The stride-preferred rule from SetPermissions.ps1.
            function Get-PreferredMean {
                param($Pools, [int]$StrideFloor = 30)

                if ($Pools.Stride.Count -ge $StrideFloor) {
                    return @{
                        Mean  = ($Pools.Stride | Measure-Object -Average).Average
                        Basis = 'stride'
                    }
                }

                $combined = [System.Collections.Generic.List[double]]::new()
                $combined.AddRange($Pools.Stride)
                $combined.AddRange($Pools.Warmup)

                if ($combined.Count -eq 0) { return @{ Mean = 0; Basis = 'none' } }

                return @{
                    Mean  = ($combined | Measure-Object -Average).Average
                    Basis = 'warmup+stride'
                }
            }
        }

        It 'samples exactly the first N items and then every Nth' {
            # Fully determined: no timing, so this can assert the exact set.
            $costs = 1..1000 | ForEach-Object { [double]$_ }
            $pools = Invoke-SamplingRule -Costs $costs -SampleMask 63 -SampleWarmup 300

            $pools.Warmup.Count | Should -Be 300

            # Items 301..1000 sampled where the 1-based counter is a multiple of
            # 64: 320, 384, ... 960 -> 11 items.
            $pools.Stride.Count | Should -Be 11
            $pools.Stride[0] | Should -Be 320
            $pools.Stride[-1] | Should -Be 960
        }

        It 'recovers a known steady-state cost exactly from the stride pool' {
            # 20,000 items: 300 expensive cold ones, the rest at a flat 10.
            # The stride pool must return exactly 10, with no timing noise to
            # hide behind.
            $costs = @(
                1..300 | ForEach-Object { 100.0 }
                301..20000 | ForEach-Object { 10.0 }
            )

            $pools = Invoke-SamplingRule -Costs $costs
            $result = Get-PreferredMean -Pools $pools

            $result.Basis | Should -Be 'stride'
            $result.Mean | Should -Be 10.0
        }

        It 'does not let the contiguous warm-up skew a large-tree mean' {
            # The regression guard: pooling the warm-up with the stride samples
            # measured up to 1.9x the true value on real timings, because the
            # warm-up covers the first N items IN ORDER and those are the
            # coldest. This asserts the two rules genuinely differ, so nobody
            # can re-merge the pools without failing a test.
            $costs = @(
                1..300 | ForEach-Object { 100.0 }
                301..8000 | ForEach-Object { 10.0 }
            )

            $pools = Invoke-SamplingRule -Costs $costs
            $strideMean = ($pools.Stride | Measure-Object -Average).Average

            $combined = [System.Collections.Generic.List[double]]::new()
            $combined.AddRange($pools.Stride)
            $combined.AddRange($pools.Warmup)
            $mixedMean = ($combined | Measure-Object -Average).Average

            Write-Host ("      stride-only {0:N1} vs mixed-pool {1:N1} (true steady-state 10.0)" -f $strideMean, $mixedMean)

            $strideMean | Should -Be 10.0
            $mixedMean | Should -BeGreaterThan 20.0
        }

        It 'falls back to the combined pool on a small tree' {
            # 200 items yields 3 stride samples, below the floor, so the warm-up
            # pool is what rescues the estimate. This is the case the warm-up was
            # added for.
            #
            # The floor is named rather than repeated as a literal, so that
            # changing it in Get-PreferredMean cannot leave this assertion
            # asserting against a number the code no longer uses.
            $strideFloor = 30
            $costs = 1..200 | ForEach-Object { 7.0 }

            $pools = Invoke-SamplingRule -Costs $costs
            $result = Get-PreferredMean -Pools $pools -StrideFloor $strideFloor

            $pools.Stride.Count | Should -BeLessThan $strideFloor
            $result.Basis | Should -Be 'warmup+stride'
            $result.Mean | Should -Be 7.0
        }

        It 'reports no basis when nothing was sampled' {
            $result = Get-PreferredMean -Pools (Invoke-SamplingRule -Costs @())

            $result.Basis | Should -Be 'none'
            $result.Mean | Should -Be 0
        }

        It 'produces a sane figure against real timing' {
            <#
             The only test here that touches the clock, and it is deliberately
             loose. Its job is to prove the timing plumbing measures real work
             and lands in the right order of magnitude — not to pin down a
             precision that ~300 samples on a shared host cannot deliver.
            #>
            $sampleMask = 63
            $sampleWarmup = 300
            $items = 8000
            $targetMs = 0.04

            $counter = 0
            $strideVals = [System.Collections.Generic.List[double]]::new()
            $fullTicks = 0L

            for ($i = 0; $i -lt $items; $i++) {
                $counter++
                $isWarmupSample = ($counter -le $sampleWarmup)
                $isSample = (
                    $isWarmupSample -or (($counter -band $sampleMask) -eq 0)
                )

                $ts = [System.Diagnostics.Stopwatch]::GetTimestamp()
                $spin = [System.Diagnostics.Stopwatch]::StartNew()
                while ($spin.Elapsed.TotalMilliseconds -lt $targetMs) { }
                $elapsed = [System.Diagnostics.Stopwatch]::GetTimestamp() - $ts

                $fullTicks += $elapsed
                if ($isSample -and -not $isWarmupSample) {
                    $strideVals.Add([double]$elapsed)
                }
            }

            $strideVals.Count | Should -BeGreaterThan 30

            $tickToMs = 1000.0 / [System.Diagnostics.Stopwatch]::Frequency
            $sampledMeanMs = (($strideVals | Measure-Object -Average).Average) * $tickToMs
            $fullMeanMs = ($fullTicks / $items) * $tickToMs

            Write-Host ("      {0} stride samples, sampled {1:N4} ms vs full {2:N4} ms (target {3:N4})" -f `
                    $strideVals.Count, $sampledMeanMs, $fullMeanMs, $targetMs)

            <#
             Asserted against the KNOWN WORKLOAD, not against the full mean.

             A sampled/full ratio band was tried at +/-10%, at 3 standard
             errors, and finally at 0.5x-1.5x. All three flaked, the last one at
             a ratio of 1.761 — one sampled item catching a scheduling pause is
             enough to move a 300-sample mean that far. Any ratio band tight
             enough to be meaningful is loose enough to flake, because the two
             means differ by whichever outliers the sample happened to catch.

             The workload is a known 0.04ms spin, so bound the measurement
             against that instead. The lower bound is the real check: the timer
             must actually wrap the work rather than something adjacent, and it
             cannot come out below the work it contains. The upper bound is
             deliberately generous, since outliers only ever inflate.
            #>
            $sampledMeanMs | Should -BeGreaterThan ($targetMs * 0.5)
            $sampledMeanMs | Should -BeLessThan ($targetMs * 50)
        }
    }
}