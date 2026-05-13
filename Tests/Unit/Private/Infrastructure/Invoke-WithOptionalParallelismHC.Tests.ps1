#Requires -Version 7
#Requires -Modules Pester

Describe 'Invoke-WithOptionalParallelismHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Infrastructure\Invoke-WithOptionalParallelismHC.ps1"
    }

    # =========================================================================
    Context 'Sequential branch (ThrottleLimit <= 1)' {
        It 'invokes the scriptblock once per input item' {
            $items = @('a', 'b', 'c')

            $results = Invoke-WithOptionalParallelismHC `
                -InputObject $items `
                -ThrottleLimit 1 `
                -ScriptBlock { param($x) "processed-$x" }

            $results | Should -HaveCount 3
            $results | Should -Be @('processed-a', 'processed-b', 'processed-c')
        }

        It 'preserves input order in the results' {
            $items = 1..5

            $results = Invoke-WithOptionalParallelismHC `
                -InputObject $items `
                -ThrottleLimit 1 `
                -ScriptBlock { param($x) $x * 10 }

            $results | Should -Be @(10, 20, 30, 40, 50)
        }

        It 'passes ArgumentList items as additional scriptblock arguments' {
            $items = @('a', 'b')

            $results = Invoke-WithOptionalParallelismHC `
                -InputObject $items `
                -ThrottleLimit 1 `
                -ScriptBlock { param($item, $prefix, $suffix) "$prefix-$item-$suffix" } `
                -ArgumentList 'PRE', 'POST'

            $results | Should -Be @('PRE-a-POST', 'PRE-b-POST')
        }

        It 'handles empty input without invoking the scriptblock' {
            $script:invocationCount = 0

            $results = Invoke-WithOptionalParallelismHC `
                -InputObject @() `
                -ThrottleLimit 1 `
                -ScriptBlock { $script:invocationCount++ }

            $script:invocationCount | Should -Be 0
            $results | Should -BeNullOrEmpty
        }

        It 'runs sequentially when ThrottleLimit is <Throttle>' -ForEach @(
            @{ Throttle = 0 }
            @{ Throttle = 1 }
            @{ Throttle = -1 }
        ) {
            # We can verify sequential by checking that a $script: variable
            # incremented inside the scriptblock is visible to us afterward —
            # parallel runspaces wouldn't share it.
            $script:counter = 0

            Invoke-WithOptionalParallelismHC `
                -InputObject @(1, 2, 3) `
                -ThrottleLimit $Throttle `
                -ScriptBlock { $script:counter++ } | Out-Null

            $script:counter | Should -Be 3
        }

        It 'works when ArgumentList is not supplied (uses the @() default)' {
            # If the parameter default weren't @(), splatting $ArgumentList
            # would fail when the caller omits it.
            $results = Invoke-WithOptionalParallelismHC `
                -InputObject @('hello') `
                -ThrottleLimit 1 `
                -ScriptBlock { param($x) $x.ToUpper() }

            $results | Should -Be 'HELLO'
        }
    }

    # =========================================================================
    Context 'Parallel branch (ThrottleLimit > 1)' {
        # The function rehydrates the scriptblock via [scriptblock]::Create()
        # inside each parallel runspace. That rehydrated block loses its lexical
        # context, so $using: references in the caller's scriptblock do NOT
        # work — by the time the block runs, $using: has nothing to bind to.
        #
        # The caller's scriptblock can only communicate via:
        #   (a) values passed in via -ArgumentList, or
        #   (b) values returned from the scriptblock.
        #
        # The tests below use return values where possible, and -ArgumentList
        # to inject a ConcurrentBag where we need a side-channel.

        It 'invokes the scriptblock once per input item and returns results' {
            $items = @('a', 'b', 'c', 'd')

            $results = Invoke-WithOptionalParallelismHC `
                -InputObject $items `
                -ThrottleLimit 4 `
                -ScriptBlock { param($x) "processed-$x" }

            $results | Should -HaveCount 4
            ($results | Sort-Object) | Should -Be @(
                'processed-a', 'processed-b', 'processed-c', 'processed-d'
            )
        }

        It 'passes ArgumentList items as additional scriptblock arguments' {
            $results = Invoke-WithOptionalParallelismHC `
                -InputObject @('a', 'b') `
                -ThrottleLimit 2 `
                -ScriptBlock { param($item, $prefix, $suffix) "$prefix-$item-$suffix" } `
                -ArgumentList 'PRE', 'POST'

            ($results | Sort-Object) | Should -Be @('PRE-a-POST', 'PRE-b-POST')
        }

        It 'returns numeric results from the scriptblock' {
            $items = @(1, 2, 3, 4)

            $results = Invoke-WithOptionalParallelismHC `
                -InputObject $items `
                -ThrottleLimit 4 `
                -ScriptBlock { param($x) $x * 10 }

            $results | Should -HaveCount 4
            ($results | Sort-Object) | Should -Be @(10, 20, 30, 40)
        }

        It 'handles empty input without invoking the scriptblock' {
            # Use a ConcurrentBag passed via -ArgumentList as a side-channel
            # sentinel. If the scriptblock fires, the bag will have items.
            # Unary comma (,$bag) prevents [object[]]$ArgumentList from
            # enumerating the bag during parameter binding.
            $bag = [System.Collections.Concurrent.ConcurrentBag[string]]::new()

            $results = Invoke-WithOptionalParallelismHC `
                -InputObject @() `
                -ThrottleLimit 4 `
                -ArgumentList (, $bag) `
                -ScriptBlock {
                param($item, $sink)
                $sink.Add("called-$item")
            }

            $bag.Count | Should -Be 0
            $results | Should -BeNullOrEmpty
        }

        It 'runs in parallel when ThrottleLimit is <Throttle>' -ForEach @(
            @{ Throttle = 2 }
            @{ Throttle = 10 }
        ) {
            # Side-channel via ArgumentList: the parallel runspaces all write to
            # the same ConcurrentBag instance (reference-typed, thread-safe).
            # Unary comma (,$bag) prevents the [object[]]$ArgumentList binding
            # from enumerating the bag away — without it, $sink arrives null.
            $bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new()

            Invoke-WithOptionalParallelismHC `
                -InputObject @(1, 2, 3) `
                -ThrottleLimit $Throttle `
                -ArgumentList (, $bag) `
                -ScriptBlock {
                param($item, $sink)
                $sink.Add($item)
            } | Out-Null

            $bag.Count | Should -Be 3
            ($bag | Sort-Object) | Should -Be @(1, 2, 3)
        }
    }
}