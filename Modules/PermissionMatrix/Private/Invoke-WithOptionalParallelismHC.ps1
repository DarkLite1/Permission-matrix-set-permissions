function Invoke-WithOptionalParallelismHC {
    <#
    .SYNOPSIS
        Executes a scriptblock against an array of input objects, seamlessly 
        switching between sequential and parallel execution.

    .DESCRIPTION
        A dynamic execution wrapper that allows the pipeline to scale 
        concurrency based on the provided configuration. 
        
        If the `ThrottleLimit` is greater than 1, the function utilizes 
        PowerShell 7's `ForEach-Object -Parallel` to spin up concurrent 
        runspaces. If the limit is 1 or less, it gracefully falls back to a 
        standard sequential `foreach` loop running on the main thread, making 
        it highly versatile for debugging or resource-constrained environments.

    .PARAMETER InputObject
        An array of items to process. Inside the scriptblock, the current item 
        is passed as the first positional parameter (and is also accessible via 
        `$_`).

    .PARAMETER ScriptBlock
        The code to execute against each item. Must define a `param()` block to 
        receive the input item and any additional arguments.

    .PARAMETER ThrottleLimit
        The maximum number of concurrent threads/runspaces. A value of 1 or 0 
        forces standard sequential execution.

    .PARAMETER ArgumentList
        An optional array of extra arguments to pass positionally to the 
        scriptblock after the main input item.

    .EXAMPLE
        $items = @('ServerA', 'ServerB')
        $myList = [System.Collections.Generic.List[string]]::new()
        
        Invoke-WithOptionalParallelismHC `
            -InputObject $items `
            -ThrottleLimit 5 `
            -ArgumentList (,$myList), $true `
            -ScriptBlock {
                param($ComputerName, $ListRef, $LogEnabled)
            
                if ($LogEnabled) {
                    $ListRef.Add("Processed $ComputerName")
                }
            }

    .NOTES
        ArgumentList follows the standard PowerShell convention (cf. Start-Job, 
        Invoke-Command): supplied values are passed as positional arguments 
        after the input item.

        CRITICAL: When passing an enumerable object (e.g., a ConcurrentBag, 
        Generic List, hashtable), wrap it with the unary comma to prevent the 
        binder from unrolling/enumerating it:
            -ArgumentList (,$bag) # CORRECT: passes the bag as a single argument
            -ArgumentList @($bag) # WRONG: bag gets enumerated and split

        SCOPE LIMITATION: The scriptblock is dynamically rehydrated via 
        [scriptblock]::Create() inside each parallel runspace. Because of this, 
        the `$using:` scope modifier from the caller's scope does NOT work 
        natively inside the scriptblock. You must pass all external variables 
        explicitly via the -ArgumentList parameter.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$InputObject,

        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory)]
        [int]$ThrottleLimit,

        [Parameter()]
        [object[]]$ArgumentList = @()
    )

    if ($ThrottleLimit -le 1) {
        Write-Verbose 'Running sequentially (ThrottleLimit <= 1)'

        foreach ($item in $InputObject) {
            & $ScriptBlock $item @ArgumentList
        }
    }
    else {
        Write-Verbose "Running in parallel (ThrottleLimit = $ThrottleLimit)"

        $scriptBlockString = $ScriptBlock.ToString()

        $InputObject | ForEach-Object -Parallel {
            $rehydratedBlock = [scriptblock]::Create($using:scriptBlockString)
            $splatArgs = $using:ArgumentList
            & $rehydratedBlock $_ @splatArgs
        } -ThrottleLimit $ThrottleLimit
    }
}