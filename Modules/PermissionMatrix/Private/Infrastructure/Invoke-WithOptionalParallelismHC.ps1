function Invoke-WithOptionalParallelismHC {
    <#
    .SYNOPSIS
        Runs a scriptblock against each input item, optionally in parallel.

    .NOTES
        ArgumentList follows the standard PowerShell convention (cf.
        Start-Job, Invoke-Command): supplied values are passed as positional
        arguments after the input item.

        When passing an enumerable object (e.g. a ConcurrentBag, list, hash-
        table), wrap it with the unary comma to prevent the binder from
        enumerating it:

            -ArgumentList (,$bag)        # passes the bag as a single arg
            -ArgumentList @($bag)        # WRONG: bag gets enumerated

        Also note: the scriptblock is rehydrated via [scriptblock]::Create()
        inside each parallel runspace, so $using: from the caller's scope
        does NOT work. Pass all needed data via -ArgumentList.
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