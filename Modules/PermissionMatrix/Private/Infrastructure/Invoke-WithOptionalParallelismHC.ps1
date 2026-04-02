function Invoke-WithOptionalParallelismHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [IEnumerable]$InputObject,

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

        $InputObject | ForEach-Object -Parallel $ScriptBlock `
            -ThrottleLimit $ThrottleLimit `
            -ArgumentList $ArgumentList
    }
}