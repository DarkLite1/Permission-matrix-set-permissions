function Invoke-WithOptionalParallelismHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
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
            & $rehydratedBlock $_ @($using:ArgumentList)
        } -ThrottleLimit $ThrottleLimit
    }
}