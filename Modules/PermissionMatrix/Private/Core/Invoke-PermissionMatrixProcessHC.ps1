function Invoke-PermissionMatrixProcessHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Context,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    try {
        # ------------------------------------------------------------
        # Load matrix files
        # ------------------------------------------------------------
        $matrixFiles = Get-MatrixFilesHC `
            -MatrixConfig $Context.Matrix `
            -SystemErrors $SystemErrors

        if (-not $matrixFiles) {
            Add-ErrorHC `
                -Type 'Warning' `
                -Category 'Process' `
                -Message 'No matrix files found to process.' `
                -Name 'No matrix files' `
                -SystemErrors $SystemErrors

            return @()
        }

        # ------------------------------------------------------------
        # Process matrices
        # ------------------------------------------------------------
        $results = $matrixFiles | ForEach-Object -Parallel {
            param($file, $defaults, $context)

            Import-MatrixFileHC `
                -MatrixFile $file `
                -Defaults $defaults `
                -Context $context

        } -ThrottleLimit $throttle -ArgumentList $defaults, $Context

        $allMatrices = @()

        foreach ($fileResult in $results) {

            # Promote file-level checks to SystemErrors
            foreach ($check in $fileResult.File.Check) {
                Add-ErrorHC `
                    -Type $check.Type `
                    -Name $check.Name `
                    -Message $check.Description `
                    -Category 'Matrix' `
                    -SystemErrors $SystemErrors
            }

            # Collect matrices
            $allMatrices += $fileResult.Matrices
        }

        return $allMatrices
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Category 'Runtime' `
            -Name 'PROCESS stage failure' `
            -Message "Unhandled exception occurred: $_" `
            -SystemErrors $SystemErrors

        return @()
    }
}