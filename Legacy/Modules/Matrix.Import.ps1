
function Process-MatrixObjects {
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix,
        [Parameter(Mandatory)][object]$Html
    )

    #
    # Process each matrix item:
    #   - Generate its troubleshooting log
    #   - Attach TroubleshootingLogPath property
    #

    foreach (
        $matrixItem in
        $ImportedMatrix | Sort-Object { $_.File.Item.Name }
    ) {

        $logPath = $null

        try {
            $logPath = Write-MatrixTroubleshootingLog `
                -Matrix $matrixItem `
                -Html $Html
        }
        catch {
            Write-Warning "Failed to build troubleshooting log for '$($matrixItem.File.Item.Name)': $_"
        }

        #
        # Add or update TroubleshootingLogPath on the matrix item
        #
        $matrixItem |
        Add-Member -NotePropertyName TroubleshootingLogPath `
            -NotePropertyValue $logPath `
            -Force
    }

    return $ImportedMatrix
}

function New-HtmlCheckRow {
    param(
        [Parameter(Mandatory)]
        [object]$CheckObject
    )

    # Determine CSS class based on type (Error / Warning / Info)
    $cssClass = Get-HtmlClassProbTypeHC -Name $CheckObject.Type

    # HTML-encode dynamic fields
    $name = [System.Net.WebUtility]::HtmlEncode($CheckObject.Name)
    $desc = [System.Net.WebUtility]::HtmlEncode($CheckObject.Description)

    # Optional list of values
    $listHtml = Format-HtmlList -Value $CheckObject.Value

    # Output final row
    return @"
<tr>
    <td class="$cssClass"></td>
    <td colspan="7">
        <p class="probTitle">$name</p>
        <p>$desc</p>
        $listHtml
    </td>
</tr>
"@
}
function Get-HtmlClassProbTypeHC {
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('FatalError', 'Warning', 'Information')]
        [string]$Name
    )

    switch ($Name) {
        'FatalError' { return 'probTypeError' }
        'Warning' { return 'probTypeWarning' }
        'Information' { return 'probTypeInfo' }
    }
}

function Format-HtmlList {
    param([object]$Value)
    if (-not $Value) { return '' }
    if ($Value.Count -le 5 -and $Value -isnot [hashtable]) {
        $encodedItems = @($Value).ForEach(
            { "<li>$([System.Net.WebUtility]::HtmlEncode($_))</li>" }
        ) -join ''
        return "<ul>$encodedItems</ul>"
    }
    return '<p><i>Check JSON dump for multiple items.</i></p>'
}
