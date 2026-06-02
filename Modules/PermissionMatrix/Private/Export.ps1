
function Build-ExportDataHC {
    <#
    .SYNOPSIS
        Builds aggregated export data for permissions and ServiceNow form data.

    .DESCRIPTION
        Iterates through the processed matrices and extracts the execution 
        results (Errors, Warnings, Paths, Actions) and form data into flat, 
        structured lists. 
        This output is specifically formatted to be fed directly into the HTML 
        and Excel reporting functions.

    .PARAMETER ImportedMatrix
        An array of processed matrix file objects (typically from $Context.
        FileResults) containing the settings, execution checks, and associated 
        form data.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$ImportedMatrix
    )

    $permissionsRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $formDataRows = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($I in $ImportedMatrix) {

        # Permissions export rows
        if ($I.Settings) {
            foreach ($S in $I.Settings) {
                $permissionsRows.Add([pscustomobject]@{
                        MatrixFile = $I.File.Item.Name
                        Computer   = $S.Import.ComputerName
                        Path       = $S.Import.Path
                        Action     = $S.Import.Action
                        Errors     = @(
                            $S.Check | 
                            Where-Object { $_.Type -eq 'FatalError' }).Count
                        Warnings   = @(
                            $S.Check | 
                            Where-Object { $_.Type -eq 'Warning' }).Count
                    })
            }
        }

        # FormData sheet export rows
        if ($I.FormData.Import) {
            $formDataRows.AddRange([pscustomobject[]]@($I.FormData.Import))
        }
    }

    return [pscustomobject]@{
        Permissions = $permissionsRows.ToArray()
        FormData    = $formDataRows.ToArray()
    }
}

function Export-FilesHC {
    <#
    .SYNOPSIS
        Executes all export operations based on settings.

    .DESCRIPTION
        Builds export data from imported matrices, then writes the configured
        export artifacts to disk: Permissions Excel, ServiceNow FormData Excel,
        and the standalone overview HTML page.

        The overview HTML is generated internally — callers no longer pass it
        in. The email summary body is a separate artifact built by EndHC and
        is not used here.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix,
        [Parameter(Mandatory)]      $ExportSettings
    )

    $exportData = Build-ExportDataHC -ImportedMatrix $ImportedMatrix

    $results = [ordered]@{
        Permissions  = $null
        FormData     = $null
        OverviewHtml = $null
    }

    # 1. Permissions Excel
    if ($ExportSettings.PermissionsExcelFile) {
        $results.Permissions = Export-PermissionsFileHC `
            -Rows $exportData.Permissions `
            -Path $ExportSettings.PermissionsExcelFile
    }

    # 2. ServiceNow FormData Excel
    if ($ExportSettings.ServiceNowFormDataExcelFile) {
        $results.FormData = Export-ServiceNowFormDataHC `
            -Rows $exportData.FormData `
            -Path $ExportSettings.ServiceNowFormDataExcelFile
    }

    # 3. Overview HTML (built from FormData rows; independent of the email body)
    if ($ExportSettings.OverviewHtmlFile) {
        $html = New-OverviewHtmlHC -FormData $exportData.FormData
        $results.OverviewHtml = Export-OverviewHtmlHC `
            -Html $html `
            -Path $ExportSettings.OverviewHtmlFile
    }

    return $results
}

function Export-PermissionsFileHC {
    <#
        Writes a permissions Excel export using ImportExcel module.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][string]$Path
    )

    try {
        $Rows | Export-Excel -Path $Path -WorksheetName 'Permissions' -AutoSize
        return $Path
    }
    catch {
        throw "Failed exporting Permissions Excel file: $_"
    }
}

function Export-ServiceNowFormDataHC {
    <#
        Writes ServiceNow FormData into an Excel file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][string]$Path
    )

    try {
        $Rows | Export-Excel -Path $Path -WorksheetName 'FormData' -AutoSize
        return $Path
    }
    catch {
        throw "Failed exporting ServiceNow FormData Excel: $_"
    }
}

function Export-OverviewHtmlHC {
    <#
        Writes the generated HTML overview page to a file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Html,
        [Parameter(Mandatory)][string]$Path
    )

    try {
        $Html | Out-File -LiteralPath $Path -Encoding utf8 -Force
        return $Path
    }
    catch {
        throw "Failed exporting Overview HTML file: $_"
    }
}