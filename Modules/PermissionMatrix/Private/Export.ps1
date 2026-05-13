
function Build-ExportDataHC {
    <#
        Build aggregated export data for permissions and form data.
        This is fed into actual export functions.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix
    )

    $permissionsRows = @()
    $formDataRows = @()

    foreach ($I in $ImportedMatrix) {

        # Permissions export rows
        foreach ($S in $I.Settings) {
            $permissionsRows += [pscustomobject]@{
                MatrixFile = $I.File.Item.Name
                Computer   = $S.Import.ComputerName
                Path       = $S.Import.Path
                Action     = $S.Import.Action
                Errors     = ($S.Check | Where-Object { $_.Type -eq 'FatalError' }).Count
                Warnings   = ($S.Check | Where-Object { $_.Type -eq 'Warning' }).Count
            }
        }

        # FormData sheet export rows
        if ($I.FormData.Import) {
            foreach ($fd in $I.FormData.Import) {
                $formDataRows += $fd
            }
        }
    }

    return [pscustomobject]@{
        Permissions = $permissionsRows
        FormData    = $formDataRows
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