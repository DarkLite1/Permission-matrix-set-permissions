<#
    Export.ps1
    Consolidated export logic for Toolbox.PermissionMatrixHC.

    Contains:
        - Build-ExportDataHC
        - Export-FilesHC
        - Export-PermissionsFileHC
        - Export-ServiceNowFormDataHC
        - Export-OverviewHtmlHC
#>

#region Build-ExportDataHC
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
#endregion



#region Export-FilesHC
function Export-FilesHC {
    <#
        Executes all export operations based on settings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix,
        [Parameter(Mandatory)]      $ExportSettings,
        [Parameter(Mandatory)]      $HtmlOverview,
        [Parameter(Mandatory)]      $Counters
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

    # 3. Overview HTML
    if ($ExportSettings.OverviewHtmlFile) {
        $results.OverviewHtml = Export-OverviewHtmlHC `
            -Html $HtmlOverview `
            -Path $ExportSettings.OverviewHtmlFile
    }

    return $results
}
#endregion



#region Export-PermissionsFileHC
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
#endregion



#region Export-ServiceNowFormDataHC
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
#endregion



#region Export-OverviewHtmlHC
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
#endregion