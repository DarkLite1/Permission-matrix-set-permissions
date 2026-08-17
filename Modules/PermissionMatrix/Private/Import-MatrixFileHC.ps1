function New-MatrixFileResultHC {
    <#
    .SYNOPSIS
        Creates the empty result object that carries one matrix file through
        the pipeline.

    .DESCRIPTION
        Import-MatrixFileHC builds this shape for every matrix file it reads.
        Invoke-PermissionMatrixBeginHC needs the identical shape for the
        fallback it creates when a runspace throws before the import returned
        anything, so both call this function and the two cannot drift apart.

        Keeping the shapes in sync matters more than it looks: the reporting
        stage reads .Item and .ReportFileName and assigns to .LogFolder and
        .ReportFilePath. Assigning a property that does not exist on a
        [PSCustomObject] throws, and the whole per-file log loop in
        Invoke-PermissionMatrixEndHC sits inside a single try block, so one
        incomplete result aborts log writing for every remaining file in the
        run rather than just its own.

    .PARAMETER MatrixFile
        The matrix file this result describes.

    .EXAMPLE
        $fileResult = New-MatrixFileResultHC -MatrixFile $file

        Creates an empty result object for $file, with all check collections
        initialized and ready to receive entries.
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$MatrixFile
    )

    [pscustomobject]@{
        Item           = $MatrixFile
        ExcelInfo      = $null
        Check          = [System.Collections.Generic.List[pscustomobject]]::new()
        Sheets         = @{
            Permissions = @{
                Raw       = $null
                Formatted = $null
                Check     = [System.Collections.Generic.List[pscustomobject]]::new()
            }
            Settings    = @{
                Raw       = $null
                Formatted = $null
            }
            FormData    = @{
                Raw       = $null
                Formatted = $null
                Check     = [System.Collections.Generic.List[pscustomobject]]::new()
            }
        }
        Matrices       = [System.Collections.Generic.List[pscustomobject]]::new()
        LogFolder      = $null
        ReportFileName = '00 - Execution Report.html'
        ReportFilePath = $null
    }
}

function Import-MatrixFileHC {
    <#
    .SYNOPSIS
        Safely imports and structures data from a Permission Matrix Excel file.

    .DESCRIPTION
        Reads the 'Settings', 'Permissions', and optional 'FormData' worksheets 
        from a provided Excel matrix file. 
        
        The raw Excel data is converted into normalized, formatted PowerShell 
        objects. For every 'Enabled' row found in the 'Settings' tab, the 
        script generates a distinct job execution object (Matrix) complete with 
        a unique GUID. 
        
        Architectural Note: This function avoids throwing terminating errors. 
        If a file is corrupt or missing mandatory worksheets, it safely catches 
        the exception and appends a 'FatalError' to the returned object's .
        Check property, allowing the main orchestrator to gracefully skip it 
        while continuing to process other valid files.

        Locking Note: The workbook is first copied to a temporary file and all
        worksheets are read from that copy. This keeps the user's original
        matrix file unlocked so it can still be edited while the script runs.
        The temporary copy is always removed afterwards.

    .PARAMETER MatrixFile
        A [System.IO.FileInfo] object pointing to the specific Excel (.xlsx) 
        file to be processed.

    .PARAMETER Context
        The global pipeline context object. Used to check the runtime 
        configuration (e.g., determining if the 'FormData' sheet needs to be 
        extracted based on ServiceNow/HTML export settings).

    .OUTPUTS
        System.Management.Automation.PSCustomObject.
        Returns a comprehensive $fileResult object containing:
        - Item        : The original FileInfo object.
        - Sheets      : The Raw and Formatted data extracted from the   
                        worksheets.
        - Matrices    : A List of initialized execution jobs (one for each 
                        enabled Setting).
        - Check       : A Generic List containing any structural file errors
                        (e.g., Missing worksheets).

    .EXAMPLE
        $fileInfo = Get-Item -LiteralPath 'C:\MatrixFiles\Finance_Matrix.xlsx'
        $globalContext = [pscustomobject]@{ Config = $jsonConfig }
        
        $result = Import-MatrixFileHC `
            -MatrixFile $fileInfo `
            -Context $globalContext
        
        if ($result.Check.Type -contains 'FatalError') {
            Write-Warning "File was structurally invalid!"
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$MatrixFile,

        [Parameter(Mandatory)]
        [pscustomobject]$Context
    )

    $fileResult = New-MatrixFileResultHC -MatrixFile $MatrixFile

    $tempMatrixFile = $null

    try {
        #region Work on a temporary copy so the original file stays unlocked
        # Get-ExcelWorkbookInfo and Import-Excel keep the workbook open while
        # reading it. Reading the user's original file directly can leave it
        # locked long enough that someone editing it at the same moment gets a
        # 'file in use' error. Copying to a temporary file first (a fast,
        # read-shared copy) means every read below targets the throwaway copy,
        # so the original matrix file is never held open and users can keep
        # editing it. The copy is removed in the finally block.
        $tempMatrixFile = Join-Path `
            -Path ([System.IO.Path]::GetTempPath()) `
            -ChildPath "PermissionMatrix_$([guid]::NewGuid().ToString('N'))_$($MatrixFile.Name)"

        Copy-Item `
            -LiteralPath $MatrixFile.FullName `
            -Destination $tempMatrixFile `
            -Force `
            -ErrorAction Stop
        #endregion

        #region Get Excel workbook info
        $fileResult.ExcelInfo = Get-ExcelWorkbookInfo `
            -Path $tempMatrixFile `
            -ErrorAction Stop
        #endregion

        #region Import Settings sheet
        $settingsSheet = @(
            Import-Excel `
                -Path $tempMatrixFile `
                -Sheet 'Settings' `
                -DataOnly `
                -ErrorAction Stop
        )
        $fileResult.Sheets.Settings.Raw = $settingsSheet

        <# Collect the indexes of the enabled rows rather than the rows
        themselves. Format-SettingStringsHC emits one row per input row in
        order, so the formatted sheet built after the guard below is index
        aligned with the raw sheet, and each matrix can then reference both
        its raw and its formatted row without formatting that row a second
        time.

        Status is trimmed for the comparison because formatting has not run
        yet at this point. Selecting on the raw value meant a cell containing
        'Enabled ' with a trailing space silently disabled the whole matrix
        file with no indication why. The "" wrapper keeps this safe when
        Status is absent or not a string, where .Trim() would throw.

        A 'for' loop is used because 0..($settingsSheet.Count - 1) counts
        backwards on an empty sheet. #>
        $enabledSettingIndexes = @(
            for ($i = 0; $i -lt $settingsSheet.Count; $i++) {
                if ("$($settingsSheet[$i].Status)".Trim() -eq 'Enabled') { $i }
            }
        )

        <# .Count, not '-not': a single-element array is unwrapped in a boolean
        context, so '-not @(0)' is $true and a file whose only enabled row is
        the first one would be reported as having no enabled settings. #>
        if ($enabledSettingIndexes.Count -eq 0) {
            $fileResult.Check.Add(
                [pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'No enabled matrix settings'
                    Description = 'This matrix file does not contain any enabled matrix settings row and is skipped.'
                    Value       = "No Settings row with 'Status = Enabled'"
                }
            )

            return
        }

        $fileResult.Sheets.Settings.Formatted = @(
            $fileResult.Sheets.Settings.Raw | Format-SettingStringsHC
        )
        #endregion

        #region Import Permissions sheet
        $permissionsSheet = Import-Excel `
            -Path $tempMatrixFile `
            -Sheet 'Permissions' `
            -NoHeader `
            -DataOnly `
            -ErrorAction Stop

        $fileResult.Sheets.Permissions.Raw = $permissionsSheet 

        $fileResult.Sheets.Permissions.Formatted = $fileResult.Sheets.Permissions.Raw | Format-PermissionsStringsHC
        #endregion

        #region Import optional FormData
        if ($Context.Config.Export.ServiceNowFormDataExcelFile -or
            $Context.Config.Export.OverviewHtmlFile -or
            $Context.Config.Export.PermissionsExcelFile -or
            $Context.Config.AuditReport) {

            try {
                $formDataImport = Import-Excel `
                    -Path $tempMatrixFile `
                    -Sheet 'FormData' `
                    -DataOnly `
                    -ErrorAction Stop

                $fileResult.Sheets.FormData.Raw = $formDataImport

                $formDataCheck = Test-MatrixFormDataHC -FormData $formDataImport

                if ($formDataCheck) {
                    $fileResult.Sheets.FormData.Check.Add($formDataCheck)
                }
                else {
                    $fileResult.Sheets.FormData.Formatted = $formDataImport[0] | Format-FormDataStringsHC
                }
            }
            catch {
                $fileResult.Sheets.FormData.Check.Add(
                    [pscustomobject]@{
                        Type        = 'FatalError'
                        Name        = "Worksheet 'FormData' not found"
                        Description = "Worksheet 'FormData' is required when ServiceNow export is enabled."
                        Value       = $_
                    }
                )
            }
        }
        #endregion

        #region Create matrix per enabled Settings row
        foreach ($index in $enabledSettingIndexes) {
            $matrix = [pscustomobject]@{
                ID          = [guid]::NewGuid().ToString()
                Setting     = @{
                    Raw       = $settingsSheet[$index]
                    Formatted = $fileResult.Sheets.Settings.Formatted[$index]
                }
                Check       = [System.Collections.Generic.List[pscustomobject]]::new()
                Matrix      = [System.Collections.Generic.List[pscustomobject]]::new()
                AdObjects   = @{}
                JobTime     = @{}
                # Volume and cost counters returned by SetPermissions.ps1.
                # Declared here (rather than added later) so every consumer can
                # test it without a PSObject.Properties.Match guard. Stays
                # $null for rows that never executed — a skipped row has no
                # telemetry, which is different from a row that walked nothing.
                Telemetry   = $null
                DiagnosticsFileName = $null
                FileContext = $fileResult
            }

            $fileResult.Matrices.Add($matrix)
        }
        #endregion
    }
    catch {
        $fileResult.Check.Add(
            [pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = $_
            }
        )
    }
    finally {
        #region Remove the temporary copy of the matrix file
        if ($tempMatrixFile -and (Test-Path -LiteralPath $tempMatrixFile)) {
            Remove-Item `
                -LiteralPath $tempMatrixFile `
                -Force `
                -ErrorAction SilentlyContinue
        }
        #endregion

        $fileResult
    }
}