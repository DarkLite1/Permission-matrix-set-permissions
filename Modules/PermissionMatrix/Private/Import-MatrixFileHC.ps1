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

    $fileResult = [pscustomobject]@{
        Item           = $MatrixFile
        ExcelInfo      = $null
        Check          = [System.Collections.Generic.List[pscustomobject]]::new()
        Sheets         = @{
            Permissions = @{
                Raw       = $null
                Formatted = $null
            }
            Settings    = @{
                Raw       = $null
                Formatted = $null
            }
            FormData    = @{
                Raw       = $null
                Formatted = $null
            }
        }
        Matrices       = [System.Collections.Generic.List[pscustomobject]]::new()
        LogFolder      = $null
        ReportFileName = '00 - Execution Report.html'
        ReportFilePath = $null
    }

    try {
        #region Get Excel workbook info
        $fileResult.ExcelInfo = Get-ExcelWorkbookInfo `
            -Path $matrixFile.FullName `
            -ErrorAction Stop
        #endregion

        #region Import Settings sheet
        $settingsSheet = @(
            Import-Excel `
                -Path $MatrixFile.FullName `
                -Sheet 'Settings' `
                -DataOnly `
                -ErrorAction Stop
        )
        $fileResult.Sheets.Settings.Raw = $settingsSheet

        $enabledSettings = $settingsSheet.Where({ $_.Status -eq 'Enabled' })

        if (-not $enabledSettings) {
            $fileResult.Check.Add(
                [pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'No enabled matrix settings'
                    Description = 'This matrix file does not contain any enabled matrix settings row and is skipped.'
                    Value       = "No Settings row with `Status = Enabled'"
                }
            )

            return
        }

        $fileResult.Sheets.Settings.Formatted = $fileResult.Sheets.Settings.Raw | Format-SettingStringsHC
        #endregion

        #region Import Permissions sheet
        $permissionsSheet = Import-Excel `
            -Path $MatrixFile.FullName `
            -Sheet 'Permissions' `
            -NoHeader `
            -DataOnly `
            -ErrorAction Stop

        $fileResult.Sheets.Permissions.Raw = $permissionsSheet 

        $fileResult.Sheets.Permissions.Formatted = $fileResult.Sheets.Permissions.Raw | Format-PermissionsStringsHC
        #endregion

        #region Import optional FormData
        $formData = $null
        if ($Context.Config.Export.ServiceNowFormDataExcelFile -or
            $Context.Config.Export.OverviewHtmlFile -or
            $Context.Config.Export.PermissionsExcelFile) {

            try {
                $formDataImport = Import-Excel `
                    -Path $MatrixFile.FullName `
                    -Sheet 'FormData' `
                    -DataOnly `
                    -ErrorAction Stop

                $fileResult.Sheets.FormData.Raw = $formDataImport

                $formDataCheck = Test-MatrixFormDataHC -FormData $formDataImport

                if ($formDataCheck) {
                    $fileResult.Check.Add($formDataCheck)
                }
                else {
                    $fileResult.Sheets.FormData.Formatted = $formDataImport[0] | Format-FormDataStringsHC
                }
            }
            catch {
                $fileResult.Check.Add(
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
        foreach ($enabledSetting in $enabledSettings) {
            $matrix = [pscustomobject]@{
                ID          = [guid]::NewGuid().ToString()
                Setting     = @{
                    Raw       = $enabledSetting
                    Formatted = Format-SettingStringsHC `
                        -Settings $enabledSetting
                }
                Check       = [System.Collections.Generic.List[pscustomobject]]::new()
                Matrix      = [System.Collections.Generic.List[pscustomobject]]::new()
                AdObjects   = @{}
                JobTime     = @{}
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
        $fileResult
    }
}