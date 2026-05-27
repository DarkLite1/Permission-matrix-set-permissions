function Import-MatrixFileHC {
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
            $Context.Config.Export.OverviewHtmlFile) {

            try {
                $formDataImport = Import-Excel `
                    -Path $MatrixFile.FullName `
                    -Sheet 'FormData' `
                    -DataOnly `
                    -ErrorAction Stop

                $fileResult.Sheets.FormData.Raw = $formDataImport

                $formDataCheck = Test-FormDataHC -FormData $formDataImport

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