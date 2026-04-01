function Import-MatrixFileHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$MatrixFile,

        [Parameter(Mandatory)]
        [pscustomobject]$Defaults,

        [Parameter(Mandatory)]
        [pscustomobject]$Context
    )

    $fileResult = [pscustomobject]@{
        File     = @{
            Name  = $MatrixFile.Name
            Path  = $MatrixFile.FullName
            Check = [System.Collections.Generic.List[pscustomobject]]::new()
        }
        Matrices = [System.Collections.Generic.List[pscustomobject]]::new()
    }

    try {
        # ------------------------------------------------------------
        # Import Settings sheet
        # ------------------------------------------------------------
        $settingsSheet = @(
            Import-Excel `
                -Path $MatrixFile.FullName `
                -Sheet 'Settings' `
                -DataOnly `
                -ErrorAction Stop
        ).Where({ $_.Status -eq 'Enabled' })

        if (-not $settingsSheet) {
            $fileResult.File.Check.Add([pscustomobject]@{
                    Type        = 'Warning'
                    Name        = 'Matrix disabled'
                    Description = 'Every Excel file needs at least one enabled matrix.'
                    Value       = 'No rows with Status = Enabled'
                })

            return $fileResult
        }

        # ------------------------------------------------------------
        # Import Permissions sheet ONCE
        # ------------------------------------------------------------
        $permissions = Import-Excel `
            -Path $MatrixFile.FullName `
            -Sheet 'Permissions' `
            -NoHeader `
            -DataOnly `
            -ErrorAction Stop |
        Format-PermissionsStringsHC

        # ------------------------------------------------------------
        # Optional FormData
        # ------------------------------------------------------------
        $formData = $null
        if ($Context.Export.ServiceNowFormDataExcelFile -or
            $Context.Export.OverviewHtmlFile) {

            try {
                $formDataImport = Import-Excel `
                    -Path $MatrixFile.FullName `
                    -Sheet 'FormData' `
                    -DataOnly `
                    -ErrorAction Stop

                $formDataCheck = Test-FormDataHC $formDataImport
                if ($formDataCheck) {
                    $fileResult.File.Check.Add($formDataCheck)
                }
                else {
                    $formData = $formDataImport[0]
                }
            }
            catch {
                $fileResult.File.Check.Add([pscustomobject]@{
                        Type        = 'FatalError'
                        Name        = "Worksheet 'FormData' not found"
                        Description = "Worksheet 'FormData' is required when ServiceNow export is enabled."
                        Value       = $_
                    })
            }
        }

        # ------------------------------------------------------------
        # Create ONE matrix per enabled Settings row
        # ------------------------------------------------------------
        foreach ($S in $settingsSheet) {
            $matrix = [pscustomobject]@{
                ID          = $null
                Import      = Format-SettingStringsHC -Settings $S
                Check       = [System.Collections.Generic.List[pscustomobject]]::new()
                Matrix      = [System.Collections.Generic.List[pscustomobject]]::new()
                AdObjects   = @{}
                JobTime     = @{}
                Defaults    = $Defaults
                Permissions = $permissions
                FormData    = $formData
            }

            # Optional: validate settings row here
            # Add to $matrix.Check if needed

            $fileResult.Matrices.Add($matrix)
        }
    }
    catch {
        $fileResult.File.Check.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = $_
            })
    }

    return $fileResult
}