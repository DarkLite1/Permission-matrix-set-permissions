function Out-LogFileHC {
    <#
        Generic exporter for CSV / JSON / TXT / XLSX log files.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [PSCustomObject[]]$DataToExport,
        [Parameter(Mandatory)] [String]$PartialPath,
        [Parameter(Mandatory)] [String[]]$FileExtensions,
        [hashtable]$ExcelFile = @{
            SheetName = 'Overview'
            TableName = 'Overview'
            CellStyle = $null
        },
        [Switch]$Append
    )

    $allPaths = @()

    foreach ($ext in ($FileExtensions | Sort-Object -Unique)) {

        $logFilePath = "$PartialPath$ext"

        try {
            switch ($ext) {

                '.csv' {
                    $DataToExport |
                    Export-Csv -LiteralPath $logFilePath -Delimiter ';' `
                        -Append:$Append -NoTypeInformation
                    break
                }

                '.json' {
                    $converted = foreach ($item in $DataToExport) {
                        foreach ($p in $item.PSObject.Properties) {
                            if ($p.Value -is [System.Management.Automation.ErrorRecord]) {
                                $item.$($p.Name) = $p.Value.Exception.Message
                            }
                        }
                        $item
                    }

                    if ($Append -and (Test-Path $logFilePath)) {
                        $existing = Get-Content -LiteralPath $logFilePath -Raw | ConvertFrom-Json
                        $converted = @($converted) + @($existing)
                    }

                    $converted |
                    ConvertTo-Json -Depth 7 |
                    Out-File -LiteralPath $logFilePath -Encoding utf8 -Force
                    break
                }

                '.txt' {
                    $DataToExport |
                    Format-List * |
                    Out-File -LiteralPath $logFilePath -Append:$Append
                    break
                }

                '.xlsx' {
                    if (-not $Append -and (Test-Path $logFilePath)) {
                        Remove-Item -LiteralPath $logFilePath -Force
                    }

                    $params = @{
                        Path          = $logFilePath
                        Append        = $true
                        AutoNameRange = $true
                        AutoSize      = $true
                        FreezeTopRow  = $true
                        WorksheetName = $ExcelFile.SheetName
                        TableName     = $ExcelFile.TableName
                    }

                    if ($ExcelFile.CellStyle) {
                        $params.CellStyleSB = $ExcelFile.CellStyle
                    }

                    $DataToExport | Export-Excel @params
                    break
                }

                default {
                    throw "Unsupported file extension '$ext'."
                }
            }

            $allPaths += $logFilePath
        }
        catch {
            Write-Warning "Failed to export log '$logFilePath': $_"
        }
    }

    return $allPaths
}