function Format-SettingStringsHC {
    <#
    .SYNOPSIS
        String manipulations on values in the 'Settings' sheet.

    .DESCRIPTION
        Remove leading and trailing spaces from strings. Add the domain name to
        the ComputerName property when it's not there. Remove trailing slashes
        from the Path. ...

        Spaces are converted to NULL values.

    .PARAMETER Settings
        One row in the Excel sheet 'Settings'.
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject]$Settings
    )

    process {
        $Obj = [ordered]@{}

        foreach ($Prop in $Settings.PSObject.Properties) {
            if ([string]::IsNullOrWhiteSpace($Prop.Value)) {
                $Obj[$Prop.Name] = $null
                continue
            }

            $Value = $Prop.Value.ToString().Trim()

            $Obj[$Prop.Name] = switch ($Prop.Name) {
                { $_ -in 'Action', 'Status' } {
                    $Value.Substring(0, 1).ToUpper() + $Value.Substring(1).ToLower()
                }
                'ComputerName' {
                    $Value = $Value.ToUpper()
                    $Domain = $env:USERDNSDOMAIN
                    
                    if ($Domain -and $Value -like "*.$Domain") {
                        $Value = $Value -ireplace [regex]::Escape(".$Domain"), ''
                    }
                    $Value
                }
                'Path' {
                    $Value.TrimEnd('\')
                }
                default {
                    $Value
                }
            }
        }

        [PSCustomObject]$Obj
    }
}
function Test-MatrixSettingHC {
    <#
    .SYNOPSIS
        Verify input for the Excel sheet 'Settings'.

    .DESCRIPTION
        Verify if one Excel row in the Excel sheet 'Settings' is correct. A FatalError object is created
        for each incorrect setting found (missing ComputerName parameter, ...)

        All rows are tested. In case only 'Status -eq Enabled' rows need to be tested,
        a filter needs to be applied upfront.

    .PARAMETER Setting
        Represents one row in the Excel sheet 'Settings', as retrieved by Import-Excel.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject]$Setting
    )

    process {
        try {
            $Properties = ($Setting | Get-Member -MemberType NoteProperty).Name

            # Define our mandatory properties once
            $MandatoryProps = @('ComputerName', 'Path', 'Action')

            #region Missing property
            $MissingProperty = $MandatoryProps.Where({ $_ -notin $Properties })
            if ($MissingProperty) {
                [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Missing column header'
                    Description = "The column headers 'ComputerName', 'Path' and 'Action' are mandatory."
                    Value       = $MissingProperty -join ', '
                }
            }
            #endregion

            #region Blank property value
            $BlankProperty = $MandatoryProps.Where({ 
                    $_ -in $Properties -and [string]::IsNullOrWhiteSpace($Setting.$_)
                })

            if ($BlankProperty) {
                [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Missing value'
                    Description = "Values for 'ComputerName', 'Path' and 'Action' are mandatory."
                    Value       = $BlankProperty -join ', '
                }
            }
            #endregion

            #region Action can only be New, Fix or Check
            if (-not [string]::IsNullOrWhiteSpace($Setting.Action) -and $Setting.Action -notmatch '^(New|Fix|Check)$') {
                [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Action value incorrect'
                    Description = "Only the values 'New', 'Fix' or 'Check' are supported in the field 'Action'."
                    Value       = $Setting.Action
                }
            }
            #endregion

            #region Path needs to be valid local path
            if (
                (-not [string]::IsNullOrWhiteSpace($Setting.Path)) -and 
                ($Setting.Path -notmatch '^(?!.*\\\s)(?!.*\s\\)(?!.*\s$)[a-zA-Z]:\\[^<>:"/|?*]+$')
            ) {
                [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Path value incorrect'
                    Description = "The 'Path' needs to be defined as a local folder (Ex. 'E:\Department\Finance')."
                    Value       = $Setting.Path
                }
            }
            #endregion

            #region JobsAtOnce is not a number between 0-9
            if (-not [string]::IsNullOrWhiteSpace($Setting.JobsAtOnce)) {
                if ($Setting.JobsAtOnce -notmatch '^[0-9]$') {
                    [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'JobsAtOnce is not a valid number'
                        Description = "The value for 'JobsAtOnce' needs to be a number between 0 and 9."
                        Value       = $Setting.JobsAtOnce
                    }
                }
            }
            #endregion
        }
        catch {
            throw "Failed testing the Excel sheet 'Settings' row for incorrect data: $_"
        }
    }
}