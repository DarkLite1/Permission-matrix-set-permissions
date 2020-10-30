<# 
.SYNOPSIS
    Generate AD object names from a matrix and send the results to the user.

.DESCRIPTION
    Generate AD object names from a matrix. Each matrix needs to have a sheet 'Permissions' 
    and 'Settings'. Send a mail to the user with an Excel sheet in attachment containing
    the AD object names and the matrix file name.

.PARAMETER MailTo
    SMTP mail addresses

.PARAMETER Path
    Can be an Excel file or a folder containing multiple Excel files

.EXAMPLE
    &'Permission matrix\Get groups\Get groups.ps1' -ScriptName 'Matrix AD object names (BNL)' -Path 'Matrix.xlsx' -MailTo 'Brecht.Gijbels@heidelbergcement.com'

.NOTES
	CHANGELOG
    2018/06/14 Script born
    2020/10/05 Rewrote Pester tests to be compliant with Pester 5
    2020/10/21 Adjusted catch clause to simply throw

	AUTHOR Brecht.Gijbels@heidelbergcement.com #>

Param (
    [Parameter(Mandatory)]
    [String]$ScriptName = 'Matrix AD object names (BNL)',
    [Parameter(Mandatory)]
    [String]$Path,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [String]$LogFolder = "\\$env:COMPUTERNAME\Log",
    [String]$ScriptAdmin = 'Brecht.Gijbels@heidelbergcement.com'
)

Begin {
    Try {
        $null = Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        $LogParams = @{
            LogFolder    = New-FolderHC -Path $LogFolder -ChildPath "Permission matrix\Get groups\$ScriptName"
            Name         = $ScriptName
            Date         = 'ScriptStartTime'
            NoFormatting = $true
        }
        $LogFile = New-LogFileNameHC @LogParams
        #endregion

        if (-not (Test-Path $Path)) {
            throw "Path '$Path' not found."
        }

        [Array]$Result = Get-MatrixAdObjectNamesHC -Path $Path
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}

Process {
    Try {
        $MailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = 'Success'
            LogFolder = $LogParams.LogFolder
            Message   = 'AD object names successfully generated.'
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        $LogFile = $LogFile + '.xlsx'
        $LogFile | Remove-Item -EA Ignore

        $ExportParams = @{
            Path         = $LogFile
            AutoSize     = $true
            FreezeTopRow = $true
        }

        $i = 0
        foreach ($R in  $Result) {
            $i++
            $ExportParams.WorkSheetName = "$i $($R.FileName)"
            $ExportParams.TableName = "Table$i"

            $R.ADObject | Select-Object @{N = $R.FileName; E = { $_ } } | 
            Export-Excel @ExportParams

            $MailParams.Attachments = $LogFile
        }

        #region Format HTML
        $MailParams.Message += "<h3>Enabled matrix Excel files:</h3>"

        $MailParams.Message += $Result | Sort-Object FileName | 
        Select-Object FileName, @{N = 'ADObjects'; E = { $_.ADObject.Count } } | 
        ConvertTo-Html -Fragment -As Table

        $MailParams.Message += '<p><i>* Check the attachment for details</i></p>'
        #endregion

        $null = Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}