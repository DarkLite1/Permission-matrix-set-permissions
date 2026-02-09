#Requires -Version 7
#Requires -Modules ImportExcel, ServiceNow

<# 
    .SYNOPSIS
        Upload data to ServiceNow

    .DESCRIPTION
        Send data to the ServiceNow API that will be used to create the ServiceNow form. This form is used by users to select the correct
        security group to gain access to the correct files and folders defined
        in a matrix Excel file.
        
        This script will first delete all records in the ServiceNow table and
        then upload all new records to to table.

    .PARAMETER CredentialsFilePath
        The .JSON file containing the che logon credentials for ServiceNow.

    .PARAMETER Environment
        The ServiceNow environment where the data will be send to. This 
        environment should be available in the CredentialsFilePath.

    .PARAMETER FormDataExcelFilePath
        The Excel file containing the records to be uploaded.

    .PARAMETER ExcelFileWorksheetName
        The name of the worksheet used in the FormDataExcelFilePath.

    .PARAMETER TableName
        The name of the table in ServiceNow where to store the records.

    .PARAMETER MaxRetries
        How many times to try to upload the data on connection errors.
#>

param (
    [Parameter(Mandatory)]
    [String]$CredentialsFilePath,
    [Parameter(Mandatory)]
    [String]$Environment,
    [Parameter(Mandatory)]
    [String]$FormDataExcelFilePath,
    [Parameter(Mandatory)]
    [String]$ExcelFileWorksheetName,
    [Parameter(Mandatory)]
    [String]$TableName,
    [int]$MaxRetries = 3
)

begin {
    function Get-StringValueHC {
        <#
        .SYNOPSIS
            Retrieve a string from the environment variables or a regular
            string.

        .DESCRIPTION
            This function checks the 'Name' property. If the value starts with
            'ENV:', it attempts to retrieve the string value from the specified
            environment variable. Otherwise, it returns the value directly.

        .PARAMETER Name
            Either a string starting with 'ENV:'; a plain text string or NULL.

        .EXAMPLE
            Get-StringValueHC -Name 'ENV:passwordVariable'

            # Output: the environment variable value of $ENV:passwordVariable
            # or an error when the variable does not exist

        .EXAMPLE
            Get-StringValueHC -Name 'mySecretPassword'

            # Output: mySecretPassword

        .EXAMPLE
            Get-StringValueHC -Name ''

            # Output: NULL
        #>
        param (
            [String]$Name
        )

        if (-not $Name) {
            return $null
        }
        elseif (
            $Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            $envVariableName = $Name.Substring(4).Trim()
            $envStringValue = Get-Item -Path "Env:\$envVariableName" -EA Ignore
            if ($envStringValue) {
                return $envStringValue.Value
            }
            else {
                throw "Environment variable '$envVariableName' not found."
            }
        }
        else {
            return $Name
        }
    }
    function New-ServiceNowSessionHC {
        [CmdletBinding()]
        param (
            [parameter(Mandatory)]
            [String]$Uri,
            [parameter(Mandatory)]
            [String]$UserName,
            [parameter(Mandatory)]
            [String]$Password,
            [parameter(Mandatory)]
            [String]$ClientId,
            [parameter(Mandatory)]
            [String]$ClientSecret
        )
        try {
            $userCred = New-Object System.Management.Automation.PSCredential(
                $UserName,
                ($Password | ConvertTo-SecureString -AsPlainText -Force)
            )

            $clientCred = New-Object System.Management.Automation.PSCredential(
                $ClientId,
                ($ClientSecret | ConvertTo-SecureString -AsPlainText -Force)
            )

            Write-Verbose "Create new ServiceNow session to '$Uri'"

            $params = @{
                Url              = $Uri
                Credential       = $userCred
                ClientCredential = $clientCred
            }
            New-ServiceNowSession @params
        }
        catch {
            $errorMessage = $_; $Error.RemoveAt(0)
            throw "Failed to create a ServiceNow session with Uri '$Uri' UserName '$UserName' ClientId '$ClientId': $errorMessage"
        }
    }

    $ErrorActionPreference = 'Stop'

    try {
        #region Import .JSON file
        Write-Verbose "Import .json file '$CredentialsFilePath'"

        $serviceNowJsonFileContent = Get-Content $CredentialsFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
        #endregion

        #region Test .JSON file properties
        Write-Verbose 'Test .json file properties'

        $serviceNowEnvironment = $serviceNowJsonFileContent.($Environment)
    
        if (-not $serviceNowEnvironment) {
            throw "Failed to find environment '$($Environment)' in the ServiceNow environment file '$($CredentialsFilePath)'"
        }
    
        @(
            'Uri', 'UserName', 'Password', 'ClientId', 'ClientSecret'
        ).where(
            { -not $serviceNowEnvironment.$_ }
        ).foreach(
            { 
                throw "Property '$_' not found for environment '$($Environment)' in file '$($CredentialsFilePath)'"
            }
        )
        #endregion
    }
    catch {
        throw "ServiceNow credentials file '$CredentialsFilePath': $_"
    }
}

process {
    #region Import records to upload
    try {
        Write-Verbose 'Import records to upload from .XLSX file'
    
        $params = @{
            Path          = $FormDataExcelFilePath 
            WorksheetName = $ExcelFileWorksheetName
        }
        $recordsToUpload = Import-Excel @params
    }
    catch {
        throw "Failed to import records to upload from file '$FormDataExcelFilePath  with worksheet name '$ExcelFileWorksheetName': $_"
    }
    #endregion
    
    if ($recordsToUpload) {
        #region Create global variable $ServiceNowSession
        $params = @{
            Uri          = Get-StringValueHC -Name $serviceNowEnvironment.Uri
            UserName     = Get-StringValueHC -Name $serviceNowEnvironment.UserName
            Password     = Get-StringValueHC -Name $serviceNowEnvironment.Password
            ClientId     = Get-StringValueHC -Name $serviceNowEnvironment.ClientId
            ClientSecret = Get-StringValueHC -Name $serviceNowEnvironment.ClientSecret
        }
        New-ServiceNowSessionHC @params
        #endregion

        #region Get all table records
        try {
            Write-Verbose "Get all records in ServiceNow table '$TableName'"

            $allTableRecords = Get-ServiceNowRecord -Table $TableName -IncludeTotalCount -First 300
        }
        catch {
            throw "Failed to retrieve all table records in ServiceNow table '$TableName': $_"
        }
        #endregion

        #region Remove all records in the ServiceNow table
        if ($allTableRecords) {
            Write-Verbose "Remove all records in ServiceNow table '$TableName'"

            $currentRecordCount = 0
            $totalRecordCount = $allTableRecords.Count

            $removeParams = @{
                Confirm = $false
                Verbose = $false
            }

            foreach ($tableRecord in $allTableRecords) {
                $attempt = 0
                $currentRecordCount++
    
                while ($attempt -lt $MaxRetries) {
                    $attempt++
                
                    try {
                        Write-Verbose "($currentRecordCount/$totalRecordCount) Remove record '$($tableRecord.sys_id)' $(if($attempt -gt 1) {' - Retry'})"

                        $tableRecord | Remove-ServiceNowRecord @removeParams

                        break
                    }
                    catch {
                        Write-Warning "Failed to remove record '$($tableRecord.sys_id)': $_"
                        
                        Start-Sleep -Seconds 3
                    }
                }
            }
        }
        #endregion

        #region Create new records in the ServiceNow table
        Write-Verbose "Create $($recordsToUpload.Count) records in the ServiceNow table '$TableName'"
        
        $currentRecordCount = 0
        $totalRecordCount = $recordsToUpload.Count

        $createParams = @{
            Table   = $TableName
            Verbose = $false
        }

        foreach ($record in $recordsToUpload) {
            $attempt = 0
            $currentRecordCount++
    
            while ($attempt -lt $MaxRetries) {
                $attempt++
                
                try {
                    Write-Verbose "($currentRecordCount/$totalRecordCount) Create record for matrix '$($record.u_matrixfilename)' AD object '$($record.u_adobjectname)' $(if($attempt -gt 1) {' - Retry'})"

                    $record | New-ServiceNowRecord @createParams

                    break
                }
                catch {
                    Write-Warning "Failed to create record for matrix '$($record.u_matrixfilename)' AD object '$($record.u_adobjectname)': $_"
                        
                    Start-Sleep -Seconds 3
                }
            }
        }
        #endregion
    }
}
