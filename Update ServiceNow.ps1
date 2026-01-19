param (
    [Parameter(Mandatory)]
    [String]$FormDataFile,
    [Parameter(Mandatory)]
    [PSCustomObject]$ServiceNow,

    [string]$PermissionMatrixAdObjectNamesFile = '\\grouphc.net\bnl\LOCAPPS\Scripts\Matrix\2.Nightly\Cherwell\AD object names.csv',
    [string]$PermissionMatrixFormDataFile = '\\grouphc.net\bnl\LOCAPPS\Scripts\Matrix\2.Nightly\Cherwell\Form data.csv',

    [String]$TableName = 'u_bnl_roles',
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
    
    #region Import matrix AD object names
    try {
        Write-Verbose 'Import data from files'
    
        $permissionMatrixAdObjectNames = Import-Csv -LiteralPath $PermissionMatrixAdObjectNamesFile
    }
    catch {
        throw "Failed to read data from file '$PermissionMatrixAdObjectNamesFile': $_"
    }
    #endregion
    
    #region Import matrix form data
    try {
        $permissionMatrixFormData = Import-Csv -LiteralPath $PermissionMatrixFormDataFile   
    }
    catch {
        throw "Failed to read data from file '$PermissionMatrixFormDataFile': $_"
    }
    #endregion
}

process {
    #region Create objects for ServiceNow
    Write-Verbose 'Create objects for ServiceNow'

    $recordsToCreate = foreach (
        $adObjectName in 
        $permissionMatrixAdObjectNames
    ) {
    
        $formData = $permissionMatrixFormData.Where(
            { 
                $adObjectName.MatrixFileName -eq $_.MatrixFileName
            }, 'first'
        )
    
        if ((-not $formData) -or ($formData.MatrixFormStatus -ne 'Enabled')) {
            continue
        }

        $adObjectName | ForEach-Object {
            @{
                u_matrixcategoryname    = $formData.MatrixCategoryName
                u_matrixsubcategoryname = $formData.MatrixSubCategoryName
                u_matrixfilename        = $_.MatrixFileName
                u_matrixresponsible     = $formData.MatrixResponsible
                u_matrixfolderpath      = $formData.MatrixFolderPath 
                u_adobjectname          = $_.SamAccountName
            }
        }
    }
    #endregion
    
    if ($recordsToCreate) {
        if ((-not $ServiceNowSession) -or ($ServiceNowSession.Count -eq 0)) {
            #region Test ServiceNow parameters
            @(
                'CredentialsFilePath', 'Environment', 'TicketFields'
            ).where(
                { -not $ServiceNow.$_ }
            ).foreach(
                { throw "Property 'ServiceNow.$_' not found" }
            )
    
            try {
                $serviceNowJsonFileContent = Get-Content $ServiceNow.CredentialsFilePath -Raw -EA Stop | ConvertFrom-Json
            }
            catch {
                throw "Failed to import the ServiceNow environment file '$($ServiceNow.CredentialsFilePath)': $_"
            }
    
            $serviceNowEnvironment = $serviceNowJsonFileContent.($ServiceNow.Environment)
    
            if (-not $serviceNowEnvironment) {
                throw "Failed to find environment '$($ServiceNow.Environment)' in the ServiceNow environment file '$($ServiceNow.CredentialsFilePath)'"
            }
    
            @(
                'Uri', 'UserName', 'Password', 'ClientId', 'ClientSecret'
            ).where(
                { -not $serviceNowEnvironment.$_ }
            ).foreach(
                { 
                    throw "Property '$_' not found for environment '$($ServiceNow.Environment)' in file '$($ServiceNow.CredentialsFilePath)'"
                }
            )
            #endregion

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
            Write-Verbose "Create $($recordsToCreate.Count) records in the ServiceNow table '$TableName'"
        
            $currentRecordCount = 0
            $totalRecordCount = $recordsToCreate.Count

            $createParams = @{
                Table   = $TableName
                Verbose = $false
            }

            foreach ($record in $recordsToCreate) {
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
}
