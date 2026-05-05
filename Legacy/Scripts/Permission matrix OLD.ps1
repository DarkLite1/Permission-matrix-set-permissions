#Requires -Version 7
#Requires -Modules ImportExcel
#Requires -Modules Toolbox.PermissionMatrix, Toolbox.ActiveDirectory

<#
    .SYNOPSIS
        Apply or verify file and folder permissions.

    .DESCRIPTION
        Read an input file that contains all the parameters for this script.

        This script applies NTFS and SMB permissions to files and folders. It
        reads an Excel file as input and performs the request actions (Check,
        Fix, New).

        Permissions in the Excel file are defined as:
        - L : List
        - R : Read
        - W : Write
        - F : Full Control
        - I : Ignore

    .PARAMETER ConfigurationJsonFile
        Contains all the parameters used by the script.
        See 'Example.json' for a detailed explanation of parameters.
#>

[CmdLetBinding()]
param (
    [Parameter(Mandatory)]
    [String]$ConfigurationJsonFile,
    [HashTable]$ScriptPath = @{
        TestRequirementsFile = "$PSScriptRoot\Test requirements.ps1"
        SetPermissionFile    = "$PSScriptRoot\Set permissions.ps1"
        UpdateServiceNow     = "$PSScriptRoot\Update ServiceNow.ps1"
    }
)

begin {
    function Import-PermissionMatrixModuleHC {
        param(
            [Parameter(Mandatory)]
            [string]$ScriptRoot,

            [Parameter()]
            [ref]$SystemErrors
        )

        try {
            # Local module path first (repo-local override)
            $localModulePath = Join-Path $ScriptRoot 'Modules\Toolbox.PermissionMatrixHC'

            if (Test-Path $localModulePath) {
                Import-Module $localModulePath -Force -ErrorAction Stop
                return
            }
        }
        catch {
            $msg = "Failed to import Toolbox.PermissionMatrixHC module: $_"
        
            if ($SystemErrors) {
                $SystemErrors.Value.Add(
                    [pscustomobject]@{
                        DateTime = Get-Date
                        Message  = $msg
                    }
                )
            }

            throw $msg  # hard stop: script cannot continue without module
        }
    }
    function Invoke-BeginSafe {
        param(
            [scriptblock]$Action,
            [string]$MessageOnError
        )

        if ($script:fatalBeginError) {
            return 
        }

        try {
            & $Action
        }
        catch {
            Add-FatalBeginError "$MessageOnError $_"
        }
    }
    function Add-FatalBeginError {
        param(
            [string]$Message
        )

        Write-Warning $Message

        $systemErrors.Add(
            [PSCustomObject]@{ 
                Message  = $Message 
                DateTime = Get-Date
            }
        )

        $script:fatalBeginError = $true
    }
    
    $ErrorActionPreference = 'stop'
    $fatalBeginError = $false

    # global state
    $systemErrors = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()
    $eventLogData = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()

    # pre-declare major objects so END never sees unbound variables
    $Matrix = $null
    $Export = $null
    $ServiceNow = $null
    $MaxConcurrent = $null
    $Settings = $null
    $PSSessionConfiguration = $nul
    $scriptStartTime = Get-Date

    $mailParams = @{}

    $eventLogData.Add(
        [PSCustomObject]@{
            Message   = 'Script started'
            DateTime  = $scriptStartTime
            EntryType = 'Information'
            EventID   = '100'
        }
    )

    # region Load modules
    Import-PermissionMatrixModuleHC `
        -ScriptRoot $PSScriptRoot `
        -SystemErrors ([ref]$systemErrors)
    #endregion

    #region Import .json file
    Invoke-BeginSafe {
        $script:jsonFileItem = Get-Item -LiteralPath $ConfigurationJsonFile -ErrorAction Stop
        $script:jsonFileContent = Get-Content $jsonFileItem -Raw -Encoding UTF8 | ConvertFrom-Json
    } 'Failed to load JSON file'
    #endregion
    
    #region Set script wide variables
    $Matrix = $jsonFileContent.Matrix
    $Export = $jsonFileContent.Export
    $ServiceNow = $jsonFileContent.ServiceNow
    $MaxConcurrent = $jsonFileContent.MaxConcurrent
    $AdGroupPlaceHolders = $Matrix.AdGroupPlaceHolders
    $Settings = $jsonFileContent.Settings
    $PSSessionConfiguration = $jsonFileContent.PSSessionConfiguration
    $DetailedLog = $Settings.SaveLogFiles.Detailed
    $LogFolder = $Settings.SaveLogFiles.Where.Folder
    #endregion
    
    #region Test script paths exist
    $scriptPathItem = @{}

    foreach ($Key in $ScriptPath.Keys) {
        $Value = $ScriptPath[$Key]

        try {
            $scriptPathItem[$Key] = (Get-Item -LiteralPath $Value -ErrorAction Stop).FullName
        }
        catch {
            Add-FatalBeginError "ScriptPath.$Key '$Value' not found: $_"
        }
    }
    #endregion

    #region Validate input file
    Invoke-BeginSafe {
        $schemaErrors = Validate-JsonSchema -JsonObject $jsonFileContent
        foreach ($e in $schemaErrors) {
            Add-FatalBeginError "Schema validation failed: $($e.Message)"
        }
    } 'JSON schema validation failed'
    #endregion
   
    #region Create log folder
    Invoke-BeginSafe {
        if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
            New-Item -ItemType Directory -Path $LogFolder -ErrorAction Stop
        }
    } "Failed to create log folder '$LogFolder'"
    #endregion

    #region Map share with Excel files
    Invoke-BeginSafe {
        if (-not (Test-Path -LiteralPath 'MatrixFolderPath:\')) {
            $retryCount = 0
            
            $maxRetries = if ($null -ne $Settings.Advanced.DriveMapMaxRetries) { 
                $Settings.Advanced.DriveMapMaxRetries 
            }
            else { 5 }
            
            $sleep = if ($null -ne $Settings.Advanced.DriveMapSleepSeconds) { 
                $Settings.Advanced.DriveMapSleepSeconds 
            }
            else { 5 }

            $isDriveMapped = $false

            while (-not $isDriveMapped -and -not $fatalBeginError) {
                try {
                    New-PSDrive -Root $Matrix.FolderPath -Name 'MatrixFolderPath' -PSProvider FileSystem -Scope Global -ErrorAction Stop
                    $isDriveMapped = $true
                }
                catch {
                    if ($retryCount -ge $maxRetries) {
                        Add-FatalBeginError "Property 'Matrix.FolderPath' path '$($Matrix.FolderPath)' not found: $_"
                        break
                    }

                    Start-Sleep -Seconds $sleep
                    $retryCount++
                }
            }
        }
    } 'Failed during drive mapping'
    #endregion

    #region Get matrix folder path
    Invoke-BeginSafe {
        $script:Matrix.FolderPath = Get-Item -LiteralPath $Matrix.FolderPath -ErrorAction Stop
    } "Matrix folder path '$($Matrix.FolderPath)' not found"    
    #endregion

    #region Get defaults settings
    $mailToDefaultsFile = [System.Collections.Generic.List[string]]::new()

    Invoke-BeginSafe {
        try {
            $script:DefaultsItem = Get-Item -LiteralPath $Matrix.DefaultsFile -ErrorAction Stop
        }
        catch {
            throw "Property 'Matrix.DefaultsFile' path '$($Matrix.DefaultsFile)' not found: $_"
        }
        $DefaultsImport = Import-Excel -Path $DefaultsItem.FullName -Sheet 'Settings' -DataOnly -ErrorAction Stop        

        #region Verify mandatory column headers
        $propDefault = $DefaultsImport[0].PSObject.Properties.Name

        foreach ($Column in @('MailTo', 'ADObjectName', 'Permission')) {
            if ($Column -notin $propDefault) {
                throw "Column header '$Column' not found. The column headers 'MailTo', 'ADObjectName' and 'Permission' are mandatory."
            }
        }
        #endregion

        $DefaultAcl = Get-DefaultAclHC -Sheet $DefaultsImport

        #region Get MailTo
        foreach ($Row in $DefaultsImport) {
            if (-not [string]::IsNullOrWhiteSpace($Row.MailTo)) {
                $script:mailToDefaultsFile.Add($Row.MailTo.ToString().Trim())
            }
        }

        if ($script:mailToDefaultsFile.Count -eq 0) {
            throw "No valid mail addresses found under column header 'MailTo'."
        }
        #endregion
    } "Failed to import default settings file '$($Matrix.DefaultsFile)'"
    #endregion

    #region Archive
    $archivePath = $null

    if ($Matrix.Archive) {
        Invoke-BeginSafe {
            $script:archivePath = Join-Path -Path $Matrix.FolderPath -ChildPath 'Archive'
        
            if (-not (
                    Test-Path -LiteralPath $archivePath -PathType Container)
            ) {
                New-Item -ItemType 'Directory' -Path $archivePath -ErrorAction Stop
            }
        
        } "Failed to create archive folder '$archivePath': $_"
    }
    #endregion
}

process {
    if ($fatalBeginError -or $systemErrors.Count -gt 0) { return }

    try {
        $ID = 0

        $getParams = @{
            Path        = 'MatrixFolderPath:\*'
            Filter      = '*.xlsx'
            ErrorAction = 'Stop'
        }

        #region Get matrix files
        $matrixFiles = @(Get-ChildItem @getParams).Where(
            { 
                (-not $_.PSIsContainer) -and
                ($_.FullName -ne $DefaultsItem.FullName)
            }
        )

        Write-Verbose "Found $($matrixFiles.Count) matrix Excel files"
        #endregion

        if ($matrixFiles.Count -eq 0) {
            Write-Verbose 'No matrix Excel files found'

            return
        }

        #region Create dated log folder
        $datedLogFolderPath = $null

        if ($matrixFiles) {
            $datedLogFolderPath = Get-DatedLogFolderPathHC
        }
        #endregion

        $scriptBlock = {
            param (
                $matrixFile,
                $Matrix,
                $Export,
                $archivePath,
                $eventLogData,
                $datedLogFolderPath,
                $VerbosePreference,
                $ErrorActionPreference
            )

            try {
                Write-Verbose "Matrix file '$($matrixFile.Name)'"

                $Obj = [PSCustomObject]@{
                    File        = @{
                        Item         = $matrixFile
                        SaveFullName = $matrixFile.FullName
                        ExcelInfo    = $null
                        LogFolder    = $null
                        Check        = [System.Collections.Generic.List[PSCustomObject]]::new()
                    }
                    Settings    = [System.Collections.Generic.List[PSCustomObject]]::new()
                    Permissions = @{
                        Import = @()
                        Check  = [System.Collections.Generic.List[PSCustomObject]]::new()
                    }
                    FormData    = @{
                        Import = $null
                        Check  = [System.Collections.Generic.List[PSCustomObject]]::new()
                    }
                }

                #region Create matrix log folder
                try {
                    $matrixLogFolderPath = Join-Path -Path $datedLogFolderPath -ChildPath $matrixFile.BaseName

                    if (-not
                        (Test-Path -LiteralPath $matrixLogFolderPath -PathType Container)
                    ) {
                        $null = New-Item -ItemType 'Directory' -Path $matrixLogFolderPath -ErrorAction Stop
                    }

                    $Obj.File.LogFolder = $matrixLogFolderPath

                    Write-Verbose "Matrix log folder '$($Obj.File.LogFolder)'"
                }
                catch {
                    throw "Failed to create log folder '$matrixLogFolderPath': $_"
                }
                #endregion

                #region Copy file to log folder
                try {
                    $copyParams = @{
                        LiteralPath = $matrixFile.FullName
                        Destination = $Obj.File.LogFolder
                        PassThru    = $true
                        ErrorAction = 'Stop'
                    }

                    Write-Verbose "Copy file '$($copyParams.LiteralPath)' to '$($copyParams.Destination)'"

                    $Obj.File.SaveFullName = (Copy-Item @copyParams).FullName
                }
                catch {
                    throw "Failed to copy file '$($copyParams.LiteralPath)' to '$($copyParams.Destination)': $_"
                }
                #endregion

                #region Get Excel file details
                $Obj.File.ExcelInfo = Get-ExcelWorkbookInfo -Path $matrixFile.FullName -ErrorAction Stop

                Write-Verbose "File '$($matrixFile.Name)': LastModifiedBy '$($Obj.File.ExcelInfo.LastModifiedBy)' LastModifiedDate '$($Obj.File.ExcelInfo.Modified.ToString('dd/MM/yyyy HH:mm:ss'))'"
                #endregion

                #region Import sheets Settings, Permissions, FormData
                try {
                    #region Import sheet Settings
                    $verboseMessage = "File '$($matrixFile.Name)': Import worksheet 'Settings'"

                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = $verboseMessage
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '2'
                        }
                    )
                    Write-Verbose $verboseMessage

                    $importParams = @{
                        Path        = $matrixFile.FullName
                        DataOnly    = $true
                        ErrorAction = 'Stop'
                    }

                    $settingsSheet = @(
                        Import-Excel @importParams -Sheet 'Settings'
                    ).Where(
                        { $_.Status -eq 'Enabled' }
                    )
                    #endregion

                    if ($settingsSheet) {
                        foreach ($S in $settingsSheet) {
                            $Obj.Settings.Add(
                                [PSCustomObject]@{
                                    ID        = $null
                                    Import    = Format-SettingStringsHC -Settings $S
                                    Check     = [System.Collections.Generic.List[PSCustomObject]]::new()
                                    Matrix    = [System.Collections.Generic.List[PSCustomObject]]::new()
                                    AdObjects = @{}
                                    JobTime   = @{}
                                }
                            )
                        }

                        #region Import sheet Permissions
                        $verboseMessage = "File '$($matrixFile.Name)': Import worksheet 'Permissions'"

                        $eventLogData.Add(
                            [PSCustomObject]@{
                                Message   = $verboseMessage
                                DateTime  = Get-Date
                                EntryType = 'Information'
                                EventID   = '2'
                            }
                        )
                        Write-Verbose $verboseMessage

                        $Obj.Permissions.Import = @(
                            Import-Excel @importParams -Sheet 'Permissions' -NoHeader | Format-PermissionsStringsHC
                        )
                        #endregion

                        #region Import sheet FormData
                        if (
                            $Export.ServiceNowFormDataExcelFile -or
                            $Export.OverviewHtmlFile
                        ) {
                            try {
                                $verboseMessage = "File '$($matrixFile.Name)': Import worksheet 'FormData'"

                                $eventLogData.Add(
                                    [PSCustomObject]@{
                                        Message   = $verboseMessage
                                        DateTime  = Get-Date
                                        EntryType = 'Information'
                                        EventID   = '2'
                                    }
                                )
                                Write-Verbose $verboseMessage

                                $formData = Import-Excel @importParams -Sheet 'FormData' -ErrorVariable importFail

                                $formDataValidation = Test-FormDataHC $formData
                                if ($formDataValidation) {
                                    $Obj.FormData.Check.Add($formDataValidation)
                                }
                                else {
                                    $Obj.FormData.Import = $formData[0]
                                }
                            }
                            catch {
                                $Obj.File.Check.Add(
                                    [PSCustomObject]@{
                                        Type        = 'FatalError'
                                        Name        = "Worksheet 'FormData' not found"
                                        Description = "When the argument 'Export.ServiceNowFormDataExcelFile' is used the Excel file needs to have a worksheet 'FormData'."
                                        Value       = @($_)
                                    }
                                )
                            }
                        }
                        #endregion
                    }
                    else {
                        $Obj.File.Check.Add(
                            [PSCustomObject]@{
                                Type        = 'Warning'
                                Name        = 'Matrix disabled'
                                Description = 'Every Excel file needs at least one enabled matrix.'
                                Value       = "The worksheet 'Settings' does not contain a row with 'Status' set to 'Enabled'."
                            }
                        )

                        $verboseMessage = "File '$($matrixFile.Name)': No lines found with status 'Enabled' in the worksheet 'Settings'"

                        $eventLogData.Add(
                            [PSCustomObject]@{
                                Message   = $verboseMessage
                                DateTime  = Get-Date
                                EntryType = 'Information'
                                EventID   = '2'
                            }
                        )
                        Write-Warning $verboseMessage
                    }
                }
                catch {
                    $errorMessage = switch -Wildcard ($_) {
                        "*Worksheet 'Settings' not found*" {
                            "Worksheet 'Settings' not found"; break
                        }
                        "*worksheet 'Settings': No column headers found on top row '1'*" {
                            "Worksheet 'Settings' is empty"; break
                        }
                        "*Worksheet 'Permissions' not found*" {
                            "Worksheet 'Permissions' not found"; break
                        }
                        "*worksheet 'Permissions': No column headers found on top row '1'*" {
                            "Worksheet 'Permissions' is empty"; break
                        }
                        default {
                            "Failed importing the Excel file '$($matrixFile.FullName)': $_"
                        }
                    }
                    $Obj.File.Check.Add(
                        [PSCustomObject]@{
                            Type        = 'FatalError'
                            Name        = 'Excel file incorrect'
                            Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                            Value       = $errorMessage
                        }
                    )
                }
                #endregion

                if ($archivePath) {
                    try {
                        $verboseMessage = "File '$($matrixFile.Name)': Move file to archive folder '$archivePath'"

                        $eventLogData.Add(
                            [PSCustomObject]@{
                                Message   = $verboseMessage
                                DateTime  = Get-Date
                                EntryType = 'Information'
                                EventID   = '2'
                            }
                        )
                        Write-Verbose $verboseMessage

                        Move-Item -LiteralPath $matrixFile -Destination $archivePath -Force -ErrorAction Stop
                    }
                    catch {
                        $Obj.File.Check.Add(
                            [PSCustomObject]@{
                                Type        = 'Warning'
                                Name        = 'Archiving failed'
                                Description = "When the '-Archive' switch is used the file is moved to the archive folder. In case a file is still in use, the move operation might fail."
                                Value       = @($_)
                            }
                        )
                    }
                }

                $Obj
            }
            catch {
                $verboseMessage = "File '$($matrixFile.Name)': $_"

                $systemErrors.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = $verboseMessage
                    }
                )

                Write-Warning $verboseMessage
            }
        }

        #region Run code serial or parallel
        $importedMatrix = if ($MaxConcurrent.Computers -eq 1) {
            $matrixFiles | ForEach-Object {
                $params = @{
                    matrixFile            = $_
                    Matrix                = $Matrix
                    Export                = $Export
                    archivePath           = $archivePath
                    eventLogData          = $eventLogData
                    datedLogFolderPath    = $datedLogFolderPath
                    VerbosePreference     = $VerbosePreference
                    ErrorActionPreference = $ErrorActionPreference
                }
                & $scriptBlock @params
            }
        }
        else {
            $processScriptBlockString = $scriptBlock.ToString()

            $matrixFiles |
            ForEach-Object -ThrottleLimit $MaxConcurrent.Computers -Parallel {
                $params = @{
                    matrixFile            = $_
                    Matrix                = $using:Matrix
                    Export                = $using:Export
                    archivePath           = $using:archivePath
                    eventLogData          = $using:eventLogData
                    datedLogFolderPath    = $using:datedLogFolderPath
                    VerbosePreference     = $using:VerbosePreference
                    ErrorActionPreference = $using:ErrorActionPreference
                }

                $rehydratedBlock = [scriptblock]::Create($using:processScriptBlockString)

                & $rehydratedBlock @params
            }
        }
        #endregion

        #region Assign unique ID to each matrix
        $matrixId = 0

        foreach (
            $I 
            in $importedMatrix | Sort-Object -Property { $_.File.Item.Name }
        ) {
            foreach (
                $S in
                $I.Settings | Sort-Object -Property ComputerName, Path
            ) {
                $matrixId++
                $S.ID = $matrixId
            }
        }
        #endregion

        if ($importedMatrix) {
            #region Build FormData for Export folder
            foreach ($I in $importedMatrix) {
                if (-not $I.FormData.Import) { continue }

                try {
                    $property = @{}

                    #region Convert MatrixResponsible to UserPrincipalName
                    $responsibleRaw = $I.FormData.Import.MatrixResponsible

                    $namesToProcess = if (
                        -not [string]::IsNullOrWhiteSpace($responsibleRaw)
                    ) {
                        $responsibleRaw.Split(',').Trim()
                    }
                    else {
                        @()
                    }

                    $params = @{
                        Name                  = $namesToProcess
                        ExcludeSamAccountName = $AdGroupPlaceHolders
                    }
                    $result = Get-AdUserPrincipalNameHC @params

                    $property.MatrixResponsible = $result.userPrincipalName -join ','

                    if ($result.notFound) {
                        $I.FormData.Check.Add(
                            [PSCustomObject]@{
                                Type        = 'Warning'
                                Name        = 'AD object not found'
                                Description = "The email address or SamAccountName is not found in the active directory. Multiple entries are supported with the comma ',' separator."
                                Value       = $result.notFound
                            }
                        )
                    }
                    #endregion

                    #region Add MatrixFilePath and MatrixFileName
                    $property.MatrixFilePath = if ($Matrix.Archive) {
                        Join-Path -Path $archivePath -ChildPath $I.File.Item.Name
                    }
                    else {
                        $I.File.Item.FullName
                    }

                    $property.MatrixFileName = $I.File.Item.BaseName
                    #endregion

                    $I.FormData.Import |
                    Add-Member -NotePropertyMembers $property -Force
                }
                catch {
                    $I.FormData.Check.Add(
                        [PSCustomObject]@{
                            Type        = 'FatalError'
                            Name        = 'Failed adding property'
                            Description = "The worksheet 'FormData' could not be updated correctly."
                            Value       = @($_)
                        }
                    )
                }
            }
            #endregion

            #region Build the matrix and check for incorrect input
            $verboseMessage = 'Build the matrix and check for incorrect input'

            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = $verboseMessage
                    DateTime  = Get-Date
                    EntryType = 'Information'
                    EventID   = '2'
                }
            )
            Write-Verbose $verboseMessage

            $scriptBlock = {
                param (
                    $I,
                    $eventLogData,
                    $VerbosePreference,
                    $ErrorActionPreference
                )

                Import-Module -Name Toolbox.PermissionMatrix -ErrorAction Stop

                if (
                    ($I.File.Check.Type -contains 'FatalError') -or
                    (-not $I.Settings)
                ) {
                    return $I
                }

                try {
                    Write-Verbose "Test matrix permissions for '$($I.File.Item.BaseName)'"

                    $permCheck = Test-MatrixPermissionsHC -Permissions $I.Permissions.Import
                    if ($permCheck) { $I.Permissions.Check.Add($permCheck) }

                    if ($I.Permissions.Check.Type -notcontains 'FatalError') {
                        foreach ($S in $I.Settings) {
                            $settingCheck = Test-MatrixSettingHC -Setting $S.Import
                            if ($settingCheck) { $S.Check.Add($settingCheck) }

                            #region Create AD object names
                            Write-Verbose "Create AD object names for '$($I.File.Item.BaseName)'"

                            $params = @{
                                Begin         = $S.Import.GroupName
                                Middle        = $S.Import.SiteCode
                                ColumnHeaders = $I.Permissions.Import | Select-Object -First 3
                            }
                            $adObjects = ConvertTo-MatrixADNamesHC @params

                            Write-Verbose "Test AD objects for '$($I.File.Item.BaseName)'"

                            $adCheck = Test-AdObjectsHC -ADObjects $adObjects
                            if ($adCheck) { $S.Check.Add($adCheck) }
                            #endregion

                            #region Create matrix for each settings line
                            if ($S.Check.Type -notcontains 'FatalError') {
                                Write-Verbose "Create matrix for each settings line in '$($I.File.Item.BaseName)'"

                                $S.AdObjects = $adObjects

                                $params = @{
                                    NonHeaderRows = $I.Permissions.Import | Select-Object -Skip 3
                                    ADObjects     = $adObjects
                                }

                                $aclMatrix = ConvertTo-MatrixAclHC @params
                                if ($aclMatrix) { $S.Matrix.Add($aclMatrix) }
                            }
                            #endregion
                        }
                    }
                }
                catch {
                    $I.File.Check.Add(
                        [PSCustomObject]@{
                            Type        = 'FatalError'
                            Name        = 'Unknown error'
                            Description = 'While checking the input and generating the matrix an error was reported.'
                            Value       = @($_)
                        }
                    )
                }

                return $I
            }

            #region Run code serial or parallel
            $importedMatrix = if ($MaxConcurrent.Computers -eq 1) {
                $importedMatrix | ForEach-Object {
                    $params = @{
                        I                     = $_
                        VerbosePreference     = $VerbosePreference
                        ErrorActionPreference = $ErrorActionPreference
                    }
                    & $scriptBlock @params
                }
            }
            else {
                $processScriptBlockString = $scriptBlock.ToString()

                $importedMatrix |
                ForEach-Object -ThrottleLimit $MaxConcurrent.Computers -Parallel {
                    $params = @{
                        I                     = $_
                        VerbosePreference     = $using:VerbosePreference
                        ErrorActionPreference = $using:ErrorActionPreference
                    }

                    $rehydratedBlock = [scriptblock]::Create($using:processScriptBlockString)

                    & $rehydratedBlock @params
                }
            }
            #endregion
            #endregion

            #region Test duplicate ComputerName/Path combination
            Write-Verbose 'Check duplicate ComputerName/Path combination'

            $duplicateSettings = $importedMatrix.Settings |
            Group-Object -Property { $_.Import.ComputerName }, { $_.Import.Path } |
            Where-Object Count -GE 2

            foreach ($DupGroup in $duplicateSettings) {
                foreach ($Setting in $DupGroup.Group) {

                    # Because these objects crossed a runspace boundary,
                    # Check might be an Object[] array now
                    # We must use += or cast it safely to add elements.
                    $Setting.Check += [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Duplicate ComputerName/Path combination'
                        Description = "Every 'ComputerName' combined with a 'Path' needs to be unique over all the 'Settings' worksheets found in all the active matrix files."
                        Value       = @{
                            ComputerName = $Setting.Import.ComputerName
                            Path         = $Setting.Import.Path
                        }
                    }
                }
            }
            #endregion

            #region Test expanded matrix and get AD object details
            Write-Verbose 'Check expanded matrix'

            $AdObjects = $importedMatrix.Settings.Matrix.ACL.Keys |
            Sort-Object -Unique

            $adObjectHash = @{}
            $groupManagerHash = @{}

            if ($AdObjects.Count -gt 0) {
                Write-Verbose "Get AD object details for $($AdObjects.Count) unique objects"

                $params = @{
                    ADObjectName = $AdObjects
                    Type         = 'SamAccountName'
                }
                $ADObjectDetails = @(Get-ADObjectDetailHC @params)

                #region Build Hash Table for lookups
                foreach ($ad in $ADObjectDetails) {
                    if (-not [string]::IsNullOrWhiteSpace($ad.samAccountName)) {
                        $adObjectHash[$ad.samAccountName] = $ad
                    }
                }
                #endregion

                foreach ($S in $importedMatrix.Settings) {
                    if (-not $S.Matrix) { continue }

                    Write-Verbose "Test expanded matrix for Settings row ComputerName '$($S.Import.ComputerName)' Path '$($S.Import.Path)' SiteName '$($S.Import.SiteName)' SiteCode '$($S.Import.SiteCode)' GroupName '$($S.Import.GroupName)'"

                    $params = @{
                        Matrix                 = $S.Matrix
                        ADObject               = $ADObjectDetails
                        DefaultAcl             = $DefaultAcl
                        AdGroupPlaceHolders = $AdGroupPlaceHolders
                    }

                    $expandedCheck = Test-ExpandedMatrixHC @params

                    if ($expandedCheck) {
                        $S.Check += $expandedCheck | ConvertTo-StructuredObjectHC
                    }
                }
            }
            #endregion

            #region Get AD object details for group managers
            $groupManagers = $ADObjectDetails.ADObject.ManagedBy |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique

            if ($groupManagers.Count -gt 0) {
                $verboseMessage = "Retrieve AD object details for $($groupManagers.Count) group managers"

                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = $verboseMessage
                        DateTime  = Get-Date
                        EntryType = 'Information'
                        EventID   = '2'
                    }
                )
                Write-Verbose $verboseMessage

                $params = @{
                    ADObjectName = $groupManagers
                    Type         = 'DistinguishedName'
                }
                $groupManagersAdDetails = @(Get-ADObjectDetailHC @params)

                #region Build Hash Table for Lookups
                if ($groupManagersAdDetails) {
                    foreach ($gm in $groupManagersAdDetails) {
                        if (-not [string]::IsNullOrWhiteSpace($gm.DistinguishedName)) {
                            $groupManagerHash[$gm.DistinguishedName] = $gm
                        }
                    }
                }
                #endregion
            }
            #endregion

            #region Remove group members that are in the AdGroupPlaceHolders
            if ($AdGroupPlaceHolders) {
                $allAdObjects = @($ADObjectDetails) + @($groupManagersAdDetails)

                foreach ($adObject in $allAdObjects) {
                    if (-not $adObject.adGroupMember) { continue }

                    $adObject.adGroupMember = @(
                        $adObject.adGroupMember.Where(
                            { $_.SamAccountName -notin $AdGroupPlaceHolders }
                        )
                    )
                }
            }
            #endregion

            #region Test server requirements
            $executableMatrix = @(Get-ExecutableMatrixHC -From $importedMatrix)

            if ($executableMatrix.Count -gt 0) {
                $verboseMessage = 'Test server requirements'

                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = $verboseMessage
                        DateTime  = Get-Date
                        EntryType = 'Information'
                        EventID   = '2'
                    }
                )
                Write-Verbose $verboseMessage

                $testRequirementsBlock = {
                    param (
                        $computerName,
                        $pathsToCheck,
                        $scriptPathItem,
                        $PSSessionConfiguration,
                        $VerbosePreference,
                        $ErrorActionPreference
                    )

                    try {
                        $params = @{
                            FilePath          = $scriptPathItem.TestRequirementsFile
                            ArgumentList      = $pathsToCheck, $true
                            ConfigurationName = $PSSessionConfiguration
                            ComputerName      = $computerName
                            ErrorAction       = 'Stop'
                        }

                        $result = Invoke-Command @params

                        return [PSCustomObject]@{
                            ComputerName = $computerName
                            Result       = $result
                        }
                    }
                    catch {
                        $problem = [PSCustomObject]@{
                            Type        = 'FatalError'
                            Name        = 'Computer requirements'
                            Value       = @($_)
                            Description = "Failed checking the computer for the minimal requirements with the 'Test requirements' script."
                        }
                        return [PSCustomObject]@{
                            ComputerName = $computerName
                            Result       = $problem
                        }
                    }
                }

                $matrixGroups = $executableMatrix |
                Group-Object -Property { $_.Import.ComputerName }

                $testRequirementsBlockString = $testRequirementsBlock.ToString()

                # DTO FLATTENING: Protects deep properties from runspace truncation
                $safeReqGroups = foreach ($group in $matrixGroups) {
                    [PSCustomObject]@{
                        ComputerName = $group.Name
                        PathsToCheck = @($group.Group.Import.Path)
                    }
                }

                #region Run code serial or parallel
                $runspaceOutput = if ($MaxConcurrent.Computers -eq 1) {
                    $safeReqGroups | ForEach-Object {
                        $params = @{
                            computerName           = $_.ComputerName
                            pathsToCheck           = $_.PathsToCheck
                            scriptPathItem         = $scriptPathItem
                            PSSessionConfiguration = $PSSessionConfiguration
                            VerbosePreference      = $VerbosePreference
                            ErrorActionPreference  = $ErrorActionPreference
                        }
                        & $testRequirementsBlock @params
                    }
                }
                else {
                    $safeReqGroups | ForEach-Object -ThrottleLimit $MaxConcurrent.Computers -Parallel {
                        $params = @{
                            computerName           = $_.ComputerName
                            pathsToCheck           = $_.PathsToCheck
                            scriptPathItem         = $using:scriptPathItem
                            PSSessionConfiguration = $using:PSSessionConfiguration
                            VerbosePreference      = $using:VerbosePreference
                            ErrorActionPreference  = $using:ErrorActionPreference
                        }

                        $rehydratedBlock = [scriptblock]::Create($using:testRequirementsBlockString)
                        & $rehydratedBlock @params
                    }
                }
                #endregion

                # Main Thread Application
                foreach ($output in $runspaceOutput) {
                    if ($output -and $output.Result) {
                        $targetGroups = $matrixGroups.Where(
                            { $_.Name -eq $output.ComputerName }
                        )
                        foreach ($group in $targetGroups) {
                            foreach ($matrix in $group.Group) {
                                $matrix.Check += $output.Result | ConvertTo-StructuredObjectHC
                            }
                        }
                    }
                }
            }
            #endregion

            #region Set permissions
            if (
                $executableMatrix = @(
                    Get-ExecutableMatrixHC -From $importedMatrix)
            ) {
                $verboseMessage = "Start 'Set permissions' script for '$($executableMatrix.Count)' matrix"

                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = $verboseMessage
                        DateTime  = Get-Date
                        EntryType = 'Information'
                        EventID   = '2'
                    }
                )
                Write-Verbose $verboseMessage

                #region Add default permissions
                if ($DefaultAcl.Count -ne 0) {
                    foreach (
                        $acl in
                        @($executableMatrix.Matrix.ACL).Where(
                            { $_.Count -ne 0 }
                        )
                    ) {
                        $DefaultAcl.GetEnumerator().Where(
                            { -not $acl.ContainsKey($_.Key) }
                        ).Foreach(
                            { $acl.Add($_.Key, $_.Value) }
                        )
                    }
                }
                #endregion

                # 1. INNER BLOCK
                $innerScriptBlock = {
                    param (
                        $matrixFileDto, # Flat object
                        $scriptPathItem,
                        $PSSessionConfiguration,
                        $MaxConcurrent,
                        $DetailedLog
                    )
                    try {
                        $startTime = Get-Date

                        # Restore the matrix array safely from JSON
                        $restoredMatrix = if (
                            -not [string]::IsNullOrWhiteSpace($matrixFileDto.MatrixJson)
                        ) {
                            @($matrixFileDto.MatrixJson | ConvertFrom-Json)
                        }
                        else { @() }

                        $params = @{
                            FilePath          = $scriptPathItem.SetPermissionFile
                            ArgumentList      = $matrixFileDto.Path, $matrixFileDto.Action, $restoredMatrix, $MaxConcurrent.FoldersPerMatrix, $DetailedLog
                            ConfigurationName = $PSSessionConfiguration
                            ComputerName      = $matrixFileDto.ComputerName
                            ErrorAction       = 'Stop'
                        }

                        $result = Invoke-Command @params

                        return [PSCustomObject]@{
                            ID       = $matrixFileDto.ID
                            Result   = $result
                            JobStart = $startTime
                            JobEnd   = Get-Date
                        }
                    }
                    catch {
                        return [PSCustomObject]@{
                            ID       = $matrixFileDto.ID
                            Result   = [PSCustomObject]@{
                                Type        = 'FatalError'
                                Name        = 'Set permissions'
                                Value       = @($_)
                                Description = "Failed applying action '$($matrixFileDto.Action)' with the 'Set permissions' script."
                            }
                            JobStart = $startTime
                            JobEnd   = Get-Date
                        }
                    }
                }

                $innerScriptBlockString = $innerScriptBlock.ToString()

                # 2. OUTER BLOCK
                $outerScriptBlock = {
                    param (
                        $ComputerGroupDto, # Safe Group object
                        $scriptPathItem,
                        $PSSessionConfiguration,
                        $MaxConcurrent,
                        $DetailedLog,
                        $innerScriptBlockString,
                        $innerScriptBlock
                    )

                    $matrixes = $ComputerGroupDto.Matrices

                    $innerResults = if (
                        $MaxConcurrent.JobsPerRemoteComputer -eq 1
                    ) {
                        # SERIAL
                        $matrixes | ForEach-Object {
                            & $innerScriptBlock -matrixFileDto $_ `
                                -scriptPathItem $scriptPathItem `
                                -PSSessionConfiguration $PSSessionConfiguration `
                                -MaxConcurrent $MaxConcurrent `
                                -DetailedLog $DetailedLog
                        }
                    }
                    else {
                        # PARALLEL
                        $matrixes | ForEach-Object -ThrottleLimit $MaxConcurrent.JobsPerRemoteComputer -Parallel {
                            $rehydratedInner = [scriptblock]::Create($using:innerScriptBlockString)

                            & $rehydratedInner -matrixFileDto $_ `
                                -scriptPathItem ($using:scriptPathItem) `
                                -PSSessionConfiguration ($using:PSSessionConfiguration) `
                                -MaxConcurrent ($using:MaxConcurrent) `
                                -DetailedLog ($using:DetailedLog)
                        }
                    }

                    return $innerResults
                }

                $outerScriptBlockString = $outerScriptBlock.ToString()

                # 3. KICKOFF
                $computerGroups = $executableMatrix |
                Group-Object -Property { $_.Import.ComputerName }

                # DTO FLATTENING: Build a shallow array and wrap the deep array in JSON
                $safeGroups = foreach ($group in $computerGroups) {
                    [PSCustomObject]@{
                        ComputerName = $group.Name
                        Matrices     = @(
                            foreach ($S in $group.Group) {
                                [PSCustomObject]@{
                                    ID           = $S.ID
                                    ComputerName = $S.Import.ComputerName
                                    Path         = $S.Import.Path
                                    Action       = $S.Import.Action
                                    MatrixJson   = ($S.Matrix | ConvertTo-Json -Depth 10 -Compress)
                                }
                            }
                        )
                    }
                }

                $allJobResults = if ($MaxConcurrent.Computers -eq 1) {
                    # SERIAL
                    $safeGroups | ForEach-Object {
                        & $outerScriptBlock -ComputerGroupDto $_ `
                            -scriptPathItem $scriptPathItem `
                            -PSSessionConfiguration $PSSessionConfiguration `
                            -MaxConcurrent $MaxConcurrent `
                            -DetailedLog $DetailedLog `
                            -innerScriptBlockString $innerScriptBlockString `
                            -innerScriptBlock $innerScriptBlock
                    }
                }
                else {
                    # PARALLEL
                    $safeGroups | ForEach-Object -ThrottleLimit $MaxConcurrent.Computers -Parallel {
                        $rehydratedOuter = [scriptblock]::Create($using:outerScriptBlockString)

                        & $rehydratedOuter -ComputerGroupDto $_ `
                            -scriptPathItem ($using:scriptPathItem) `
                            -PSSessionConfiguration ($using:PSSessionConfiguration) `
                            -MaxConcurrent ($using:MaxConcurrent) `
                            -DetailedLog ($using:DetailedLog) `
                            -innerScriptBlockString ($using:innerScriptBlockString)
                    }
                }

                # 4. MAIN THREAD APPLICATION
                foreach (
                    $payload in
                    @($allJobResults).Where({ $_ -ne $null })
                ) {
                    $matchedMatrices = $executableMatrix.Where(
                        { $_.ID -eq $payload.ID }
                    )

                    foreach ($liveMatrix in $matchedMatrices) {
                        if ($payload.Result) {
                            $liveMatrix.Check += $payload.Result | 
                            ConvertTo-StructuredObjectHC
                        }
                        $liveMatrix.JobTime.Start = $payload.JobStart
                        $liveMatrix.JobTime.End = $payload.JobEnd
                        $liveMatrix.JobTime.Duration = New-TimeSpan -Start $payload.JobStart -End $payload.JobEnd
                    }
                }
            }
            #endregion
        }
    }
    catch {
        Write-Warning $_

        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = $_
            }
        )
    }
}

end {
    function Build-Counters {
        param(
            [Parameter(Mandatory = $false)][array]$ImportedMatrix,
            [Parameter(Mandatory = $false)][object]$SystemErrors
        )

        #
        # Prepare result structure
        #
        $counters = [ordered]@{
            TotalErrors   = 0
            TotalWarnings = 0
            File          = @{ Error = 0; Warning = 0 }
            Settings      = @{ Error = 0; Warning = 0 }
            Permissions   = @{ Error = 0; Warning = 0 }
            FormData      = @{ Error = 0; Warning = 0 }
        }

        #
        # If no matrix data, only return system errors
        #
        if (-not $ImportedMatrix) {
            $counters.TotalErrors = $SystemErrors.Count
            return [PSCustomObject]$counters
        }

        #
        # Helper: Count error/warning types in a collection
        #
        function Count-CheckTypes {
            param([object[]]$Checks)

            return @{
                Error   = @(
                    $Checks | Where-Object { $_.Type -eq 'FatalError' }
                ).Count
                Warning = @(
                    $Checks | Where-Object { $_.Type -eq 'Warning' }
                ).Count
            }
        }

        #
        # Gather all check collections
        #
        $fileChecks = $ImportedMatrix | ForEach-Object { 
            $_.File.Check ?? @()
        }
        $settingsChecks = $ImportedMatrix | ForEach-Object { 
            $_.Settings ?? @() } | ForEach-Object { $_.Check ?? @() 
        }
        $permissionChecks = $ImportedMatrix | ForEach-Object { 
            $_.Permissions.Check ?? @() 
        }
        $formDataChecks = $ImportedMatrix | ForEach-Object {
            $_.FormData.Check ?? @() 
        }

        #
        # Count all categories
        #
        $counters.File = Count-CheckTypes $fileChecks
        $counters.Settings = Count-CheckTypes $settingsChecks
        $counters.Permissions = Count-CheckTypes $permissionChecks
        $counters.FormData = Count-CheckTypes $formDataChecks

        #
        # Totals
        #
        $counters.TotalErrors =
        $counters.File.Error +
        $counters.Settings.Error +
        $counters.Permissions.Error +
        $counters.FormData.Error +
        $SystemErrors.Count

        $counters.TotalWarnings =
        $counters.File.Warning +
        $counters.Settings.Warning +
        $counters.Permissions.Warning +
        $counters.FormData.Warning

        return [PSCustomObject]$counters
    }    

    try {
        #region Ensure Settings and LogFolder are defined
        $Settings = Ensure-SafeSettingsHC $Settings
        
        $LogFolder = Ensure-LogFolderHC `
            -RequestedFolder $Settings.SaveLogFiles.Where.Folder `
            -SystemErrors ([ref]$systemErrors)
        #endregion

        if (-not $fatalBeginError) {
            #
            # 1. VALIDATE CONFIGURATION
            #
            $validation = Validate-Settings `
                -Settings $Settings `
                -Matrix $Matrix `
                -Export $Export `
                -ServiceNow $ServiceNow `
                -MaxConcurrent $MaxConcurrent

            foreach ($err in $validation.Errors) {
                $systemErrors.Add($err)
            }

            if (-not $validation.IsValid) {
                Write-Warning 'Configuration validation failed. Aborting end block.'
                return
            }
        }

        #
        # 2. INITIALIZE HTML STRUCTURE
        #
        $html = Initialize-HtmlStructure
       
        #
        # 3. PROCESS MATRICES AND BUILD HTML TABLES
        #
        if ($importedMatrix) {
            $importedMatrix = Process-MatrixObjects `
                -ImportedMatrix $importedMatrix `
                -Html $html

            $html.MatrixTables = Build-MatrixEmailHtml `
                -ImportedMatrix $importedMatrix `
                -Html $html
        }

        #
        # 4. COLLECT EXPORT DATA
        #
        $dataToExport = if (
            $importedMatrix -and
            (
                $Export.ServiceNowFormDataExcelFile -or
                $Export.PermissionsExcelFile -or
                $Export.OverviewHtmlFile
            )
        ) {
            Build-ExportData `
                -ImportedMatrix $importedMatrix `
                -AdObjectHash $adObjectHash `
                -GroupManagerHash $groupManagerHash
        }

        #
        # 5. EXPORT FILES
        #
        $exportedFiles = @{}
        if ($dataToExport -and $systemErrors.Count -eq 0) {
            
            $exportLogFolderPath = Join-Path (Get-DatedLogFolderPathHC) 'Export'

            if (-not (Test-Path -LiteralPath $exportLogFolderPath)) {
                New-Item -ItemType Directory -Path $exportLogFolderPath -ErrorAction SilentlyContinue | Out-Null
            }

            $exportedFiles = Export-Files `
                -DataToExport $dataToExport `
                -ExportConfig $Export `
                -ServiceNowConfig $ServiceNow `
                -ExportLogFolder $exportLogFolderPath `
                -ScriptPathItem $scriptPathItem `
                -SystemErrors ([ref]$systemErrors)
        }

        #
        # 6. BUILD COUNTERS AND ERROR TABLES
        #
        $counter = Build-Counters `
            -ImportedMatrix $importedMatrix `
            -SystemErrors $systemErrors

        $html.ErrorWarningTable = Build-ErrorWarningTable `
            -CounterData $counter `
            -SystemErrors $systemErrors


        #
        # 7. WRITE EVENT LOG
        #
        Write-EventLogSafe `
            -EventLogData $eventLogData `
            -ScriptName $Settings.ScriptName `
            -Settings $Settings `
            -SystemErrors ([ref]$systemErrors)

        #
        # 8. CLEANUP OLD LOGS AND SAVE SYSTEM ERROR DUMP
        #
        Cleanup-OldLogs `
            -LogFolder $LogFolder `
            -RetentionDays $Settings.SaveLogFiles.DeleteLogsAfterDays `
            -SystemErrors ([ref]$systemErrors)

        Write-SystemErrorLog `
            -SystemErrors $systemErrors `
            -LogFolder $LogFolder `
            -MailParams ([ref]$mailParams)


        #
        # 9. BUILD AND SEND MAIL
        #
        $mailParams = Build-MailParameters `
            -Settings $Settings `
            -Html $html `
            -ExportedFiles $exportedFiles `
            -Counter $counter `
            -SystemErrors $systemErrors `
            -MatrixCount @($importedMatrix).Count `
            -ExistingMailParams $mailParams `
            -MailToDefaultsFile $mailToDefaultsFile `
            -LogFolder $LogFolder `
            -ScriptStartTime $scriptStartTime

        if ($systemErrors.Count -ne 0 -or $importedMatrix) {

            Send-MailSafe `
                -MailParams $mailParams `
                -SystemErrors ([ref]$systemErrors)

            if (Test-Path -LiteralPath $LogFolder -PathType Container) {
                Save-MailBodyToLog `
                    -MailParams $mailParams `
                    -LogFolder $LogFolder `
                    -SystemErrors ([ref]$systemErrors)
            }
        }
    }
    catch {
        #
        # ANY uncaught failure inside END block is severe
        #
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Unhandled error in END block: $_"
            }
        )
        Write-Warning "Unhandled fatal error: $_"
    }
    finally {
        #
        # Always remove drive if it exists
        #
        Remove-PSDrive MatrixFolderPath -ErrorAction Ignore

        #
        # Determine exit behavior
        #
        if ($systemErrors.Count -gt 0) {
            Write-Warning ('Found {0} system error(s).' -f $systemErrors.Count)
            $systemErrors |
            Sort-Object DateTime |
            ForEach-Object { Write-Warning $_.Message }

            Write-Warning 'Exit script with error code 1'
            exit 1
        }
        else {
            Write-Verbose 'Script finished successfully.'
        }
    }
}