#Requires -Version 7
#Requires -Modules ImportExcel, Toolbox.Remoting
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
    $ErrorActionPreference = 'stop'

    $eventLogData = [System.Collections.Generic.List[PSObject]]::new()
    $systemErrors = [System.Collections.Generic.List[PSObject]]::new()
    $scriptStartTime = Get-Date
        
    function ConvertTo-HtmlValueHC {
        if (-not $E.Value) {
            $null
        }
        elseif (($E.Value.Count -le 5) -and (-not ($E.Value -is [HashTable]))) {
            @"
        <ul>
            $(@($E.Value).ForEach({"<li>$_</li>"}))
        </ul>
"@
        }
        else {
            $fileName = "ID $($S.ID) - $($E.Type) - $($E.Name).txt".Split([IO.Path]::GetInvalidFileNameChars()) -join '_'

            $OutParams = @{
                LiteralPath = Join-Path -Path $I.File.LogFolder -ChildPath $fileName
                Encoding    = 'utf8'
                NoClobber   = $true
            }
            $E | ConvertTo-Json -Depth 100 | ForEach-Object {
                [System.Text.RegularExpressions.Regex]::Unescape($_)
            } | Out-File @OutParams
            @"
        <ul>
            <li><a href="$($OutParams.LiteralPath)">$("$($E.Value.Count) items")</a></li>
        </ul>
"@
        }
    }
    function Get-HTNLidTagProbTypeHC {
        [OutputType([String[]])]
        param (
            [Parameter(Mandatory)]
            [String]$Name
        )

        try {
            switch ($Name) {
                'FatalError' {
                    'probTypeError'
                    break
                }
                'Warning' {
                    'probTypeWarning'
                    break
                }
                'Information' {
                    'probTypeInfo'
                    break
                }
                default {
                    throw "Type '$_' is unknown"
                }
            }
        }
        catch {
            throw "Failed converting the HTML name '$Name' to a valid HTML ID tag: $_"
        }
    }

    try {
        $eventLogData.Add(
            [PSCustomObject]@{
                Message   = 'Script started'
                DateTime  = $scriptStartTime
                EntryType = 'Information'
                EventID   = '100'
            }
        )
        
        Get-Job | Remove-Job -Force

        #region Test path exists
        $scriptPathItem = @{}

        $ScriptPath.GetEnumerator().ForEach(
            {
                try {
                    $key = $_.Key
                    $value = $_.Value

                    $params = @{
                        Path        = $value
                        ErrorAction = 'Stop'
                    }
                    $scriptPathItem[$key] = (Get-Item @params).FullName
                }
                catch {
                    throw "ScriptPath.$key '$value' not found"
                }
            }
        )
        #endregion

        #region Import .json file
        Write-Verbose "Import .json file '$ConfigurationJsonFile'"

        $jsonFileItem = Get-Item -LiteralPath $ConfigurationJsonFile -ErrorAction Stop

        $jsonFileContent = Get-Content $jsonFileItem -Raw -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        Write-Verbose 'Test .json file properties'

        try {
            @(
                'MaxConcurrent', 'Matrix'
            ).where(
                { -not $jsonFileContent.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            @(
                'FolderPath', 'DefaultsFile'
            ).where(
                { -not $jsonFileContent.Matrix.$_ }
            ).foreach(
                { throw "Property 'Matrix.$_' not found" }
            )

            @(
                'Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer'
            ).foreach(
                {
                    if (-not $jsonFileContent.MaxConcurrent.$_) {
                        throw "Property 'MaxConcurrent.$_' not found" 
                    }
                    #region Test integer value
                    try {
                        [int]$jsonFileContent.MaxConcurrent.$_
                    }
                    catch {
                        throw "Property 'MaxConcurrent.$_' needs to be a number, the value '$($jsonFileContent.MaxConcurrent.$_)' is not supported."
                    }
                    #endregion
                }
            )

            #region Test boolean values
            foreach (
                $boolean in
                @(
                    'Archive'
                )
            ) {
                try {
                    $null = [Boolean]::Parse($jsonFileContent.Matrix.$boolean)
                }
                catch {
                    throw "Property 'Matrix.$boolean' is not a boolean value"
                }
            }
            
            try {
                $null = [Boolean]::Parse($jsonFileContent.Settings.SaveLogFiles.Detailed)
            }
            catch {
                throw "Property 'Settings.SaveLogFiles.Detailed' is not a boolean value"
            }
            #endregion

            #region Test array
            if (-not ($jsonFileContent.Matrix.ExcludedSamAccountName -is [Array])) {
                throw "Property 'Matrix.ExcludedSamAccountName' needs to be array"
            }
            #endregion
        }
        catch {
            throw "Input file '$ConfigurationJsonFile': $_"
        }
        #endregion

        $Matrix = $jsonFileContent.Matrix
        $Export = $jsonFileContent.Export
        $MaxConcurrent = $jsonFileContent.MaxConcurrent
        $ExcludedSamAccountName = $jsonFileContent.Matrix.ExcludedSamAccountName
        $DetailedLog = $jsonFileContent.Settings.SaveLogFiles.Detailed
        $LogFolder = $jsonFileContent.Settings.SaveLogFiles.Where.Folder

        #region Convert .json file
        Write-Verbose 'Convert .json file'

        #region Set PSSessionConfiguration
        $PSSessionConfiguration = $jsonFileContent.PSSessionConfiguration

        if (-not $PSSessionConfiguration) {
            $PSSessionConfiguration = 'PowerShell.7'
        }
        #endregion

        #region Test Export parameters
        if ($Export.FolderPath) {
            if (-not
                (Test-Path -LiteralPath $Export.FolderPath -PathType Container)
            ) {
                throw "Export folder '$($Export.FolderPath)' not found"
            }

            @(
                'AdObjects',
                'FormData',
                'AccessList',
                'GroupManagers',
                'ExcelOverview'
            ).Where( 
                { -not ($Export.FileName[$_]) }
            ).foreach( {
                    throw "Property 'Export.FileName.$_' is mandatory when the parameter Export.FolderPath is used."
                })
        }
        #endregion

        #region Create log folder
        try {
            $LogFolder = (New-Item -ItemType 'Directory' -Path $LogFolder -Force -EA Stop).FullName
        }
        catch {
            throw "Failed to create log folder '$LogFolder': $_"
        }
        #endregion

        #region Map share with Excel files
        if (-not (Test-Path -LiteralPath MatrixFolderPath:)) {
            $RetryCount = 0; $Completed = $false
            while (-not $Completed) {
                try {
                    $null = New-PSDrive -Name MatrixFolderPath -PSProvider FileSystem -Root $Matrix.FolderPath -EA Stop
                    $Completed = $true
                }
                catch {
                    if ($RetryCount -ge '240') {
                        throw "Drive mapping failed for '$($Matrix.FolderPath)': $_"
                    }
                    else {
                        Start-Sleep -Seconds 30
                        $RetryCount++
                        $Error.Clear()
                    }
                }
            }
        }

        $Matrix.FolderPath = Get-Item $Matrix.FolderPath -EA Stop
        #endregion

        #region Default settings file
        try {
            #region Get the defaults
            $DefaultsItem = Get-Item -LiteralPath $Matrix.DefaultsFile -EA Stop

            try {
                $DefaultsImport = Import-Excel -Path $DefaultsItem -Sheet 'Settings' -DataOnly -ErrorAction 'Stop'
            }
            catch {
                throw "worksheet 'Settings' not found*"
            }
            #endregion

            #region Verify mandatory column headers
            $propDefault = $DefaultsImport.ForEach( {
                    $_.PSObject.Properties.Name
                })

            @('MailTo', 'ADObjectName', 'Permission').Where( { $propDefault -notcontains $_ }).ForEach( {
                    throw "Column header '$_' not found. The column headers 'MailTo', 'ADObjectName' and 'Permission' are mandatory."
                })
            #endregion

            $DefaultAcl = Get-DefaultAclHC -Sheet $DefaultsImport

            #region Get MailTo
            $MailTo = $DefaultsImport.ForEach( {
                    $_.PSObject.Properties.Where( { ($_.Name -eq 'MailTo') -and ($_.Value) }).Foreach( {
                            $_.Value.ToString().Trim()
                        })
                })

            if (-not $MailTo) {
                throw "No mail addresses found under column header 'MailTo'"
            }
            #endregion
        }
        catch {
            throw "Defaults file '$($Matrix.DefaultsFile)' worksheet 'Settings': $_"
        }
        #endregion

        if ($Matrix.Archive) {
            try {
                $archivePath = Join-Path -Path $Matrix.FolderPath -ChildPath 'Archive'

                $ArchiveItem = (New-Item -ItemType 'Directory' -Path $archivePath -Force -EA Stop).FullName
            }
            catch {
                throw "Failed to create archive folder '$archivePath': $_"
            }
        }
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Input file '$ConfigurationJsonFile': $_"
            }
        )

        Write-Warning $systemErrors[-1].Message

        return
    }
}

process {
    try {
        $ID = 0

        $getParams = @{
            Path        = 'MatrixFolderPath:\*'
            Include     = '*.xlsx'
            File        = $true
            ErrorAction = 'Stop'
        }

        [Array]$importedMatrix = foreach (
            $matrixFile in
            @(Get-ChildItem @getParams).Where(
                { $_.FullName -ne $DefaultsItem.FullName })
        ) {
            try {
                Write-Verbose "Matrix file '$matrixFile'"

                $Obj = [PSCustomObject]@{
                    File        = @{
                        Item         = $matrixFile
                        SaveFullName = $matrixFile.FullName
                        ExcelInfo    = $null
                        LogFolder    = $null
                        Check        = @()
                    }
                    Settings    = @()
                    Permissions = @{
                        Import = @()
                        Check  = @()
                    }
                    FormData    = @{
                        Import = $null
                        Check  = @()
                    }
                }

                #region Create log folder
                try {
                    $logFolderPath = Join-Path -Path $LogFolder -ChildPath (
                        '{0:00}-{1:00}-{2:00} {3:00}{4:00} ({5}) - {6}' -f $scriptStartTime.Year, $scriptStartTime.Month,
                        $scriptStartTime.Day, $scriptStartTime.Hour, $scriptStartTime.Minute, $scriptStartTime.DayOfWeek, $matrixFile.BaseName)

                    $Obj.File.LogFolder = (New-Item -ItemType 'Directory' -Path $logFolderPath -Force -EA Stop).FullName
                }
                catch {
                    throw "Failed to create log folder '$logFolderPath': $_"
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
                $Obj.File.ExcelInfo = Get-ExcelWorkbookInfo -Path $matrixFile.FullName -ErrorAction 'Stop'

                Write-Verbose "File '$($matrixFile.Name)': LastModifiedBy '$($Obj.File.ExcelInfo.LastModifiedBy)' LastModifiedDate '$($Obj.File.ExcelInfo.Modified.ToString('dd/MM/yyyy HH:mm:ss'))'"
                #endregion

                #region Import sheets Settings, Permissions, FormData
                try {
                    #region Import sheet Settings
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "File '$($matrixFile.Name)': Import worksheet 'Settings'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '2'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $ImportParams = @{
                        Path        = $matrixFile.FullName
                        DataOnly    = $true
                        ErrorAction = 'Stop'
                    }
                    $Settings = @(
                        Import-Excel @ImportParams -Sheet 'Settings'
                    ).Where(
                        { $_.Status -eq 'Enabled' }
                    )
                    #endregion

                    if ($Settings) {
                        foreach ($S in $Settings) {
                            $ID++

                            $Obj.Settings += [PSCustomObject]@{
                                ID        = $ID
                                Import    = Format-SettingStringsHC -Settings $S
                                Check     = @()
                                Matrix    = @()
                                AdObjects = @{}
                                JobTime   = @{}
                            }
                        }

                        #region Import sheet Permissions
                        $eventLogData.Add(
                            [PSCustomObject]@{
                                Message   = "File '$($matrixFile.Name)': Import worksheet 'Permissions'"
                                DateTime  = Get-Date
                                EntryType = 'Information'
                                EventID   = '2'
                            }
                        )
                        Write-Verbose $eventLogData[-1].Message

                        $Obj.Permissions.Import = @(
                            Import-Excel @ImportParams -Sheet 'Permissions' -NoHeader |
                            Format-PermissionsStringsHC
                        )
                        #endregion

                        #region Import sheet FormData
                        if ($Export.FolderPath) {
                            try {
                                $eventLogData.Add(
                                    [PSCustomObject]@{
                                        Message   = "File '$($matrixFile.Name)': Import worksheet 'FormData'"
                                        DateTime  = Get-Date
                                        EntryType = 'Information'
                                        EventID   = '2'
                                    }
                                )
                                Write-Verbose $eventLogData[-1].Message
                        
                                $formData = Import-Excel @ImportParams -Sheet 'FormData' -ErrorVariable importFail

                                $Obj.FormData.Check += Test-FormDataHC $formData

                                if (-not $Obj.FormData.Check) {
                                    $Obj.FormData.Import = $formData
                                }
                            }
                            catch {
                                $Obj.File.Check += [PSCustomObject]@{
                                    Type        = 'FatalError'
                                    Name        = "Worksheet 'FormData' not found"
                                    Description = "When the argument 'Export.FolderPath' is used the Excel file needs to have a worksheet 'FormData'."
                                    Value       = @($_)
                                }
                                # remove multiple errors from Import-Excel
                                $importFail | ForEach-Object {
                                    $Error.Remove($_)
                                }
                            }
                        }
                        #endregion
                    }
                    else {
                        $Obj.File.Check += [PSCustomObject]@{
                            Type        = 'Warning'
                            Name        = 'Matrix disabled'
                            Description = 'Every Excel file needs at least one enabled matrix.'
                            Value       = "The worksheet 'Settings' does not contain a row with 'Status' set to 'Enabled'."
                        }

                        $eventLogData.Add(
                            [PSCustomObject]@{
                                Message   = "File '$($matrixFile.Name)': No lines found with status 'Enabled' in the worksheet 'Settings'"
                                DateTime  = Get-Date
                                EntryType = 'Information'
                                EventID   = '2'
                            }
                        )
                        Write-Warning $eventLogData[-1].Message
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
                            throw "Failed importing the Excel file '$($matrixFile.FullName)': $_"
                        }
                    }
                    $Obj.File.Check += [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Excel file incorrect'
                        Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                        Value       = $errorMessage
                    }

                    try { $Error.RemoveRange(0, 2) }
                    catch { throw 'Import-Excel throws 2 errors normally' }
                }
                #endregion

                if ($Matrix.Archive) {
                    try {
                        $eventLogData.Add(
                            [PSCustomObject]@{
                                Message   = "File '$($matrixFile.Name)': Move file to archive folder '$($ArchiveItem.FullName)'"
                                DateTime  = Get-Date
                                EntryType = 'Information'
                                EventID   = '2'
                            }
                        )
                        Write-Verbose $eventLogData[-1].Message

                        Move-Item -LiteralPath $matrixFile -Destination $ArchiveItem -Force -EA Stop
                    }
                    catch {
                        $Obj.File.Check += [PSCustomObject]@{
                            Type        = 'Warning'
                            Name        = 'Archiving failed'
                            Description = "When the '-Archive' switch is used the file is moved to the archive folder.In case a file is still in use, the move operation might fail."
                            Value       = @($_)
                        }

                        $Error.RemoveAt(0)
                    }
                }

                $Obj
            }
            catch {
                $systemErrors.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "File '$($matrixFile.Name)': $_"
                    }
                )

                Write-Warning $systemErrors[-1].Message
            }
        }

        if ($importedMatrix) {
            #region Build FormData for Export folder
            foreach ($I in ($importedMatrix.Where( { $_.FormData.Import }))) {
                try {
                    $property = @{}

                    #region Convert MatrixResponsible to UserPrincipalName
                    $params = @{
                        Name                  = $I.FormData.Import.MatrixResponsible.Split(',').Trim()
                        ExcludeSamAccountName = $ExcludedSamAccountName
                    }
                    $result = Get-AdUserPrincipalNameHC @params

                    $property.MatrixResponsible = $result.userPrincipalName -join ','

                    if ($result.notFound) {
                        $I.FormData.Check += [PSCustomObject]@{
                            Type        = 'Warning'
                            Name        = 'AD object not found'
                            Description = "The email address or SamAccountName is not found in the active directory. Multiple entries are supported with the comma ',' separator."
                            Value       = $result.notFound
                        }
                    }
                    #endregion

                    #region Add MatrixFilePath and MatrixFileName
                    $property.MatrixFilePath = if ($Matrix.Archive) {
                        Join-Path $ArchiveItem $I.File.Item.Name
                    }
                    else {
                        $I.File.Item.FullName
                    }

                    $property.MatrixFileName = $I.File.Item.BaseName
                    #endregion

                    $I.FormData.Import | Add-Member -NotePropertyMembers $property -Force
                }
                catch {
                    $I.FormData.Check += [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Failed adding property'
                        Description = "The worksheet 'FormData' could not be updated correctly."
                        Value       = @($_)
                    }
                }
            }
            #endregion

            #region Build the matrix and check for incorrect input
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Build the matrix and check for incorrect input'
                    DateTime  = Get-Date
                    EntryType = 'Information'
                    EventID   = '2'
                }
            )
            Write-Verbose $eventLogData[-1].Message

            foreach (
                $I in
                $importedMatrix.Where(
                    {
                        ($_.File.Check.Type -notcontains 'FatalError' ) -and
                        ($_.Settings)
                    }
                )
            ) {
                try {
                    Write-Verbose 'Test matrix permissions'

                    $I.Permissions.Check += Test-MatrixPermissionsHC -Permissions $I.Permissions.Import

                    if ($I.Permissions.Check.Type -notcontains 'FatalError') {
                        foreach ($S in $I.Settings) {
                            $S.Check += Test-MatrixSettingHC -Setting $S.Import

                            #region Create AD object names
                            Write-Verbose 'Create AD object names'

                            $params = @{
                                Begin         = $S.Import.GroupName
                                Middle        = $S.Import.SiteCode
                                ColumnHeaders = $I.Permissions.Import |
                                Select-Object -First 3
                            }
                            $adObjects = ConvertTo-MatrixADNamesHC @params

                            Write-Verbose 'Test AD objects'

                            $S.Check += Test-AdObjectsHC $adObjects
                            #endregion

                            #region Create matrix for each settings line
                            if ($S.Check.Type -notcontains 'FatalError') {
                                Write-Verbose 'Create matrix for each settings line'

                                $S.AdObjects = $adObjects

                                $params = @{
                                    NonHeaderRows = $I.Permissions.Import |
                                    Select-Object -Skip 3
                                    ADObjects     = $adObjects
                                }
                                $S.Matrix += ConvertTo-MatrixAclHC @params
                            }
                            #endregion
                        }
                    }
                }
                catch {
                    $I.File.Check += [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Unknown error'
                        Description = 'While checking the input and generating the matrix an error was reported.'
                        Value       = $_
                    }
                    $Error.RemoveAt(0)
                }
            }
            #endregion

            #region Test duplicate ComputerName/Path combination
            Write-Verbose 'Check duplicate ComputerName/Path combination'

            (
                @($importedMatrix.Settings | Group-Object @{
                        Expression = {
                            $_.Import.ComputerName + ' - ' + $_.Import.Path }
                    }
                ).Where( { $_.Count -ge 2 })
            ).Group.Foreach(
                {
                    $_.Check += [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Duplicate ComputerName/Path combination'
                        Description = "Every 'ComputerName' combined with a 'Path' needs to be unique over all the 'Settings' worksheets found in all the active matrix files."
                        Value       = @{
                            $_.Import.ComputerName = $_.Import.Path
                        }
                    }
                }
            )
            #endregion

            #region Test expanded matrix and get AD object details
            Write-Verbose 'Check expanded matrix'

            $AdObjects = $importedMatrix.Settings.Matrix.ACL.Keys

            if ($AdObjects.count -ne 0) {
                Write-Verbose 'Get AD object details'
                $params = @{
                    ADObjectName = $AdObjects | Sort-Object -Unique
                    Type         = 'SamAccountName'
                }
                $ADObjectDetails = @(Get-ADObjectDetailHC @params)

                @($importedMatrix.Settings).Where( { $_.Matrix }).Foreach(
                    {
                        Write-Verbose "Test expanded matrix for Settings row ComputerName '$($_.Import.ComputerName)' Path '$($_.Import.Path)' SiteName '$($_.Import.SiteName)' SiteCode '$($_.Import.SiteCode)' GroupName '$($_.Import.GroupName)'"

                        $params = @{
                            Matrix                 = $_.Matrix
                            ADObject               = $ADObjectDetails
                            DefaultAcl             = $DefaultAcl
                            ExcludedSamAccountName = $ExcludedSamAccountName
                        }
                        $_.Check += Test-ExpandedMatrixHC @params
                    }
                )
            }
            #endregion

            #region Get AD object details for group managers
            if (
                $groupManagers = $ADObjectDetails.ADObject.ManagedBy |
                Sort-Object -Unique
            ) {
                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = "Retrieve AD object details for $($groupManagers.Count) group managers"
                        DateTime  = Get-Date
                        EntryType = 'Information'
                        EventID   = '2'
                    }
                )
                Write-Verbose $eventLogData[-1].Message

                $params = @{
                    ADObjectName = $groupManagers
                    Type         = 'DistinguishedName'
                }
                $groupManagersAdDetails = Get-ADObjectDetailHC @params
            }
            #endregion

            #region Remove group members that are in the ExcludedSamAccountName
            if ($ExcludedSamAccountName) {
                foreach ($adObject in $ADObjectDetails) {
                    $adObject.adGroupMember = $adObject.adGroupMember |
                    Where-Object {
                        $ExcludedSamAccountName -notcontains $_.SamAccountName
                    }
                }
                foreach ($adObject in $groupManagersAdDetails) {
                    $adObject.adGroupMember = $adObject.adGroupMember |
                    Where-Object {
                        $ExcludedSamAccountName -notcontains $_.SamAccountName
                    }
                }
            }
            #endregion

            #region Test server requirements
            if (
                $executableMatrix = @(
                    Get-ExecutableMatrixHC -From $importedMatrix)
            ) {
                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = 'Test server requirements'
                        DateTime  = Get-Date
                        EntryType = 'Information'
                        EventID   = '2'
                    }
                )
                Write-Verbose $eventLogData[-1].Message

                $scriptBlock = {
                    try {
                        #region Declare variables for parallel execution
                        if (-not $MaxConcurrentComputers) {
                            $scriptPathItem = $using:scriptPathItem
                            $PSSessionConfiguration = $using:PSSessionConfiguration
                            $eventLogData = $using:eventLogData
                        }
                        #endregion

                        $matrix = $_.Group
                        $computerName = $_.Name

                        $params = @{
                            FilePath          = $scriptPathItem.TestRequirementsFile
                            ArgumentList      = $matrix.Import.Path, $true
                            ConfigurationName = $PSSessionConfiguration
                            ComputerName      = $computerName
                            ErrorAction       = 'Stop'
                        }
                        if ($result = Invoke-Command @params) {
                            $matrix | ForEach-Object { $_.Check += $result }
                        }
                    }
                    catch {
                        $problem = [PSCustomObject]@{
                            Type        = 'FatalError'
                            Name        = 'Computer requirements'
                            Value       = $_
                            Description = "Failed checking the computer for the minimal requirements with the 'Test requirements' script."
                        }
                        $Error.RemoveAt(0)
                        $matrix | ForEach-Object { $_.Check += $problem }
                    }
                }

                #region Run code serial or parallel
                $foreachParams = if ($MaxConcurrent.Computers -eq 1) {
                    @{
                        Process = $scriptBlock
                    }
                }
                else {
                    @{
                        Parallel      = $scriptBlock
                        ThrottleLimit = $MaxConcurrent.Computers
                    }
                }
                #endregion

                $executableMatrix |
                Group-Object -Property { $_.Import.ComputerName } |
                ForEach-Object @foreachParams
            }
            #endregion

            #region Set permissions
            if (
                $executableMatrix = @(
                    Get-ExecutableMatrixHC -From $importedMatrix)
            ) {
                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = "Start 'Set permissions' script for '$($executableMatrix.Count)' matrix"
                        DateTime  = Get-Date
                        EntryType = 'Information'
                        EventID   = '2'
                    }
                )
                Write-Verbose $eventLogData[-1].Message

                #region Add default permissions
                <#
                    In case of conflict the acl in the matrix will win
                    over the acl in the Matrix.DefaultsFile.
                #>
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

                $matrixes = $null

                $outerScriptBlock = {
                    # $VerbosePreference = 'Continue'

                    $matrixes = $_.Group

                    #region Declare variables for parallel execution
                    if (-not $MaxConcurrent) {
                        $MaxConcurrent = $using:MaxConcurrent
                        $scriptPathItem = $using:scriptPathItem
                        $PSSessionConfiguration = $using:PSSessionConfiguration
                        $DetailedLog = $using:DetailedLog
                        $eventLogData = $using:eventLogData
                    }
                    #endregion

                    $innerScriptBlock = {
                        try {
                            # $VerbosePreference = 'Continue'

                            $matrix = $_

                            #region Declare variables for parallel execution
                            if (-not $MaxConcurrent) {
                                $MaxConcurrent = $using:MaxConcurrent
                                $scriptPathItem = $using:scriptPathItem
                                $PSSessionConfiguration = $using:PSSessionConfiguration
                                $DetailedLog = $using:DetailedLog
                                $eventLogData = $using:eventLogData
                            }
                            #endregion

                            $matrix.JobTime.Start = Get-Date

                            $params = @{
                                FilePath          = $scriptPathItem.SetPermissionFile
                                ArgumentList      = $matrix.Import.Path, $matrix.Import.Action, $matrix.Matrix, $MaxConcurrent.FoldersPerMatrix, $DetailedLog
                                ConfigurationName = $PSSessionConfiguration
                                ComputerName      = $matrix.Import.ComputerName
                                ErrorAction       = 'Stop'
                            }
                            if ($result = Invoke-Command @params) {
                                $matrix.Check += $result
                            }
                        }
                        catch {
                            $problem = [PSCustomObject]@{
                                Type        = 'FatalError'
                                Name        = 'Set permissions'
                                Value       = $_
                                Description = "Failed applying action '$($matrix.Import.Action)' with the 'Set permissions' script."
                            }
                            $Error.RemoveAt(0)
                            $matrix.Check += $problem
                        }
                        finally {
                            $matrix.JobTime.End = Get-Date
                            $matrix.JobTime.Duration = New-TimeSpan -Start $matrix.JobTime.Start -End $matrix.JobTime.End
                        }
                    }

                    $innerForeachParams = if (
                        $MaxConcurrent.JobsPerRemoteComputer -gt 1
                    ) {
                        @{
                            Parallel      = $innerScriptBlock
                            ThrottleLimit = $MaxConcurrent.JobsPerRemoteComputer
                        }
                    }
                    else {
                        @{
                            Process = $innerScriptBlock
                        }    
                    }

                    $matrixes | ForEach-Object @innerForeachParams
                }

                $foreachParams = if ($MaxConcurrent.Computers -gt 1) {
                    @{
                        Parallel      = $outerScriptBlock
                        ThrottleLimit = $MaxConcurrent.Computers
                    }
                }
                else {
                    @{
                        Process = $outerScriptBlock
                    }
                }

                $executableMatrix |
                Group-Object -Property { $_.Import.ComputerName } |
                ForEach-Object @foreachParams -Verbose
            }
            #endregion
        }
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = $_
            }
        )

        Write-Warning $systemErrors[-1].Message
    }
    finally {
        if ($psSessions.Values.Session) {
            # Only close PS Sessions and not the WinPSCompatSession
            # used by Write-EventLog
            # https://github.com/PowerShell/PowerShell/issues/24227
            $psSessions.Values.Session | Remove-PSSession -EA Ignore
        }
    }
}

end {
    try {
        $settings = $jsonFileContent.Settings

        $scriptName = $settings.ScriptName
        $saveInEventLog = $settings.SaveInEventLog
        $sendMail = $settings.SendMail
        $saveLogFiles = $settings.SaveLogFiles

        $allLogFilePaths = @()
        $baseLogName = $null
        $logFolderPath = $null

        #region Get script name
        if (-not $scriptName) {
            Write-Warning "No 'Settings.ScriptName' found in import file."
            $scriptName = 'Default script name'
        }
        #endregion

        $matrixLogFile = Join-Path -Path $LogFolder -ChildPath (
            '{0:00}-{1:00}-{2:00} {3:00}{4:00} ({5})' -f
            $scriptStartTime.Year, $scriptStartTime.Month, $scriptStartTime.Day,
            $scriptStartTime.Hour, $scriptStartTime.Minute, $scriptStartTime.DayOfWeek
        )

        if ($importedMatrix) {
            $groupManagersSheet = @()
            $accessListSheet = @()

            #region Export to matrix Excel log file
            foreach ($I in $importedMatrix) {
                #region Get unique SamAccountNames for all matrix in Settings
                $matrixSamAccountNames = $i.Settings.AdObjects.Values.SamAccountName |
                Select-Object -Property @{
                    Name       = 'name'
                    Expression = { "$($_)".Trim() }
                } -Unique |
                Select-Object -ExpandProperty name

                Write-Verbose "Matrix '$($i.File.Item.Name)' has '$($matrixSamAccountNames.count)' unique SamAccountNames"
                #endregion

                #region Create Excel worksheet 'AccessList'
                $accessListToExport = foreach ($S in $matrixSamAccountNames) {
                    $adData = $ADObjectDetails |
                    Where-Object { $S -eq $_.samAccountName }

                    if (-not $adData.adObject) {
                        $eventLogData.Add(
                            [PSCustomObject]@{
                                Message   = "Matrix '$($i.File.Item.Name)' SamAccountName '$s' not found in AD"
                                DateTime  = Get-Date
                                EntryType = 'Information'
                                EventID   = '2'
                            }
                        )
                        Write-Warning $eventLogData[-1].Message
                    }
                    elseif (-not $adData.adGroupMember) {
                        $adData | Select-Object -Property SamAccountName,
                        @{Name = 'Name'; Expression = { $_.adObject.Name } },
                        @{Name = 'Type'; Expression = { $_.adObject.ObjectClass } },
                        MemberName, MemberSamAccountName
                    }
                    else {
                        $adData.adGroupMember | Select-Object -Property @{
                            Name       = 'SamAccountName'
                            Expression = { $S }
                        },
                        @{Name = 'Name'; Expression = { $adData.adObject.Name } },
                        @{Name = 'Type'; Expression = { $adData.adObject.ObjectClass } },
                        @{Name = 'MemberName'; Expression = { $_.Name } },
                        @{Name = 'MemberSamAccountName'; Expression = { $_.SamAccountName } }
                    }
                }
                #endregion

                #region Create Excel worksheet 'GroupManagers'
                $groupManagersToExport = foreach ($S in $matrixSamAccountNames) {
                    $adData = (
                        $ADObjectDetails | Where-Object {
                            ($S -eq $_.samAccountName) -and
                            ($_.adObject.ObjectClass -eq 'group')
                        }
                    )
                    if ($adData) {
                        $groupManager = $groupManagersAdDetails | Where-Object {
                            $_.DistinguishedName -eq $adData.adObject.ManagedBy
                        }

                        if (-not $groupManager) {
                            [PSCustomObject]@{
                                GroupName         = $adData.adObject.Name
                                ManagerName       = $null
                                ManagerType       = $null
                                ManagerMemberName = $null
                            }
                        }
                        elseif (-not $groupManager.adGroupMember) {
                            [PSCustomObject]@{
                                GroupName         = $adData.adObject.Name
                                ManagerName       = $groupManager.adObject.Name
                                ManagerType       = $groupManager.adObject.ObjectClass
                                ManagerMemberName = $null
                            }
                        }
                        else {
                            foreach ($user in $groupManager.adGroupMember) {
                                [PSCustomObject]@{
                                    GroupName         = $adData.adObject.Name
                                    ManagerName       = $groupManager.adObject.Name
                                    ManagerType       = $groupManager.adObject.ObjectClass
                                    ManagerMemberName = $user.Name
                                }
                            }
                        }
                    }
                }
                #endregion

                if ($accessListToExport) {
                    #region Export to Excel worksheet 'AccessList'
                    $excelParams = @{
                        Path               = $I.File.SaveFullName
                        AutoSize           = $true
                        WorksheetName      = 'AccessList'
                        TableName          = 'AccessList'
                        FreezeTopRow       = $true
                        NoNumberConversion = '*'
                        ClearSheet         = $true
                    }

                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export $($accessListToExport.Count) AD objects to Excel file '$($excelParams.Path)' worksheet '$($excelParams.WorksheetName)'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message
                        
                    $accessListToExport | Export-Excel @excelParams
                    #endregion

                    #region Create 'AccessList' to export
                    if ($Export.FolderPath) {
                        $accessListSheet += $accessListToExport |
                        Select-Object @{
                            Name       = 'MatrixFileName'
                            Expression = { $I.File.Item.BaseName }
                        }, *
                    }
                    #endregion

                    if ($groupManagersToExport) {
                        #region Export to Excel worksheet 'GroupManagers'
                        $excelParams.WorksheetName = $excelParams.TableName = 'GroupManagers'

                        $eventLogData.Add(
                            [PSCustomObject]@{
                                Message   = "Export $($groupManagersToExport.Count) AD objects to Excel file '$($excelParams.Path)' worksheet '$($excelParams.WorksheetName)'"
                                DateTime  = Get-Date
                                EntryType = 'Information'
                                EventID   = '1'
                            }
                        )
                        Write-Verbose $eventLogData[-1].Message

                        $groupManagersToExport | Export-Excel @excelParams
                        #endregion

                        #region Create 'GroupManagers' to export
                        if ($Export.FolderPath) {
                            $groupManagersSheet += $groupManagersToExport |
                            Select-Object @{
                                Name       = 'MatrixFileName'
                                Expression = { $I.File.Item.BaseName }
                            }, *
                        }
                        #endregion
                    }
                }
            }
            #endregion

            #region Export data to .XLSX and .CSV files
            $formDataSheet = @()
            $adObjectNamesSheet = @()
            $dataToExport = @{}

            if (
                $Export.FolderPath -and
                ($importedMatrix.FormData.Check.Type -notcontains 'FatalError')
            ) {
                #region Create data to export hashtable
                foreach ($property in $Export.FileName) {
                    $dataToExport[$property.Name] = @{
                        LogFilePath    = "$matrixLogFile - Export - $($property.Value)"
                        ExportFilePath = Join-Path $Export.FolderPath $property.Value
                        Data           = @()
                    }
                }
                #endregion

                #region Remove old exported log files
                $dataToExport.GetEnumerator() | ForEach-Object {
                    $fileToRemove = $_.Value.LogFilePath
                    
                    if (Test-Path $fileToRemove) {
                        Write-Verbose "Remove file '$fileToRemove'"
                        Remove-Item -Path $fileToRemove -ErrorAction Ignore
                    }
                }
                #endregion

                #region Create AdObjectNames and FormData to export
                foreach ($I in $importedMatrix) {
                    $adObjects = foreach (
                        $S in
                        $I.Settings.Where( { $_.AdObjects.Count -ne 0 })
                    ) {
                        foreach ($A in ($S.AdObjects.GetEnumerator())) {
                            [PSCustomObject]@{
                                MatrixFileName	= $I.File.Item.BaseName
                                SamAccountName = $A.Value.SamAccountName
                                GroupName      = $A.Value.Converted.Begin
                                SiteCode       = $A.Value.Converted.Middle
                                Name           = $A.Value.Converted.End
                            }
                        }
                    }

                    if ($adObjects) {
                        $formDataSheet += $I.FormData.Import

                        $adObjectNamesSheet += $adObjects |
                        Group-Object SamAccountName |
                        ForEach-Object { $_.Group[0] }
                    }
                }
                #endregion

                #region Create parameters
                $ExportParams = @{
                    Path         = "$matrixLogFile - Cherwell - $($Export.FileName.ExcelOverview)"
                    AutoSize     = $true
                    FreezeTopRow = $true
                }

                $exportCsvAdParams = @{
                    literalPath       = Join-Path $Export.FolderPath $Export.FileName.AdObjects
                    Encoding          = 'utf8'
                    NoTypeInformation = $true
                }

                $exportCsvFormParams = @{
                    literalPath       = Join-Path $Export.FolderPath $Export.FileName.FormData
                    Encoding          = 'utf8'
                    NoTypeInformation = $true
                }

                $exportCsvGroupManagersParams = @{
                    literalPath       = Join-Path $Export.FolderPath $Export.FileName.GroupManagers
                    Encoding          = 'utf8'
                    NoTypeInformation = $true
                }

                $exportCsvAccessListParams = @{
                    literalPath       = Join-Path $Export.FolderPath $Export.FileName.AccessList
                    Encoding          = 'utf8'
                    NoTypeInformation = $true
                }
                #endregion

                if ($AdObjectNamesSheet) {
                    #region Export AD object names to .XLSX file
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export $($AdObjectNamesSheet.Count) AD object names to '$($ExportParams.Path)'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $AdObjectNamesSheet |
                    Export-Excel @ExportParams -WorksheetName 'AdObjectNames' -TableName 'AdObjectNames'
                    #endregion

                    #region Export AD object names to a .CSV file
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export $($AdObjectNamesSheet.Count) AD object names to '$($exportCsvAdParams.literalPath)'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $adObjectNamesSheet | Export-Csv @exportCsvAdParams
                    #endregion

                    #region Copy csv file to log folder
                    $copyParams = @{
                        LiteralPath = $exportCsvAdParams.literalPath
                        Destination = "$matrixLogFile - Cherwell - $($Export.FileName.AdObjects)"
                    }
                    Copy-Item @copyParams
                    #endregion
                }

                if ($groupManagersSheet) {
                    #region Export group managers to .XLSX file
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export $($groupManagersSheet.Count) group managers to '$($ExportParams.Path)'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $groupManagersSheet |
                    Export-Excel @ExportParams -WorksheetName 'GroupManagers' -TableName 'GroupManagers'
                    #endregion

                    #region Export group managers to a .CSV file
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export $($groupManagersSheet.Count) group managers to '$($exportCsvGroupManagersParams.literalPath)'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $groupManagersSheet |
                    Export-Csv @exportCsvGroupManagersParams
                    #endregion

                    #region Copy csv file to log folder
                    $copyParams = @{
                        LiteralPath = $exportCsvGroupManagersParams.literalPath
                        Destination = "$matrixLogFile - Cherwell - $($Export.FileName.GroupManagers)"
                    }
                    Copy-Item @copyParams
                    #endregion
                }

                if ($accessListSheet) {
                    #region Export access list to .XLSX file
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export $($accessListSheet.Count) access list to '$($ExportParams.Path)'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $accessListSheet |
                    Export-Excel @ExportParams -WorksheetName 'AccessList' -TableName 'AccessList'
                    #endregion

                    #region Export access list to .CSV file
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export access list to '$($exportCsvAccessListParams.literalPath)'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $accessListSheet |
                    Export-Csv @exportCsvAccessListParams
                    #endregion

                    #region Copy csv file to log folder
                    $copyParams = @{
                        LiteralPath = $exportCsvAccessListParams.literalPath
                        Destination = "$matrixLogFile - Cherwell - $($Export.FileName.AccessList)"
                    }
                    Copy-Item @copyParams
                    #endregion
                }

                if ($formDataSheet) {
                    #region Export FormData to .XLSX file
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export FormData to '$($ExportParams.Path)'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $formDataSheet |
                    Export-Excel @ExportParams -WorksheetName 'FormData' -TableName 'FormData'
                    #endregion

                    #region Export FormData to .CSV file
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export FormData to '$($exportCsvFormParams.literalPath)'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $formDataSheet | Export-Csv @exportCsvFormParams
                    #endregion

                    #region Export FormData to an HTML file
                    $htmlFileContent = @(
                        @'
<style>
  body {
    background-color: #f0f0f0;
    color: #004e2b;
    font-family: Arial, sans-serif;
    padding: 20px;
  }

  a {
    color: #004e2b;
    text-decoration: none;
  }
  a:hover {
    color: #00dd39;
    text-decoration: underline;
  }

  h1 {
    border-bottom: 2px solid #004e2b;
    padding-bottom: 10px;
    margin-bottom: 25px;
    color: #004e2b;
    text-transform: uppercase;
    font-size: 1.8em;
  }

  table {
    width: 100%;
    max-width: 1200px;
    margin: 20px 0;
    border-collapse: separate;
    border-spacing: 0;
    box-shadow: 0 6px 15px rgba(0, 0, 0, 0.2);
    background-color: #ffffff;
    border-radius: 8px;
    overflow: hidden;
    table-layout: auto;
    border: none;
  }

  table th {
    background-color: #004e2b;
    color: #ffffff;
    text-align: left;
    padding: 15px 20px;
    font-weight: bold;
    text-transform: uppercase;
    border: none;
    font-size: 0.9em;
  }

  table thead tr:first-child th:first-child {
    border-top-left-radius: 8px;
  }
  table thead tr:first-child th:last-child {
    border-top-right-radius: 8px;
  }

  table th:nth-child(3) {
    text-align: left;
    word-break: normal;
  }

  table td {
    text-align: center;
    padding: 10px 15px;
    border: none;
    border-bottom: 1px solid #e0e0e0;
    vertical-align: middle;
    color: #004e2b;
  }

  table tbody tr:last-child td {
    border-bottom: none;
  }

  table td:nth-child(3) {
    text-align: left;
    white-space: nowrap;
    word-break: normal;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  table td:nth-child(4) {
    text-align: left;
    white-space: nowrap;
    word-break: normal;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  table td:nth-child(5) {
    text-align: left;
    white-space: nowrap;
    word-break: normal;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  table tbody tr:nth-child(even) {
    background-color: #f8f8f8b7;
  }
  table tbody tr:nth-child(odd) {
    background-color: #ffffff;
  }

  table tbody tr:hover {
    background-color: #c2ebcf;
    color: #004e2b;
  }

  table tbody tr td a {
    display: block;
    width: 100%;
    height: 100%;
    color: #004e2b;
  }
  table td:last-child a {
    display: inline;
    color: #004e2b;
  }

  table tbody tr:hover td a {
    color: #004e2b;
  }
</style>
'@,
                        '<h1>Matrix files overview</h1>'
                    )

                    $htmlMatrixTableRows = '<tr>
                        <th>Category</th>
                        <th>Subcategory</th>
                        <th>Folder</th>
                        <th>Link to the matrix</th>
                        <th>Responsible</th>
                    </tr>
                    '

                    $htmlMatrixTableRows += $formDataSheet | 
                    Sort-Object -Property 'MatrixCategoryName', 'MatrixSubCategoryName', 'MatrixFolderDisplayName' | 
                    ForEach-Object {
                        $emailsMatrixResponsible = foreach (
                            $email in
                            $_.MatrixResponsible -split ','
                        ) {
                            "<a href=`"mailto:$email`">$email</a>"
                        }

                        "<tr>
                                <td>$($_.MatrixCategoryName)</td>
                                <td>$($_.MatrixSubCategoryName)</td>
                                <td><a href=`"$($_.MatrixFolderDisplayName)`">$($_.MatrixFolderDisplayName)</a></td>
                                <td><a href=`"$($_.MatrixFilePath)`">$($_.MatrixFileName)</a></td>
                                <td>$emailsMatrixResponsible</td>
                            </tr>"
                    }

                    $htmlFileContent += "<table>$htmlMatrixTableRows</table>"

                    $joinParams = @{
                        Path      = $Export.FolderPath
                        ChildPath = $Export.FileName.ExcelOverview.Replace('.xlsx', '.html')
                    }
                    $htmlFilePath = Join-Path @joinParams

                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = "Export FormData to '$htmlFilePath'"
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '1'
                        }
                    )
                    Write-Verbose $eventLogData[-1].Message

                    $htmlFileContent | Out-File -LiteralPath $htmlFilePath -Encoding utf8 -Force
                    #endregion

                    #region Copy csv file to log folder
                    $copyParams = @{
                        LiteralPath = $exportCsvFormParams.literalPath
                        Destination = "$matrixLogFile - Cherwell - $($Export.FileName.FormData)"
                    }
                    Copy-Item @copyParams
                    #endregion

                    #region Start ServiceNow FormData upload
                    try {
                        $params = @{
                            ServiceNowCredentialsFilePath = $jsonFileContent.ServiceNow.CredentialsFilePath
                            Environment                   = $jsonFileContent.ServiceNow.Environment
                            FormDataFile                  = $exportCsvFormParams.literalPath
                            TableName                     = $jsonFileContent.ServiceNow.TableName
                        }
                        & $scriptPathItem.UpdateServiceNow @params
                    }
                    catch {
                        $systemErrors.Add(
                            [PSCustomObject]@{
                                DateTime = Get-Date
                                Message  = "Failed executing script '$($scriptPathItem.UpdateServiceNow.FullName)': $_"
                            }
                        )

                        Write-Warning $systemErrors[-1].Message
                    }
                    #endregion
                }

                if ($adObjectNamesSheet -or $formDataSheet -or
                    $accessListSheet -or $groupManagersSheet) {
                    #region Copy Excel file from log folder to Export folder
                    $copyParams = @{
                        LiteralPath = $ExportParams.Path
                        Destination = Join-Path $Export.FolderPath $Export.FileName.ExcelOverview
                    }
                    Copy-Item @copyParams
                    #endregion
                }
            }
            #endregion

            #region HTML <style> for Mail and Settings
            Write-Verbose 'Format HTML'

            $htmlStyle = @'
<style>
    a {
        color: black;
        text-decoration: underline;
    }
    a:hover {
        color: blue;
    }

    #overviewTable {
        border-collapse: collapse;
        border: 1px solid Black;
        table-layout: fixed;
    }

    #overviewTable th {
        font-weight: normal;
        text-align: left;
    }
    #overviewTable td {
        text-align: center;
    }

    #matrixTable {
        border: 1px solid Black;
        /* padding-bottom: 60px; */
        /* border-spacing: 0.5em; */
        border-collapse: separate;
        border-spacing: 0px 0.6em;
        /* padding: 10px; */
        width: 600px;
    }

    #matrixTitle {
        border: none;
        background-color: lightgrey;
        text-align: center;
        padding: 6px;
    }

    #matrixHeader {
        font-weight: normal;
        letter-spacing: 5pt;
        font-style: italic;
    }

    #matrixFileInfo {
        font-weight: normal;
        font-size: 12px;
        font-style: italic;
        text-align: center;
    }

    #LegendTable {
        border-collapse: collapse;
        border: 1px solid Black;
        table-layout: fixed;
    }

    #LegendTable td {
        text-align: center;
    }

    #probTitle {
        font-weight: bold;
    }

    #probTypeWarning {
        background-color: orange;
    }
    #probTextWarning {
        color: orange;
        font-weight: bold;
    }

    #probTypeError {
        background-color: red;
    }
    #probTextError {
        color: red;
        font-weight: bold;
    }

    #probTypeInfo {
        background-color: lightgrey;
    }

    table tbody tr td a {
        display: block;
        width: 100%;
        height: 100%;
    }
</style>
'@
            #endregion

            #region HTML LegendTable for Mail and Settings
            $htmlLegend = @'
<table id="LegendTable">
    <tr>
        <td id="probTypeError" style="border: 1px solid Black;width: 150px;">Error</td>
        <td id="probTypeWarning" style="border: 1px solid Black;width: 150px;">Warning</td>
        <td id="probTypeInfo" style="border: 1px solid Black;width: 150px;">Information</td>
    </tr>
</table>
'@
            #endregion

            #region HTML Mail overview & Settings detail
            $htmlMatrixTables = foreach ($I in $importedMatrix) {
                #region HTML File
                $FileCheck = if ($I.File.Check) {
                    @'
                    <th id="matrixHeader" colspan="8">File</th>
'@

                    foreach ($F in $I.File.Check) {
                        $ProbType = Get-HTNLidTagProbTypeHC -Name $F.Type

                        $ProbValue = if ($F.Value) {
                            '<ul>'
                            @($F.Value).ForEach( { "<li>$_</li>" })
                            '</ul>'
                        }

                        @"
                        <tr>
                            <td id="$ProbType"></td>
                            <td colspan="7">
                                <p id="probTitle">$($F.Name)</p>
                                <p>$($F.Description)</p>
                                $ProbValue
                            </td>
                        </tr>
"@
                    }
                }
                #endregion

                #region HTML FormData
                $FormDataCheck = if ($I.FormData.Check) {
                    @'
                    <th id="matrixHeader" colspan="8">FormData</th>
'@

                    foreach ($F in $I.FormData.Check) {
                        $ProbType = Get-HTNLidTagProbTypeHC -Name $F.Type

                        $ProbValue = if ($F.Value) {
                            '<ul>'
                            @($F.Value).ForEach( { "<li>$_</li>" })
                            '</ul>'
                        }

                        @"
                        <tr>
                            <td id="$ProbType"></td>
                            <td colspan="7">
                                <p id="probTitle">$($F.Name)</p>
                                <p>$($F.Description)</p>
                                $ProbValue
                            </td>
                        </tr>
"@
                    }
                }
                #endregion

                #region HTML Permissions
                $PermissionsCheck = if ($I.Permissions.Check) {
                    @'
                    <th id="matrixHeader" colspan="8">Permissions</th>
'@

                    foreach ($F in $I.Permissions.Check) {
                        $ProbType = Get-HTNLidTagProbTypeHC -Name $F.Type

                        $ProbValue = if ($F.Value) {
                            '<ul>'
                            @($F.Value).ForEach( {
                                    "<li>$_</li>"
                                })
                            '</ul>'
                        }

                        @"
                        <tr>
                            <td id="$ProbType"></td>
                            <td colspan="7">
                                <p id="probTitle">$($F.Name)</p>
                                <p>$($F.Description)</p>
                                $ProbValue
                            </td>
                        </tr>
"@
                    }
                }
                #endregion

                #region HTML Mail overview Settings table $ Settings detail file
                $MailSettingsTable = $null

                if (
                    ($I.Settings) -and
                    ($I.File.Check.Type -notcontains 'FatalError') -and
                    ($I.Permissions.Check.Type -notcontains 'FatalError')
                ) {
                    $HtmlSettingsHeader = @'
                    <th id="matrixHeader" colspan="8">Settings</th>
                    <tr>
                        <td></td>
                        <td>ID</td>
                        <td>ComputerName</td>
                        <td>Path</td>
                        <td>Action</td>
                        <td>Duration</td>
                    </tr>
'@

                    $MailSettingsTable = $HtmlSettingsHeader

                    foreach ($S in $I.Settings) {
                        #region Get problem color
                        $ProbType = if ($S.Check.Type -contains 'FatalError') {
                            Get-HTNLidTagProbTypeHC -Name 'FatalError'
                        }
                        elseif ($S.Check.Type -contains 'Warning') {
                            Get-HTNLidTagProbTypeHC -Name 'Warning'
                        }
                        elseif ($S.Check.Type -contains 'Information') {
                            Get-HTNLidTagProbTypeHC -Name 'Information'
                        }
                        #endregion

                        #region HTML Settings Create tables
                        $SettingsDetailFatalError = foreach ($E in @($S.Check).Where( { $_.Type -eq 'FatalError' })) {
                            $htmlValue = ConvertTo-HtmlValueHC
                            @"
                            <tr>
                                <td id="probTypeError"></td>
                                <td colspan="7">
                                    <p id="probTitle">$($E.Name)</p>
                                    <p>$($E.Description)</p>
                                    $htmlValue
                                </td>
                            </tr>
"@
                        }

                        $SettingsDetailWarning = foreach ($E in @($S.Check).Where( { $_.Type -eq 'Warning' })) {
                            $htmlValue = ConvertTo-HtmlValueHC
                            @"
                            <tr>
                                <td id="probTypeWarning"></td>
                                <td colspan="7">
                                    <p id="probTitle">$($E.Name)</p>
                                    <p>$($E.Description)</p>
                                    $htmlValue
                                </td>
                            </tr>
"@
                        }

                        $SettingsDetailInfo = foreach ($E in @($S.Check).Where( { $_.Type -eq 'Information' })) {
                            $htmlValue = ConvertTo-HtmlValueHC
                            @"
                            <tr>
                                <td id="probTypeInfo"></td>
                                <td colspan="7">
                                    <p id="probTitle">$($E.Name)</p>
                                    <p>$($E.Description)</p>
                                    $htmlValue
                                </td>
                            </tr>
"@
                        }
                        #endregion

                        #region HTML Settings Create file
                        $SettingsDetail = @"
                            <!DOCTYPE html>
                            <html>
                            <head>
                                <style type="text/css">
                                    body {
                                        font-family: verdana;
                                        background-color: white;
                                    }

                                    h1 {
                                        background-color: black;
                                        color: white;
                                        margin-bottom: 10px;
                                        text-indent: 10px;
                                        page-break-before: always;
                                    }

                                    h2 {
                                        background-color: lightGrey;
                                        margin-bottom: 10px;
                                        text-indent: 10px;
                                        page-break-before: always;
                                    }

                                    h3 {
                                        background-color: lightGrey;
                                        margin-bottom: 10px;
                                        font-size: 16px;
                                        text-indent: 10px;
                                        page-break-before: always;
                                    }

                                    p {
                                        font-size: 14px;
                                        margin-left: 10px;
                                    }

                                    p.italic {
                                        font-style: italic;
                                        font-size: 12px;
                                    }

                                    table {
                                        font-size: 14px;
                                        border-collapse: collapse;
                                        border: 1px none;
                                        padding: 3px;
                                        text-align: left;
                                        padding-right: 10px;
                                        margin-left: 10px;
                                    }

                                    td,
                                    th {
                                        font-size: 14px;
                                        border-collapse: collapse;
                                        border: 1px none;
                                        padding: 3px;
                                        text-align: left;
                                        padding-right: 10px
                                    }

                                    li {
                                        font-size: 14px;
                                    }

                                    base {
                                        target="_blank"
                                    }
                                </style>
                            </head>

                            <body>
                                $htmlStyle
                                <table id="matrixTable">
                                <tr>
                                    <th id="matrixTitle" colspan="8"><a href="$($I.File.SaveFullName)">$($I.File.Item.Name)</a></th>
                                </tr>
                                $HtmlSettingsHeader
                                <tr>
                                    <td id="$ProbType"></td>
                                    <td>$($S.ID)</td>
                                    <td>$($S.Import.ComputerName)</td>
                                    <td>$($S.Import.Path)</td>
                                    <td>$($S.Import.Action)</td>
                                    <td>$(if($D = $S.JobTime.Duration){ '{0:00}:{1:00}:{2:00}' -f $D.Hours, $D.Minutes, $D.Seconds}else{'NA'})</td>
                                </tr>

                                $(if ($SettingsDetailFatalError) {'<th id="matrixHeader" colspan="8">Error</th>' + $SettingsDetailFatalError})
                                $(if ($SettingsDetailWarning) {'<th id="matrixHeader" colspan="8">Warning</th>' + $SettingsDetailWarning})
                                $(if ($SettingsDetailInfo) {'<th id="matrixHeader" colspan="8">Information</th>' + $SettingsDetailInfo})

                                </table>
                                <br>
                                $htmlLegend
                                    <h2>About</h2>
                                <table>
                                    <tr>
                                        <th>GroupName</th>
                                        <td>$($S.Import.GroupName)</td>
                                    </tr>
                                    <tr>
                                        <th>SiteCode</th>
                                        <td>$($S.Import.SiteCode)</td>
                                    </tr>
                                    <tr>
                                        <th>Start time</th>
                                        <td>$(
                                            if ($D = $S.JobTime.Start) {
                                                $D.ToString('dd/MM/yyyy HH:mm:ss (dddd)')
                                            }
                                            else {
                                                'NA'
                                            }
                                            )
                                        </td>
                                    </tr>
                                    <tr>
                                        <th>End time</th>
                                        <td>$(
                                            if ($D = $S.JobTime.End) {
                                                $D.ToString('dd/MM/yyyy HH:mm:ss (dddd)')
                                            }
                                            else {
                                                'NA'
                                            }
                                            )
                                        </td>
                                    </tr>
                                </table>
                            </body>
                        </html>
"@

                        $SettingsFile = Join-Path -Path $I.File.LogFolder -ChildPath "ID $($S.ID) - Settings.html"
                        $SettingsDetail | Out-File -FilePath $SettingsFile -Encoding utf8
                        #endregion

                        $MailSettingsTable += @"
                        <tr>
                            <td id="$ProbType"></td>
                            <td><a href="$SettingsFile">$($S.ID)</a></td>
                            <td><a href="$SettingsFile">$($S.Import.ComputerName)</a></td>
                            <td><a href="$SettingsFile">$($S.Import.Path)</a></td>
                            <td><a href="$SettingsFile">$($S.Import.Action)</a></td>
                            <td><a href="$SettingsFile">$(if($D = $S.JobTime.Duration){ '{0:00}:{1:00}:{2:00}' -f $D.Hours, $D.Minutes, $D.Seconds}else{'NA'})</a></td>
                        </tr>
"@
                    }
                }
                #endregion

                @"
                <table id="matrixTable">
                    <tr>
                        <th id="matrixTitle" colspan="8"><a href="$($I.File.SaveFullName)">$($I.File.Item.Name)</a></th>
                    </tr>
                    <tr>
                        <th id="matrixFileInfo" colspan="8">Last change: $($I.File.ExcelInfo.LastModifiedBy) @ $($I.File.ExcelInfo.Modified.ToString('dd/MM/yyyy HH:mm:ss'))</th>
                    </tr>
                    $FileCheck
                    $FormDataCheck
                    $PermissionsCheck
                    $MailSettingsTable
                </table>
                <br><br>
"@
            }

            #region FatalError and warning count
            $counter = @{
                FormData    = @{
                    Error   = @(
                        $importedMatrix.FormData.Check |
                        Where-Object Type -EQ 'FatalError'
                    ).count
                    Warning = @(
                        $importedMatrix.FormData.Check |
                        Where-Object Type -EQ 'Warning'
                    ).count
                }
                Permissions = @{
                    Error   = @(
                        $importedMatrix.Permissions.Check |
                        Where-Object Type -EQ 'FatalError'
                    ).count
                    Warning = @(
                        $importedMatrix.Permissions.Check |
                        Where-Object Type -EQ 'Warning'
                    ).count
                }
                Settings    = @{
                    Error   = @(
                        $importedMatrix.Settings.Check |
                        Where-Object Type -EQ 'FatalError'
                    ).count
                    Warning = @(
                        $importedMatrix.Settings.Check |
                        Where-Object Type -EQ 'Warning'
                    ).count
                }
                File        = @{
                    Error   = @(
                        $importedMatrix.File.Check |
                        Where-Object Type -EQ 'FatalError'
                    ).count
                    Warning = @(
                        $importedMatrix.File.Check |
                        Where-Object Type -EQ 'Warning'
                    ).count
                }
                Total       = @{
                    Errors   = 0
                    Warnings = 0
                }
            }

            $counter.Total.Errors = (
                $counter.FormData.error + $counter.Permissions.error +
                $counter.Settings.error + $counter.File.error
            )
            $counter.Total.Warnings = (
                $counter.FormData.warning + $counter.Permissions.warning +
                $counter.Settings.warning + $counter.File.warning
            )
            #endregion

            $htmlFormData = if ($Export.FolderPath) {
                @"
            <p><b>Export to <a href="$($Export.FolderPath)">folder</a>:</b></p>
            <table id="overviewTable">
            <tr>
                <th>
                $(
                    if ($accessListSheet.count -and
                        $exportCsvAccessListParams.literalPath -and
                        (Test-Path -LiteralPath $exportCsvAccessListParams.literalPath)
                    ) {
@"
                        <a href="$($exportCsvAccessListParams.literalPath)">Access list</a>
"@
                    }
                    else {'Access list'}
                )
                </th>
                <td>$($accessListSheet.count)</td>
            </tr>
            <tr>
                <th>
                $(
                    if ($adObjectNamesSheet.count -and
                        $exportCsvAdParams.literalPath -and
                        (Test-Path -LiteralPath $exportCsvAdParams.literalPath)
                    ) {
@"
                        <a href="$($exportCsvAdParams.literalPath)">AD objects</a>
"@
                    }
                    else {'AD objects'}
                )
                </th>
                <td>$($adObjectNamesSheet.count)</td>
            </tr>
            <tr>
                <th>
                $(
                    if ($groupManagersSheet.count -and
                        $exportCsvGroupManagersParams.literalPath -and
                        (Test-Path -LiteralPath $exportCsvGroupManagersParams.literalPath)
                    ) {
@"
                        <a href="$($exportCsvGroupManagersParams.literalPath)">Group managers</a>
"@
                    }
                    else {'Group managers'}
                )
                </th>
                <td>$($groupManagersSheet.count)</td>
            </tr>
            <tr>
                <th>
            $(
                if ($formDataSheet.count -and
                    $exportCsvFormParams.literalPath -and
                    (Test-Path -LiteralPath $exportCsvFormParams.literalPath)
                ) {
@"
                    <a href="$($exportCsvFormParams.literalPath)">Form data</a>
"@
                }
                else {'From data'}
            )
                </th>
                <td>$($formDataSheet.count)</td>
            </tr>
            </table>
            $(
                if (
                    ($ExportParams.Path) -and
                    (Test-Path -LiteralPath $ExportParams.Path)
                ) {
@"
<p><i>* Check the <a href="$($ExportParams.Path)">overview</a> for details.</i></p>
"@
                }
            )
            <hr style="width:50%;text-align:left;margin-left:0">
"@
            }

        
            $htmlErrorWarningTable = if ($counter.Total.Errors + $counter.Total.Warnings) {
                @"
            <p><b>Detected issues:</b></p>
            <table id="overviewTable">
            <tr>
                <td></td>
                <td>Errors</td>
                <td>Warnings</td>
            </tr>
            $(
                foreach ($item in ($counter.GetEnumerator())) {
                    if ($item.Value.Error + $item.Value.Warning) {
@"
                    <tr>
                        <th>$($item.Key)</th>
                        <td{0}>$($item.Value.Error)</td>
                        <td{1}>$($item.Value.Warning)</td>
                    </tr>
"@ -f $(if ($item.Value.Error) {' id="probTextError"'}),
$(if ($item.Value.Warning) {' id="probTextWarning"'})
                    }
                }
            )
            </table>
            <p><i>* Check the matrix results below for details.</i></p>
            <hr style="width:50%;text-align:left;margin-left:0">
"@
            }

            $htmlMail = @"
                $htmlStyle
                $htmlErrorWarningTable
                $htmlFormData
                <p><b>Matrix results per file:</b></p>
                $htmlMatrixTables
                $htmlLegend
"@

            $Subject = "$(@($importedMatrix).Count) matrix file{0}{1}{2}" -f $(
                if (@($importedMatrix).Count -ne 1) { 's' }
            ),
            $(
                if ($counter.Total.Errors) {
                    ", $($counter.Total.Errors) error{0}" -f $(
                        if ($counter.Total.Errors -ne 1) { 's' }
                    )
                }
            ),
            $(
                if ($counter.Total.Warnings) {
                    ", $($counter.Total.Warnings) warning{0}" -f $(
                        if ($counter.Total.Warnings -ne 1) { 's' }
                    )
                }
            )

            $MailParams = @{
                To        = $MailTo
                Bcc       = $ScriptAdmin
                Priority  = if ($counter.Total.Errors + $counter.Total.Warnings) { 'High' }
                else { 'Normal' }
                Subject   = $Subject
                Message   = $htmlMail
                Save      = "$matrixLogFile - Mail - $Subject.html"
                Header    = $ScriptName
                LogFolder = $LogFolder
            }
            Get-ScriptRuntimeHC -Stop
            Send-MailHC @MailParams
            #endregion

            #region Non terminating errors are reported to the admin
            # usually when Get-ADObjectDetailHC times out for groups too large
            if ($error) {
                $MailParams = @{
                    To        = $ScriptAdmin
                    Priority  = 'High'
                    Subject   = "FAILURE - $($error.count) non terminating errors"
                    Message   = "While running the permission matrix the following non terminating errors where reported: $($error.Exception.Message | Where-Object { $_  } | ConvertTo-HtmlListHC -Spacing Wide )"
                    Save      = "$matrixLogFile - Mail - $($error.count) non terminating errors.html"
                    Header    = $ScriptName
                    LogFolder = $LogFolder
                }
                Send-MailHC @MailParams
            }
            #endregion
        }
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = $_
            }
        )

        Write-Warning $systemErrors[-1].Message
    }
    finally {
        Get-Job | Remove-Job -Force -EA Ignore
        Remove-PSDrive MatrixFolderPath -EA Ignore

        if ($systemErrors) {
            $M = 'Found {0} system error{1}' -f
            $systemErrors.Count,
            $(if ($systemErrors.Count -ne 1) { 's' })
            Write-Warning $M

            $systemErrors | ForEach-Object {
                Write-Warning $_.Message
            }

            Write-Warning 'Exit script with error code 1'
            exit 1
        }
        else {
            Write-Verbose 'Script finished successfully'
        }
    }
}