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
            '<ul>{0}</ul>' -f $(@($E.Value).ForEach({ "<li>$_</li>" }))
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
            '<ul>
                <li><a href="{0}">{1} items</a></li>
            </ul>' -f $($OutParams.LiteralPath), $($E.Value.Count)
        }
    }
    function Get-HTNLidTagProbTypeHC {
        [OutputType([String])]
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
    function Get-StringValueHC {
        <#
        .SYNOPSIS
            Retrieve a string from the environment variables or a regular string.

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
    function Remove-FileHC {
        param (
            [parameter(Mandatory)]
            [string]$FilePath
        )
                
        if (Test-Path -LiteralPath $FilePath -PathType Leaf) {
            Write-Verbose "Remove file '$FilePath'"
            Remove-Item -Path $FilePath -ErrorAction Ignore
        }
    }
    function Send-MailKitMessageHC {
        <#
            .SYNOPSIS
                Send an email using MailKit and MimeKit assemblies.

            .DESCRIPTION
                This function sends an email using the MailKit and MimeKit
                assemblies. It requires the assemblies to be installed before
                calling the function:

                $params = @{
                    Source           = 'https://www.nuget.org/api/v2'
                    SkipDependencies = $true
                    Scope            = 'AllUsers'
                }
                Install-Package @params -Name 'MailKit'
                Install-Package @params -Name 'MimeKit'

            .PARAMETER MailKitAssemblyPath
                The path to the MailKit assembly.

            .PARAMETER MimeKitAssemblyPath
                The path to the MimeKit assembly.

            .PARAMETER SmtpServerName
                The name of the SMTP server.

            .PARAMETER SmtpPort
                The port of the SMTP server.

            .PARAMETER SmtpConnectionType
                The connection type for the SMTP server.

                Valid values are:
                - 'None'
                - 'Auto'
                - 'SslOnConnect'
                - 'StartTlsWhenAvailable'
                - 'StartTls'

            .PARAMETER Credential
                The credential object containing the username and password.

            .PARAMETER From
                The sender's email address.

            .PARAMETER FromDisplayName
            The display name to show for the sender.

            Email clients may display this differently. It is most likely
            to be shown if the sender's email address is not recognized
                (e.g., not in the address book).

            .PARAMETER To
                The recipient's email address.

            .PARAMETER Body
            The body of the email, HTML is supported.

            .PARAMETER Subject
            The subject of the email.

            .PARAMETER Attachments
            An array of file paths to attach to the email.

            .PARAMETER Priority
            The email priority.

            Valid values are:
            - 'Low'
            - 'Normal'
            - 'High'

            .EXAMPLE
            # Send an email with StartTls and credential

            $SmtpUserName = 'smtpUser'
            $SmtpPassword = 'smtpPassword'

            $securePassword = ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($SmtpUserName, $securePassword)

            $params = @{
                SmtpServerName = 'SMT_SERVER@example.com'
                SmtpPort = 587
                SmtpConnectionType = 'StartTls'
                Credential = $credential
                from = 'm@example.com'
                To = '007@example.com'
                Body = '<p>Mission details in attachment</p>'
                Subject = 'For your eyes only'
                Priority = 'High'
                Attachments = @('c:\Mission.ppt', 'c:\ID.pdf')
                MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
            }

            Send-MailKitMessageHC @params

            .EXAMPLE
            # Send an email without authentication

            $params = @{
                SmtpServerName      = 'SMT_SERVER@example.com'
                SmtpPort            = 25
                From                = 'hacker@example.com'
                FromDisplayName     = 'White hat hacker'
                Bcc                 = @('james@example.com', 'mike@example.com')
                Body                = '<h1>You have been hacked</h1>'
                Subject             = 'Oops'
                MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
            }

            Send-MailKitMessageHC @params
            #>

        [CmdletBinding()]
        param (
            [parameter(Mandatory)]
            [string]$MailKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$MimeKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$SmtpServerName,
            [parameter(Mandatory)]
            [ValidateSet(25, 465, 587, 2525)]
            [int]$SmtpPort,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
            [parameter(Mandatory)]
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [string]$FromDisplayName,
            [string[]]$To,
            [string[]]$Bcc,
            [int]$MaxAttachmentSize = 20MB,
            [ValidateSet(
                'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
            )]
            [string]$SmtpConnectionType = 'None',
            [ValidateSet('Normal', 'Low', 'High')]
            [string]$Priority = 'Normal',
            [string[]]$Attachments,
            [PSCredential]$Credential
        )

        begin {
            function Test-IsAssemblyLoaded {
                param (
                    [String]$Name
                )
                foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
                    if ($assembly.FullName -like "$Name, Version=*") {
                        return $true
                    }
                }
                return $false
            }

            function Add-Attachments {
                param (
                    [string[]]$Attachments,
                    [MimeKit.Multipart]$BodyMultiPart
                )

                $attachmentList = New-Object System.Collections.ArrayList($null)

                foreach (
                    $attachmentPath in
                    $Attachments | Sort-Object -Unique
                ) {
                    try {
                        #region Test if file exists
                        try {
                            $attachmentItem = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop

                            if ($attachmentItem.PSIsContainer) {
                                Write-Warning "Attachment '$attachmentPath' is a folder, not a file"
                                continue
                            }
                        }
                        catch {
                            Write-Warning "Attachment '$attachmentPath' not found"
                            continue
                        }
                        #endregion

                        $totalSizeAttachments += $attachmentItem.Length

                        $null = $attachmentList.Add($attachmentItem)

                        #region Check size of attachments
                        if ($totalSizeAttachments -ge $MaxAttachmentSize) {
                            $M = 'The maximum allowed attachment size of {0} MB has been exceeded ({1} MB). No attachments were added to the email. Check the log folder for details.' -f
                            ([math]::Round(($MaxAttachmentSize / 1MB))),
                            ([math]::Round(($totalSizeAttachments / 1MB), 2))

                            Write-Warning $M

                            return [PSCustomObject]@{
                                AttachmentLimitExceededMessage = $M
                            }
                        }
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentPath': $_"
                    }
                }
                #endregion

                foreach (
                    $attachmentItem in
                    $attachmentList
                ) {
                    try {
                        Write-Verbose "Add mail attachment '$($attachmentItem.Name)'"

                        $attachment = New-Object MimeKit.MimePart

                        #region Create a MemoryStream to hold the file content
                        $memoryStream = New-Object System.IO.MemoryStream

                        try {
                            $fileStream = [System.IO.File]::OpenRead($attachmentItem.FullName)
                            $fileStream.CopyTo($memoryStream)
                        }
                        finally {
                            if ($fileStream) {
                                $fileStream.Dispose()
                            }
                        }

                        $memoryStream.Position = 0
                        #endregion

                        $attachment.Content = New-Object MimeKit.MimeContent($memoryStream)

                        $attachment.ContentDisposition = New-Object MimeKit.ContentDisposition

                        $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64

                        $attachment.FileName = $attachmentItem.Name

                        $bodyMultiPart.Add($attachment)
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentItem': $_"
                    }
                }
            }

            try {
                #region Test To or Bcc required
                if (-not ($To -or $Bcc)) {
                    throw "Either 'To' to 'Bcc' is required for sending emails"
                }
                #endregion

                #region Test To
                foreach ($email in $To) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "To email address '$email' not valid."
                    }
                }
                #endregion

                #region Test Bcc
                foreach ($email in $Bcc) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "Bcc email address '$email' not valid."
                    }
                }
                #endregion

                #region Load MimeKit assembly
                if (-not(Test-IsAssemblyLoaded -Name 'MimeKit')) {
                    try {
                        Write-Verbose "Load MimeKit assembly '$MimeKitAssemblyPath'"
                        Add-Type -Path $MimeKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MimeKit assembly '$MimeKitAssemblyPath': $_"
                    }
                }
                #endregion

                #region Load MailKit assembly
                if (-not(Test-IsAssemblyLoaded -Name 'MailKit')) {
                    try {
                        Write-Verbose "Load MailKit assembly '$MailKitAssemblyPath'"
                        Add-Type -Path $MailKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MailKit assembly '$MailKitAssemblyPath': $_"
                    }
                }
                #endregion
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
        }

        process {
            try {
                $message = New-Object -TypeName 'MimeKit.MimeMessage'

                #region Create body with attachments
                $bodyPart = New-Object MimeKit.TextPart('html')
                $bodyPart.Text = $Body

                $bodyMultiPart = New-Object MimeKit.Multipart('mixed')
                $bodyMultiPart.Add($bodyPart)

                if ($Attachments) {
                    $params = @{
                        Attachments   = $Attachments
                        BodyMultiPart = $bodyMultiPart
                    }
                    $addAttachments = Add-Attachments @params

                    if ($addAttachments.AttachmentLimitExceededMessage) {
                        $bodyPart.Text += '<p><i>{0}</i></p>' -f
                        $addAttachments.AttachmentLimitExceededMessage
                    }
                }

                $message.Body = $bodyMultiPart
                #endregion

                $fromAddress = New-Object MimeKit.MailboxAddress(
                    $FromDisplayName, $From
                )
                $message.From.Add($fromAddress)

                foreach ($email in $To) {
                    $message.To.Add($email)
                }

                foreach ($email in $Bcc) {
                    $message.Bcc.Add($email)
                }

                $message.Subject = $Subject

                #region Set priority
                switch ($Priority) {
                    'Low' {
                        $message.Headers.Add('X-Priority', '5 (Lowest)')
                        break
                    }
                    'Normal' {
                        $message.Headers.Add('X-Priority', '3 (Normal)')
                        break
                    }
                    'High' {
                        $message.Headers.Add('X-Priority', '1 (Highest)')
                        break
                    }
                    default {
                        throw "Priority type '$_' not supported"
                    }
                }
                #endregion

                $smtp = New-Object -TypeName 'MailKit.Net.Smtp.SmtpClient'

                try {
                    $smtp.Connect(
                        $SmtpServerName, $SmtpPort,
                        [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType
                    )
                }
                catch {
                    throw "Failed to connect to SMTP server '$SmtpServerName' on port '$SmtpPort' with connection type '$SmtpConnectionType': $_"
                }

                if ($Credential) {
                    try {
                        $smtp.Authenticate(
                            $Credential.UserName,
                            $Credential.GetNetworkCredential().Password
                        )
                    }
                    catch {
                        throw "Failed to authenticate with user name '$($Credential.UserName)' to SMTP server '$SmtpServerName': $_"
                    }
                }

                Write-Verbose "Send mail to '$To' with subject '$Subject'"

                $null = $smtp.Send($message)
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
            finally {
                if ($smtp) {
                    $smtp.Disconnect($true)
                    $smtp.Dispose()
                }
                if ($message) {
                    $message.Dispose()
                }
            }
        }
    }
    function Write-EventsToEventLogHC {
        <#
        .SYNOPSIS
            Write events to the event log.

        .DESCRIPTION
            The use of this function will allow standardization in the Windows
            Event Log by using the same EventID's and other properties across
            different scripts.

            Custom Windows EventID's based on the PowerShell standard streams:

            PowerShell Stream     EventIcon    EventID   EventDescription
            -----------------     ---------    -------   ----------------
            [i] Info              [i] Info     100       Script started
            [4] Verbose           [i] Info     4         Verbose message
            [1] Output/Success    [i] Info     1         Output on success
            [3] Warning           [w] Warning  3         Warning message
            [2] Error             [e] Error    2         Fatal error message
            [i] Info              [i] Info     199       Script ended successfully

        .PARAMETER Source
            Specifies the script name under which the events will be logged.

        .PARAMETER LogName
            Specifies the name of the event log to which the events will be
            written. If the log does not exist, it will be created.

        .PARAMETER Events
            Specifies the events to be written to the event log. This should be
            an array of PSCustomObject with properties: Message, EntryType, and
            EventID.

        .PARAMETER Events.xxx
            All properties that are not 'EntryType' or 'EventID' will be used to
            create a formatted message.

        .PARAMETER Events.EntryType
            The type of the event.

            The following values are supported:
            - Information
            - Warning
            - Error
            - SuccessAudit
            - FailureAudit

            The default value is Information.

        .PARAMETER Events.EventID
            The ID of the event. This should be a number.
            The default value is 4.

        .EXAMPLE
            $eventLogData = [System.Collections.Generic.List[PSObject]]::new()

            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script started'
                    EntryType = 'Information'
                    EventID   = '100'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message  = 'Failed to read the file'
                    FileName = 'C:\Temp\test.txt'
                    DateTime = Get-Date
                    EntryType = 'Error'
                    EventID   = '2'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message  = 'Created file'
                    FileName = 'C:\Report.xlsx'
                    FileSize = 123456
                    DateTime = Get-Date
                    EntryType = 'Information'
                    EventID   = '1'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script finished'
                    EntryType = 'Information'
                    EventID   = '199'
                }
            )

            $params = @{
                Source  = 'Test (Brecht)'
                LogName = 'HCScripts'
                Events  = $eventLogData
            }
            Write-EventsToEventLogHC @params
        #>

        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [String]$Source,
            [Parameter(Mandatory)]
            [String]$LogName,
            [PSCustomObject[]]$Events
        )

        try {
            if (
                -not(
                    ([System.Diagnostics.EventLog]::Exists($LogName)) -and
                    [System.Diagnostics.EventLog]::SourceExists($Source)
                )
            ) {
                Write-Verbose "Create event log '$LogName' and source '$Source'"
                New-EventLog -LogName $LogName -Source $Source -ErrorAction Stop
            }

            foreach ($eventItem in $Events) {
                $params = @{
                    LogName     = $LogName
                    Source      = $Source
                    EntryType   = $eventItem.EntryType
                    EventID     = $eventItem.EventID
                    Message     = ''
                    ErrorAction = 'Stop'
                }

                if (-not $params.EntryType) {
                    $params.EntryType = 'Information'
                }
                if (-not $params.EventID) {
                    $params.EventID = 4
                }

                foreach (
                    $property in
                    $eventItem.PSObject.Properties | Where-Object {
                        ($_.Name -ne 'EntryType') -and ($_.Name -ne 'EventID')
                    }
                ) {
                    $params.Message += "`n- $($property.Name) '$($property.Value)'"
                }

                Write-Verbose "Write event to log '$LogName' source '$Source' message '$($params.Message)'"

                Write-EventLog @params
            }
        }
        catch {
            throw "Failed to write to event log '$LogName' source '$Source': $_"
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

        $Matrix = $jsonFileContent.Matrix
        $Export = $jsonFileContent.Export
        $ServiceNow = $jsonFileContent.ServiceNow
        $MaxConcurrent = $jsonFileContent.MaxConcurrent
        $ExcludedSamAccountName = $jsonFileContent.Matrix.ExcludedSamAccountName
        $DetailedLog = $jsonFileContent.Settings.SaveLogFiles.Detailed
        $LogFolder = $jsonFileContent.Settings.SaveLogFiles.Where.Folder

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

        #region Set PSSessionConfiguration
        $PSSessionConfiguration = $jsonFileContent.PSSessionConfiguration

        if (-not $PSSessionConfiguration) {
            $PSSessionConfiguration = 'PowerShell.7'
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
    if ($systemErrors) { return }

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
                        if ($Export.ServiceNowFormDataExcelFile) {
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
                                    Description = "When the argument 'Export.ServiceNowFormDataExcelFile' is used the Excel file needs to have a worksheet 'FormData'."
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

                return
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
}

end {
    try {
        $settings = $jsonFileContent.Settings

        $scriptName = $settings.ScriptName
        $saveInEventLog = $settings.SaveInEventLog
        $sendMail = $settings.SendMail
        $saveLogFiles = $settings.SaveLogFiles

        #region Get script name
        if (-not $scriptName) {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Input file '$ConfigurationJsonFile': No 'Settings.ScriptName' found."
                }
            )

            Write-Warning $systemErrors[-1].Message

            $scriptName = 'Default script name'
        }
        #endregion

        $mailParams = @{
            From                = Get-StringValueHC $sendMail.From
            Subject             = "$($counter.Total.MovedFiles) moved"
            SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
            SmtpPort            = Get-StringValueHC $sendMail.Smtp.Port
            MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
            MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
        }

        if (-not $systemErrors) {
            if ($importedMatrix) {
                $dataToExport = @{
                    AccessList    = @()
                    AdObjects     = @()
                    FormData      = @()
                    GroupManagers = @()
                }

                #region Get matrix log file name
                $matrixLogFileBasePath = Join-Path -Path $LogFolder -ChildPath (
                    '{0:00}-{1:00}-{2:00} {3:00}{4:00} ({5})' -f
                    $scriptStartTime.Year, $scriptStartTime.Month, $scriptStartTime.Day,
                    $scriptStartTime.Hour, $scriptStartTime.Minute, $scriptStartTime.DayOfWeek
                )
                #endregion

                #region Add sheets to Matrix log files and collect data
                foreach ($I in $importedMatrix) {
                    $excelParams = @{
                        Path               = $I.File.SaveFullName
                        AutoSize           = $true
                        ClearSheet         = $true
                        FreezeTopRow       = $true
                        NoNumberConversion = '*'
                    }
                
                    #region Get SamAccountNames for rows Settings sheet
                    $matrixSamAccountNames = $i.Settings.AdObjects.Values.SamAccountName |
                    Select-Object -Property @{
                        Name       = 'name'
                        Expression = { "$($_)".Trim() }
                    } -Unique |
                    Select-Object -ExpandProperty name

                    Write-Verbose "Matrix '$($i.File.Item.Name)' has '$($matrixSamAccountNames.count)' unique SamAccountNames"
                    #endregion

                    #region AccessList
                    #region Create object
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

                    if ($accessListToExport) {
                        #region Export to Excel
                        $excelParams.WorksheetName = 'AccessList'
                        $excelParams.TableName = 'AccessList'
              
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

                        #region Add to export
                        $dataToExport['AccessList'] += $accessListToExport |
                        Select-Object @{
                            Name       = 'MatrixFileName'
                            Expression = { $I.File.Item.BaseName }
                        }, *
                        #endregion
                    }
                    #endregion

                    #region GroupManagers
                    #region Create objects
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

                    if ($groupManagersToExport) {
                        #region Export to Excel
                        $excelParams.WorksheetName = 'GroupManagers'
                        $excelParams.TableName = 'GroupManagers'

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

                        #region Add to export
                        $dataToExport['GroupManagers'] += $groupManagersToExport |
                        Select-Object @{
                            Name       = 'MatrixFileName'
                            Expression = { $I.File.Item.BaseName }
                        }, *
                        #endregion
                    }
                    #endregion

                    #region AdObjects
                    $adObjects = foreach (
                        $S in
                        $I.Settings.Where( { $_.AdObjects.Count -ne 0 })
                    ) {
                        foreach ($A in ($S.AdObjects.GetEnumerator())) {
                            [PSCustomObject]@{
                                MatrixFileName = $I.File.Item.BaseName
                                SamAccountName = $A.Value.SamAccountName
                                GroupName      = $A.Value.Converted.Begin
                                SiteCode       = $A.Value.Converted.Middle
                                Name           = $A.Value.Converted.End
                            }
                        }
                    }

                    if ($adObjects) {
                        #region Export to Excel
                        $excelParams.WorksheetName = 'AdObjects'
                        $excelParams.TableName = 'AdObjects'
                        
                        $AdObjects | Export-Excel @ExportParams
                        #endregion

                        #region Add to export
                        $dataToExport['AdObjects'] += $adObjects |
                        Group-Object SamAccountName |
                        ForEach-Object { $_.Group[0] }
                        #endregion
                    }
                    #endregion

                    #region FormData
                    $dataToExport['FormData'] += $I.FormData.Import                    
                    #endregion
                }
                #endregion

                $exportedFiles = @{}

                #region Create Permissions Excel files
                if ($Export.PermissionsExcelFile) {
                    Remove-FileHC -FilePath $Export.PermissionsExcelFile

                    #region Add sheets to Permissions log file
                    $permissionsExcelLogFileParams = @{
                        Path         = "$matrixLogFileBasePath - AllMatrix - Permissions.xlsx"
                        AutoSize     = $true
                        FreezeTopRow = $true
                    }

                    foreach (
                        $property in    
                        $dataToExport.GetEnumerator() | Where-Object { 
                            $_.Value
                        }
                    ) {
                        try {
                            $name = $property.Name
                            $data = $property.Value

                            $permissionsExcelLogFileParams.WorksheetName = $name
                            $permissionsExcelLogFileParams.TableName = $name

                            $eventLogData.Add(
                                [PSCustomObject]@{
                                    Message   = "Export $($data.Count) $Name objects to '$($permissionsExcelLogFileParams.Path)'"
                                    DateTime  = Get-Date
                                    EntryType = 'Information'
                                    EventID   = '1'
                                }
                            )
                            Write-Verbose $eventLogData[-1].Message
                   
                            $data | Export-Excel @permissionsExcelLogFileParams
                        }
                        catch {
                            $systemErrors.Add(
                                [PSCustomObject]@{
                                    DateTime = Get-Date
                                    Message  = "Failed to export all matrix data to .XLSX and .CSV files for '$($property.Name)': $_"
                                }
                            )
                                
                            Write-Warning $systemErrors[-1].Message
                        }
                    }
                    #endregion
                        
                    #region Copy Permissions Excel file from log to prod folder
                    if (
                        $Export.PermissionsExcelFile -and
                        (Test-Path -LiteralPath $permissionsExcelLogFileParams.Path -PathType Leaf)
                    ) {
                        $copyParams = @{
                            LiteralPath = $permissionsExcelLogFileParams.Path
                            Destination = $Export.PermissionsExcelFile
                        }

                        $eventLogData.Add(
                            [PSCustomObject]@{
                                Message   = "Copy file '$($copyParams.LiteralPath)' to '$($copyParams.Destination)'"
                                DateTime  = Get-Date
                                EntryType = 'Information'
                                EventID   = '1'
                            }
                        )
                        Write-Verbose $eventLogData[-1].Message

                        Copy-Item @copyParams
                    }
                    #endregion

                    $exportedFiles['PermissionsExcelFile'] = $Export.PermissionsExcelFile
                }
                #endregion

                #region Create ServiceNowFormData and OverviewHTML file
                if (
                    $dataToExport['FormData'] -and
                    $importedMatrix.FormData.Check.Type -notcontains 'FatalError'
                ) {
                    if ($Export.ServiceNowFormDataExcelFile) {
                        Remove-FileHC -FilePath $Export.ServiceNowFormDataExcelFile

                        #region Create objects for ServiceNow
                        Write-Verbose 'Create objects for ServiceNow form'

                        $serviceNowFormData = foreach (
                            $adObjectName in 
                            $dataToExport.AdObjects
                        ) {
    
                            $formData = $dataToExport.FormData.Where(
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

                        #region Export to Excel
                        $params = @{
                            Path         = $Export.ServiceNowFormDataExcelFile
                            AutoSize     = $true
                            FreezeTopRow = $true
                        }

                        $serviceNowFormData | Export-Excel @params
                        #endregion

                        $exportedFiles['ServiceNowFormDataExcelFile'] = $params.Path
              
                        #region Start ServiceNow FormData upload
                        if (
                            $ServiceNow.CredentialsFilePath -and
                            $ServiceNow.Environment -and
                            $ServiceNow.TableName -and
                            $serviceNowFormData
                        ) {
                            try {
                                $params = @{
                                    CredentialsFilePath = $ServiceNow.CredentialsFilePath
                                    Environment         = $ServiceNow.Environment
                                    TableName           = $ServiceNow.TableName
                                    FormDataFile        = $params.Path
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
                        }
                        else {
                            $systemErrors.Add(
                                [PSCustomObject]@{
                                    DateTime = Get-Date
                                    Message  = 'Parameter missing to upload data to ServiceNow'
                                }
                            )

                            Write-Warning $systemErrors[-1].Message
                        }
                        #endregion
                    }
                    if ($Export.OverviewHtmlFile) {
                        Remove-FileHC -FilePath $Export.OverviewHtmlFile
                
                        #region Export FormData to HTML file
                        try {
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
                                '<h1>Matrix files overview</h1>',
                                '<table>
                                    <tr>
                                        <th>Category</th>
                                        <th>Subcategory</th>
                                        <th>Folder</th>
                                        <th>Link to the matrix</th>
                                        <th>Responsible</th>
                                    </tr>'
                            )

                            $htmlFileContent += $dataToExport['FormData'] | 
                            Sort-Object -Property 'MatrixCategoryName', 
                            'MatrixSubCategoryName', 
                            'MatrixFolderDisplayName' | 
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

                            $htmlFileContent += '</table>'

                            $params = @{
                                LiteralPath = $Export.OverviewHtmlFile 
                                Encoding    = 'utf8'
                                Force       = $true
                            }

                            $eventLogData.Add(
                                [PSCustomObject]@{
                                    Message   = "Export FormData to '$($params.LiteralPath)'"
                                    DateTime  = Get-Date
                                    EntryType = 'Information'
                                    EventID   = '1'
                                }
                            )
                            Write-Verbose $eventLogData[-1].Message

                            $htmlFileContent | Out-File @params
                        }
                        catch {
                            $systemErrors.Add(
                                [PSCustomObject]@{
                                    DateTime = Get-Date
                                    Message  = "Failed to export FormData to HTML file '$($Export.OverviewHtmlFile)': $_"
                                }
                            )

                            Write-Warning $systemErrors[-1].Message
                        }
                        #endregion

                        $exportedFiles['OverviewHtmlFile'] = $Export.OverviewHtmlFile
                    }
                }
                #endregion

                $html = @{}

                #region HTML Style and table legend
                Write-Verbose 'Format HTML'

                $html.Style = '<style>
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
                </style>'

                $html.LegendTable = '
                <table id="LegendTable">
                    <tr>
                        <td id="probTypeError" style="border: 1px solid Black;width: 150px;">Error</td>
                        <td id="probTypeWarning" style="border: 1px solid Black;width: 150px;">Warning</td>
                        <td id="probTypeInfo" style="border: 1px solid Black;width: 150px;">Information</td>
                    </tr>
                </table>'
                #endregion

                #region HTML Mail overview & Settings detail
                $htmlMatrixTables = foreach ($I in $importedMatrix) {
                    #region HTML File
                    $FileCheck = if ($I.File.Check) {
                        @'
                    <th id="matrixHeader" colspan="8">File</th>
'@

                        foreach ($F in $I.File.Check) {
                            $problem = @{
                                Type        = Get-HTNLidTagProbTypeHC -Name $F.Type
                                Details     = if ($F.Value) {
                                    '<ul>'
                                    @($F.Value).ForEach( { "<li>$_</li>" })
                                    '</ul>'
                                }
                                Name        = $F.Name
                                Description = $F.Description
                            }

                            '<tr>
                                <td id="{0}"></td>
                                <td colspan="7">
                                    <p id="probTitle">{1}</p>
                                    <p>{2}</p>
                                    {3}
                                </td>
                            </tr>' -f 
                            $($problem.Type), 
                            $($problem.Name), 
                            $($problem.Description), 
                            $($problem.Details)
                        }
                    }
                    #endregion

                    #region HTML FormData
                    $FormDataCheck = if ($I.FormData.Check) {
                        @'
                    <th id="matrixHeader" colspan="8">FormData</th>
'@

                        foreach ($F in $I.FormData.Check) {
                            $problem = @{
                                Type        = Get-HTNLidTagProbTypeHC -Name $F.Type
                                Details     = if ($F.Value) {
                                    '<ul>'
                                    @($F.Value).ForEach( { "<li>$_</li>" })
                                    '</ul>'
                                }
                                Name        = $F.Name
                                Description = $F.Description
                            }

                            '<tr>
                                <td id="{0}"></td>
                                <td colspan="7">
                                    <p id="probTitle">{1}</p>
                                    <p>{2}</p>
                                    {3}
                                </td>
                            </tr>' -f 
                            $($problem.Type), 
                            $($problem.Name), 
                            $($problem.Description), 
                            $($problem.Details)
                        }
                    }
                    #endregion

                    #region HTML Permissions
                    $PermissionsCheck = if ($I.Permissions.Check) {
                        '<th id="matrixHeader" colspan="8">Permissions</th>'

                        foreach ($F in $I.Permissions.Check) {
                            $problem = @{
                                Type        = Get-HTNLidTagProbTypeHC -Name $F.Type
                                Details     = if ($F.Value) {
                                    '<ul>'
                                    @($F.Value).ForEach( { "<li>$_</li>" })
                                    '</ul>'
                                }
                                Name        = $F.Name
                                Description = $F.Description
                            }

                            '<tr>
                                <td id="{0}"></td>
                                <td colspan="7">
                                    <p id="probTitle">{1}</p>
                                    <p>{2}</p>
                                    {3}
                                </td>
                            </tr>' -f 
                            $($problem.Type), 
                            $($problem.Name), 
                            $($problem.Description), 
                            $($problem.Details)
                        }
                    }
                    #endregion

                    #region HTML Mail overview Settings table $ Settings detail file
                    $html.Mail = @{}

                    if (
                        ($I.Settings) -and
                        ($I.File.Check.Type -notcontains 'FatalError') -and
                        ($I.Permissions.Check.Type -notcontains 'FatalError')
                    ) {
                        $html.Mail.SettingsHeader = '
                        <th id="matrixHeader" colspan="8">Settings</th>
                        <tr>
                            <td></td>
                            <td>ID</td>
                            <td>ComputerName</td>
                            <td>Path</td>
                            <td>Action</td>
                            <td>Duration</td>
                        </tr>'

                        $html.Mail.SettingsTable = $html.Mail.SettingsHeader

                        foreach ($S in $I.Settings) {
                            $problem = @{}

                            #region Get problem color
                            $problem.Type = if ($S.Check.Type -contains 'FatalError') {
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
                            $html.MatrixLogFile = @{}

                            $html.MatrixLogFile.FatalError = foreach ($E in @($S.Check).Where( { $_.Type -eq 'FatalError' })) {
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

                            $html.MatrixLogFile.Warning = foreach ($E in @($S.Check).Where( { $_.Type -eq 'Warning' })) {
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

                            $html.MatrixLogFile.Info = foreach ($E in @($S.Check).Where( { $_.Type -eq 'Information' })) {
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
                            $html.MatrixLogFile.Table =
                            '<!DOCTYPE html>
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
                            </head>'

                            $html.MatrixLogFile.Table += @"
                            <body>
                                $($html.Style)
                                <table id="matrixTable">
                                <tr>
                                    <th id="matrixTitle" colspan="8"><a href="$($I.File.SaveFullName)">$($I.File.Item.Name)</a></th>
                                </tr>
                                $($html.Mail.SettingsHeader)
                                <tr>
                                    <td id="$($problem.Type)"></td>
                                    <td>$($S.ID)</td>
                                    <td>$($S.Import.ComputerName)</td>
                                    <td>$($S.Import.Path)</td>
                                    <td>$($S.Import.Action)</td>
                                    <td>$(
                                        if($D = $S.JobTime.Duration) {
                                            '{0:00}:{1:00}:{2:00}' -f 
                                            $D.Hours, $D.Minutes, $D.Seconds
                                        }
                                        else{'NA'}
                                        )
                                    </td>
                                </tr>

                                $(if ($html.MatrixLogFile.FatalError) {'<th id="matrixHeader" colspan="8">Error</th>' + $html.MatrixLogFile.FatalError})
                                $(if ($html.MatrixLogFile.Warning) {'<th id="matrixHeader" colspan="8">Warning</th>' + $html.MatrixLogFile.Warning})
                                $(if ($html.MatrixLogFile.Info) {'<th id="matrixHeader" colspan="8">Information</th>' + $html.MatrixLogFile.Info})

                                </table>
                                <br>
                                $($html.LegendTable)
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

                            $matrixLogFileParams = @{
                                FilePath = Join-Path -Path $I.File.LogFolder -ChildPath "ID $($S.ID) - Settings.html"
                                Encoding = 'utf8'
                            }
                            $html.MatrixLogFile.Table | 
                            Out-File @matrixLogFileParams
                            #endregion

                            $html.Mail.SettingsTable += "
                        <tr>
                            <td id=`"$($problem.Type)`"></td>
                            <td><a href=`"{0}`">$($S.ID)</a></td>
                            <td><a href=`"{0}`">$($S.Import.ComputerName)</a></td>
                            <td><a href=`"{0}`">$($S.Import.Path)</a></td>
                            <td><a href=`"{0}`">$($S.Import.Action)</a></td>
                            <td><a href=`"{0}`">{1}</a></td>
                        </tr>" -f 
                            $($matrixLogFileParams.FilePath), 
                            $(
                                if ($D = $S.JobTime.Duration) {
                                    '{0:00}:{1:00}:{2:00}' -f
                                    $D.Hours, $D.Minutes, $D.Seconds
                                }
                                else { 'NA' })
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
                    $($html.Mail.SettingsTable)
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

                $htmlExportFiles = if ($exportedFiles.Count) {
                    @"
            <p><b>Exported $($exportedFiles.Count) files:</b></p>
            <ul>
            $(
                $exportedFiles.GetEnumerator() | ForEach-Object {
                    "<li><a href=`"$($_.Value)`">$($_.Key)</a></li>"
                }
            )
            </ul>
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
                $($html.Style)
                $htmlErrorWarningTable
                $htmlExportFiles
                <p><b>Matrix results per file:</b></p>
                $htmlMatrixTables
                $($html.LegendTable)
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
                    Save      = "$matrixLogFileBasePath - Mail - $Subject.html"
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
                        Save      = "$matrixLogFileBasePath - Mail - $($error.count) non terminating errors.html"
                        Header    = $ScriptName
                        LogFolder = $LogFolder
                    }
                    Send-MailHC @MailParams
                }
                #endregion
            }
        }

        #region Send email
        try {
            $isSendMail = $false

            if ($ReportOnly) {
                $isSendMail = $true
            }
            else {
                switch ($sendMail.When) {
                    'Never' {
                        break
                    }
                    'Always' {
                        $isSendMail = $true
                        break
                    }
                    'OnError' {
                        if ($counter.Total.Errors) {
                            $isSendMail = $true
                        }
                        break
                    }
                    'OnErrorOrAction' {
                        if ($counter.Total.Errors -or $logFileData) {
                            $isSendMail = $true
                        }
                        break
                    }
                    default {
                        throw "SendMail.When '$($sendMail.When)' not supported. Supported values are 'Never', 'Always', 'OnError' or 'OnErrorOrAction'."
                    }
                }
            }

            if ($isSendMail) {
                #region Test mandatory fields
                @{
                    'From'                 = $sendMail.From
                    'Smtp.ServerName'      = $sendMail.Smtp.ServerName
                    'Smtp.Port'            = $sendMail.Smtp.Port
                    'AssemblyPath.MailKit' = $sendMail.AssemblyPath.MailKit
                    'AssemblyPath.MimeKit' = $sendMail.AssemblyPath.MimeKit
                }.GetEnumerator() |
                Where-Object { -not $_.Value } | ForEach-Object {
                    throw "Input file property 'Settings.SendMail.$($_.Key)' cannot be blank"
                }
                #endregion

                $mailParams = @{
                    From                = Get-StringValueHC $sendMail.From
                    Subject             = "$($counter.Total.MovedFiles) moved"
                    SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
                    SmtpPort            = Get-StringValueHC $sendMail.Smtp.Port
                    MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
                    MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
                }

                $mailParams.Body = @"
<!DOCTYPE html>
<html>
<head>
<style type="text/css">
    body {
        font-family:verdana;
        font-size:14px;
        background-color:white;
    }
    h1 {
        margin-bottom: 0;
    }
    h2 {
        margin-bottom: 0;
    }
    h3 {
        margin-bottom: 0;
    }
    p.italic {
        font-style: italic;
        font-size: 12px;
    }
    table {
        border-collapse:collapse;
        border:0px none;
        padding:3px;
        text-align:left;
    }
    td, th {
        border-collapse:collapse;
        border:1px none;
        padding:3px;
        text-align:left;
    }
    #aboutTable th {
        color: rgb(143, 140, 140);
        font-weight: normal;
    }
    #aboutTable td {
        color: rgb(143, 140, 140);
        font-weight: normal;
    }
    base {
        target="_blank"
    }
</style>
</head>
<body>
<table>
    <h1>$scriptName</h1>
    <hr size="2" color="#06cc7a">

    $($sendMail.Body)

    $(
        if ($systemErrors.Count) {
            '<table>
                <tr style="background-color: #ffe5ec;">
                    <th>System errors</th>
                    <td>{0}</td>
                </tr>
            </table>' -f $($systemErrors.Count)
        }
    )

    $(
        if ($ReportOnly) {
            '<p>Summary of all SFTP actions <b>executed today</b>:</p>'
        }
        else {
            '<p>Summary of SFTP actions:</p>'
        }
    )

    $htmlTable

    $(
        if ($allLogFilePaths) {
            '<p><i>* Check the attachment(s) for details</i></p>'
        }
    )

    <hr size="2" color="#06cc7a">
    <table id="aboutTable">
        $(
            if ($scriptStartTime) {
                '<tr>
                    <th>Start time</th>
                    <td>{0:00}/{1:00}/{2:00} {3:00}:{4:00} ({5})</td>
                </tr>' -f
                $scriptStartTime.Day,
                $scriptStartTime.Month,
                $scriptStartTime.Year,
                $scriptStartTime.Hour,
                $scriptStartTime.Minute,
                $scriptStartTime.DayOfWeek
            }
        )
        $(
            if ($scriptStartTime) {
                $runTime = New-TimeSpan -Start $scriptStartTime -End (Get-Date)
                '<tr>
                    <th>Duration</th>
                    <td>{0:00}:{1:00}:{2:00}</td>
                </tr>' -f
                $runTime.Hours, $runTime.Minutes, $runTime.Seconds
            }
        )
        $(
            if ($logFolderPath) {
                '<tr>
                    <th>Log files</th>
                    <td><a href="{0}">Open log folder</a></td>
                </tr>' -f $logFolderPath
            }
        )
        <tr>
            <th>Host</th>
            <td>$($host.Name)</td>
        </tr>
        <tr>
            <th>PowerShell</th>
            <td>$($PSVersionTable.PSVersion.ToString())</td>
        </tr>
        <tr>
            <th>Computer</th>
            <td>$env:COMPUTERNAME</td>
        </tr>
        <tr>
            <th>Account</th>
            <td>$env:USERDNSDOMAIN\$env:USERNAME</td>
        </tr>
    </table>
</table>
</body>
</html>
"@

                if ($sendMail.FromDisplayName) {
                    $mailParams.FromDisplayName = Get-StringValueHC $sendMail.FromDisplayName
                }

                if ($sendMail.Subject) {
                    $mailParams.Subject = '{0}, {1}' -f
                    $mailParams.Subject, $sendMail.Subject
                }

                if ($sendMail.To) {
                    $mailParams.To = $sendMail.To
                }

                if ($sendMail.Bcc) {
                    $mailParams.Bcc = $sendMail.Bcc
                }

                if ($counter.Total.Errors) {
                    $mailParams.Priority = 'High'
                    $mailParams.Subject = '{0} error{1}, {2}' -f
                    $counter.Total.Errors,
                    $(if ($counter.Total.Errors -ne 1) { 's' }),
                    $mailParams.Subject
                }

                if ($allLogFilePaths) {
                    $mailParams.Attachments = $allLogFilePaths |
                    Sort-Object -Unique
                }

                if ($sendMail.Smtp.ConnectionType) {
                    $mailParams.SmtpConnectionType = Get-StringValueHC $sendMail.Smtp.ConnectionType
                }

                #region Create SMTP credential
                $smtpUserName = Get-StringValueHC $sendMail.Smtp.UserName
                $smtpPassword = Get-StringValueHC $sendMail.Smtp.Password

                if ( $smtpUserName -and $smtpPassword) {
                    try {
                        $securePassword = ConvertTo-SecureString -String $smtpPassword -AsPlainText -Force

                        $credential = New-Object System.Management.Automation.PSCredential($smtpUserName, $securePassword)

                        $mailParams.Credential = $credential
                    }
                    catch {
                        throw "Failed to create credential: $_"
                    }
                }
                elseif ($smtpUserName -or $smtpPassword) {
                    throw "Both 'Settings.SendMail.Smtp.Username' and 'Settings.SendMail.Smtp.Password' are required when authentication is needed."
                }
                #endregion

                Send-MailKitMessageHC @mailParams
            }
        }
        catch {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed sending email: $_"
                }
            )

            Write-Warning $systemErrors[-1].Message

            if ($baseLogName -and $isLog.systemErrors) {
                $params = @{
                    DataToExport   = $systemErrors[-1]
                    PartialPath    = "$baseLogName - Errors"
                    FileExtensions = $logFileExtensions
                }
                $null = Out-LogFileHC @params -EA Ignore
            }
        }
        #endregion
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