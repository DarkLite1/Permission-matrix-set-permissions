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
    function ConvertTo-HtmlValueHC {
        param(
            [Parameter(Mandatory)]
            $ErrorObj,
            [Parameter(Mandatory)]
            $SettingId,
            [Parameter(Mandatory)]
            [string]$LogFolderPath
        )

        if (-not $ErrorObj.Value) {
            return $null
        }
        elseif (
            ($ErrorObj.Value.Count -le 5) -and 
            (-not ($ErrorObj.Value -is [hashtable]))
        ) {
            return '<ul>{0}</ul>' -f $(@($ErrorObj.Value).ForEach({ "<li>$_</li>" }))
        }
        else {
            $safeName = "ID $SettingId - $($ErrorObj.Type) - $($ErrorObj.Name).txt".Split([IO.Path]::GetInvalidFileNameChars()) -join '_'

            $OutParams = @{
                LiteralPath = Join-Path -Path $LogFolderPath -ChildPath $safeName
                Encoding    = 'utf8'
                NoClobber   = $true
            }
            $ErrorObj | ConvertTo-Json -Depth 100 | ForEach-Object {
                [System.Text.RegularExpressions.Regex]::Unescape($_)
            } | Out-File @OutParams

            return '<ul><li><a href="{0}">{1} items</a></li></ul>' -f $OutParams.LiteralPath, $ErrorObj.Value.Count
        }
    }
    function ConvertTo-StructuredObjectHC {
        <#
        .SYNOPSIS
            Normalizes various input types into a standard PSCustomObject for 
            HealthCheck reports.
        
        .DESCRIPTION
            This function takes strings, hashtables, or existing objects and 
            ensures they conform to a specific schema (Name, Description, Type, 
            Value). If properties are missing or  null, it injects "Missing 
            data" and sets the Type to "FatalError".
        
        .PARAMETER Objects
            The input data to be converted. Can be a single item or an array.
        
        .EXAMPLE
            "System Error" | ConvertTo-StructuredObjectHC

            Converts a simple string into a full object with Name: 
            'Error during execution'.
        
        .EXAMPLE
            ConvertTo-StructuredObjectHC -Objects @{ 
                Name = "Disk Check"
                Type = "Warning" 
            }
            Converts a hashtable into an object and fills in the missing 
            'Description' property.
        #>

        [CmdletBinding()]
        param(
            [Parameter(ValueFromPipeline = $true)]
            [AllowEmptyCollection()]
            [array]$Objects
        )

        process {
            foreach ($checkObj in @($Objects)) {
                if ($null -eq $checkObj) { continue }

                # 1. Normalize into a PSCustomObject
                $current = if (
                    $checkObj -is [string] -or 
                    $checkObj -is [System.ValueType]
                ) {
                    [PSCustomObject]@{
                        Name        = 'Error during execution'
                        Description = "Primitive value received: $checkObj"
                        Type        = 'FatalError'
                        Value       = $checkObj
                    }
                }
                elseif ($checkObj -is [hashtable]) {
                    [PSCustomObject]$checkObj
                }
                else {
                    # Force a cast to ensure the object is extensible/malleable
                    [PSCustomObject]$checkObj
                }

                # 2. Ensure the 'Value' property exists (avoiding null reference errors later)
                if (-not (Get-Member -InputObject $current -Name 'Value')) {
                    $current | Add-Member -MemberType NoteProperty -Name 'Value' -Value $null
                }

                # 3. Validate Core Properties
                foreach ($prop in @('Name', 'Description', 'Type')) {
                    if ([string]::IsNullOrWhiteSpace($current.$prop)) {
                        $current | Add-Member -MemberType NoteProperty -Name $prop -Value 'Missing data' -Force
                        $current | Add-Member -MemberType NoteProperty -Name 'Type' -Value 'FatalError' -Force
                    }
                }

                # Output the object to the pipeline
                $current
            }
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
    function Get-DatedLogFolderPathHC {
        try {
            $datedLogFolder = Join-Path -Path $LogFolder -ChildPath (
                '{0:00}_{1:00}_{2:00}_{3:00}{4:00}{5:00} ({6})' -f $scriptStartTime.Year, $scriptStartTime.Month,
                $scriptStartTime.Day,
                $scriptStartTime.Hour, $scriptStartTime.Minute, $scriptStartTime.Second, $jsonFileItem.BaseName
            )

            return (New-Item -ItemType 'Directory' -Path $datedLogFolder -Force -EA Stop).FullName
        }
        catch {
            return $LogFolder
        }
    }
    function Out-LogFileHC {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [PSCustomObject[]]$DataToExport,
            [Parameter(Mandatory)]
            [String]$PartialPath,
            [Parameter(Mandatory)]
            [String[]]$FileExtensions,
            [hashtable]$ExcelFile = @{
                SheetName = 'Overview'
                TableName = 'Overview'
                CellStyle = $null
            },
            [Switch]$Append
        )

        $allLogFilePaths = @()

        foreach (
            $fileExtension in
            $FileExtensions | Sort-Object -Unique
        ) {
            try {
                $logFilePath = "$PartialPath$fileExtension"

                Write-Verbose "Export $($DataToExport.Count) object(s) to '$logFilePath'"

                switch ($fileExtension) {
                    '.csv' {
                        $params = @{
                            LiteralPath       = $logFilePath
                            Append            = $Append
                            Delimiter         = ';'
                            NoTypeInformation = $true
                        }
                        $DataToExport | Export-Csv @params

                        break
                    }
                    '.json' {
                        #region Convert error object to error message string
                        $convertedDataToExport = foreach (
                            $exportObject in
                            $DataToExport
                        ) {
                            foreach ($property in $exportObject.PSObject.Properties) {
                                $name = $property.Name
                                $value = $property.Value
                                if (
                                    $value -is [System.Management.Automation.ErrorRecord]
                                ) {
                                    if (
                                        $value.Exception -and $value.Exception.Message
                                    ) {
                                        $exportObject.$name = $value.Exception.Message
                                    }
                                    else {
                                        $exportObject.$name = $value.ToString()
                                    }
                                }
                            }
                            $exportObject
                        }
                        #endregion

                        if (
                            $Append -and
                            (Test-Path -LiteralPath $logFilePath -PathType Leaf)
                        ) {
                            $params = @{
                                LiteralPath = $logFilePath
                                Raw         = $true
                                Encoding    = 'UTF8'
                            }
                            $jsonFileContent = Get-Content @params | ConvertFrom-Json

                            $convertedDataToExport = [array]$convertedDataToExport + [array]$jsonFileContent
                        }

                        $convertedDataToExport |
                        ConvertTo-Json -Depth 7 |
                        Out-File -LiteralPath $logFilePath

                        break
                    }
                    '.txt' {
                        $params = @{
                            LiteralPath = $logFilePath
                            Append      = $Append
                        }

                        $DataToExport | Format-List -Property * -Force |
                        Out-File @params

                        break
                    }
                    '.xlsx' {
                        if (
                            (-not $Append) -and
                            (Test-Path -LiteralPath $logFilePath -PathType Leaf)
                        ) {
                            $logFilePath | Remove-Item
                        }

                        $excelParams = @{
                            Path          = $logFilePath
                            Append        = $true
                            AutoNameRange = $true
                            AutoSize      = $true
                            FreezeTopRow  = $true
                            WorksheetName = $ExcelFile.SheetName
                            TableName     = $ExcelFile.TableName
                            Verbose       = $false
                        }
                        if ($ExcelFile.CellStyle) {
                            $excelParams.CellStyleSB = $ExcelFile.CellStyle
                        }
                        $DataToExport | Export-Excel @excelParams

                        break
                    }
                    default {
                        throw "Log file extension '$_' not supported. Supported values are '.csv', '.json', '.txt' or '.xlsx'."
                    }
                }

                $allLogFilePaths += $logFilePath
            }
            catch {
                Write-Warning "Failed creating log file '$logFilePath': $_"
            }
        }

        $allLogFilePaths
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
                $totalSizeAttachments = 0

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
    function Validate-JsonSchema {
        param(
            [Parameter(Mandatory)]
            [object]$JsonObject
        )

        $errors = [System.Collections.Generic.List[object]]::new()

        function Add-SchemaError {
            param([string]$Message)
            $errors.Add(
                [PSCustomObject]@{ 
                    Message = $Message 
                }
            )
        }

        # --- 1. Required top-level objects ---
        foreach (
            $prop in 
            @('Matrix', 'Export', 'ServiceNow', 'MaxConcurrent', 'PSSessionConfiguration', 'Settings')
        ) {
            if ($null -eq $JsonObject.$prop) {
                Add-SchemaError "Property '$prop' not found"
            }
        }

        # If Settings missing, bail out (prevents deeper checks)
        if ($null -eq $JsonObject.Settings) {
            return $errors
        }

        # --- 2. Validate Settings structure ---
        if ($null -eq $JsonObject.Settings.SaveLogFiles.Where.Folder) {
            Add-SchemaError "Property 'Settings.SaveLogFiles.Where.Folder' not found"
        }

        if ($null -eq $JsonObject.Settings.SaveLogFiles.Detailed) {
            Add-SchemaError "Property 'Settings.SaveLogFiles.Detailed' not found"
        }
        elseif ($JsonObject.Settings.SaveLogFiles.Detailed -isnot [bool] ) {
            Add-SchemaError 'Settings.SaveLogFiles.Detailed must be boolean'
        }

        if ($null -eq $JsonObject.Settings.SendMail) {
            Add-SchemaError "Property 'Settings.SendMail' not found"
        }
        else {
            if (-not $JsonObject.Settings.SendMail.From) {
                Add-SchemaError "Property 'Settings.SendMail.From' not found"
            }

            if ($JsonObject.Settings.SendMail.To -and
                ($JsonObject.Settings.SendMail.To -isnot [string] -and
                $JsonObject.Settings.SendMail.To -isnot [array])) {
                Add-SchemaError "Property 'Settings.SendMail.To' not found"
            }

            if ($null -eq $JsonObject.Settings.SendMail.Body) {
                Add-SchemaError "Property 'Settings.SendMail.Body' not found"
            }
        }

        # --- 3. Validate Matrix structure ---
        if ($null -ne $JsonObject.Matrix) {
            if (-not $JsonObject.Matrix.FolderPath) {
                Add-SchemaError "Property 'Matrix.FolderPath' not found"
            }

            if (-not $JsonObject.Matrix.DefaultsFile) {
                Add-SchemaError "Property 'Matrix.DefaultsFile' not found"
            }

            if ($JsonObject.Matrix.ExcludedSamAccountName -and
                $JsonObject.Matrix.ExcludedSamAccountName -isnot [array]) {
                Add-SchemaError "Property 'Matrix.ExcludedSamAccountName' must be an array"
            }

            if ($null -eq $JsonObject.Matrix.Archive) {
                Add-SchemaError "Property 'Matrix.Archive' not found"
            }
            elseif ($JsonObject.Matrix.Archive -isnot [bool]) {
                Add-SchemaError 'Matrix.Archive must be boolean'
            }
        }
        else {
            Add-SchemaError 'Matrix required'
        }

        # --- 4. Validate MaxConcurrent structure ---
        if ($null -ne $JsonObject.MaxConcurrent) {
            foreach (
                $prop in
                @('Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer')
            ) {
                $val = $JsonObject.MaxConcurrent.$prop

                if ($null -eq $val) {
                    Add-SchemaError "Property 'MaxConcurrent.$prop' not found"
                    continue
                }

                if ($val -notmatch '^\d+$') {
                    Add-SchemaError "MaxConcurrent.$prop must be an integer"
                }
            }
        }
     
        # --- 5. Validate Export structure (file extensions only) ---
        if ($null -ne $JsonObject.Export) {
            if ($JsonObject.Export.PermissionsExcelFile -and
                ($JsonObject.Export.PermissionsExcelFile -isnot [string]) -and
                ($JsonObject.Export.PermissionsExcelFile -notmatch '\.xlsx$')) {
                Add-SchemaError 'Export.PermissionsExcelFile must be a string ending with .xlsx'
            }

            if ($JsonObject.Export.OverviewHtmlFile -and
                ($JsonObject.Export.OverviewHtmlFile -isnot [string]) -and
                ($JsonObject.Export.OverviewHtmlFile -notmatch '\.html?$')) {
                Add-SchemaError 'Export.OverviewHtmlFile must be a string ending with .html'
            }

            if ($JsonObject.Export.ServiceNowFormDataExcelFile -and
                ($JsonObject.Export.ServiceNowFormDataExcelFile -isnot [string]) -and
                ($JsonObject.Export.ServiceNowFormDataExcelFile -notmatch '\.xlsx$')) {
                Add-SchemaError 'Export.ServiceNowFormDataExcelFile must be a string ending with .xlsx'
            }

            if ($JsonObject.Export.ServiceNowFormDataExcelFile -and
                -not $JsonObject.ServiceNow) {
                Add-SchemaError 'ServiceNow must be defined when ServiceNowFormDataExcelFile is used'
            }
        }
        else {
            Add-SchemaError 'Export required'
        }

        return $errors
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
            if ([System.Diagnostics.EventLog]::SourceExists($Source)) {
                $existingLogName = [System.Diagnostics.EventLog]::LogNameFromSourceName($Source, '.')

                if ($existingLogName -ne $LogName) {
                    throw "The event log source '$Source' is already registered with event log name '$existingLogName', it cannot be used with log name '$LogName'."
                }
            }
            else {
                Write-Verbose "Create event log source '$Source' with log name '$LogName'"

                New-EventLog -LogName $LogName -Source $Source -EA Stop
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
    $ExcludedSamAccountName = $Matrix.ExcludedSamAccountName
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
                        ExcludeSamAccountName = $ExcludedSamAccountName
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
                        ExcludedSamAccountName = $ExcludedSamAccountName
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

            #region Remove group members that are in the ExcludedSamAccountName
            if ($ExcludedSamAccountName) {
                $allAdObjects = @($ADObjectDetails) + @($groupManagersAdDetails)

                foreach ($adObject in $allAdObjects) {
                    if (-not $adObject.adGroupMember) { continue }

                    $adObject.adGroupMember = @(
                        $adObject.adGroupMember.Where(
                            { $_.SamAccountName -notin $ExcludedSamAccountName }
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
    function Ensure-SafeSettingsHC {
        param([object]$Settings)

        if (-not $Settings) {
            # Create a safe dummy object with defaults so END block cannot break
            return [PSCustomObject]@{
                ScriptName     = 'Default script name'
                SaveLogFiles   = @{
                    Where               = @{
                        Folder = $null 
                    } 
                    Detailed            = $false 
                    DeleteLogsAfterDays = 0 
                }
                SaveInEventLog = @{ 
                    Save    = $false 
                    LogName = $null 
                }
                SendMail       = @{ 
                    From = $null 
                    To   = @()
                    Body = $null 
                    Smtp = @{ Port     = 25
                        ConnectionType = 'None' 
                    } 
                }
            }
        }

        # Ensure ScriptName exists and is safe for filenames
        if ([string]::IsNullOrWhiteSpace($Settings.ScriptName)) {
            $Settings.ScriptName = 'Default script name'
        }

        if (-not $Settings.SaveLogFiles) {
            $Settings | Add-Member -NotePropertyName SaveLogFiles -NotePropertyValue @{
                Where               = @{ 
                    Folder = $null 
                }
                Detailed            = $false
                DeleteLogsAfterDays = 0
            }
        }

        if (-not $Settings.SaveInEventLog) {
            $Settings | Add-Member -NotePropertyName SaveInEventLog -NotePropertyValue @{
                Save    = $false
                LogName = $null
            }
        }

        if (-not $Settings.SendMail) {
            $Settings | Add-Member -NotePropertyName SendMail -NotePropertyValue @{
                From = $null
                To   = @()
                Body = $null
                Smtp = @{
                    Port           = 25
                    ConnectionType = 'None'
                }
            }
        }

        return $Settings
    }

    function Ensure-LogFolderHC {
        param(
            [Parameter()]
            [string]$RequestedFolder,

            [Parameter()]
            [ref]$SystemErrors
        )

        #
        # 1 - If requested folder is null or empty → immediate fallback
        #
        if ([string]::IsNullOrWhiteSpace($RequestedFolder)) {

            $fallback = Join-Path $env:TEMP 'PermissionMatrixLogs'

            try {
                if (-not (Test-Path -LiteralPath $fallback -PathType Container)) {
                    New-Item -ItemType Directory -Path $fallback -ErrorAction Stop | Out-Null
                }
            }
            catch {
                # Last-resort fallback (extremely rare)
                $fallback = $env:TEMP
            }

            if ($SystemErrors) {
                $SystemErrors.Value.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "LogFolder missing. Using fallback folder '$fallback'."
                    }
                )
            }

            return $fallback
        }

        #
        # 2 - Try to create/verify the requested folder
        #
        try {
            if (-not (Test-Path -LiteralPath $RequestedFolder -PathType Container)) {
                New-Item -ItemType Directory -Path $RequestedFolder -ErrorAction Stop | Out-Null
            }

            return $RequestedFolder
        }
        catch {

            #
            # 3 - Requested folder invalid → use fallback
            #
            $fallback = Join-Path $env:TEMP 'PermissionMatrixLogs'

            try {
                if (-not (Test-Path -LiteralPath $fallback -PathType Container)) {
                    New-Item -ItemType Directory -Path $fallback -ErrorAction Stop | Out-Null
                }
            }
            catch {
                # last resort (TEMP folder is always safe)
                $fallback = $env:TEMP
            }

            if ($SystemErrors) {
                $SystemErrors.Value.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "LogFolder '$RequestedFolder' invalid or uncreatable. Using fallback '$fallback'. Error: $_"
                    }
                )
            }

            return $fallback
        }
    }

    function Remove-FileHC {
        param(
            [Parameter(Mandatory)]
            [string]$FilePath,
            [Parameter()]
            [ref]$SystemErrors
        )

        try {
            # If file doesn't exist, nothing to do
            if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
                return
            }

            # Remove file safely
            Remove-Item -LiteralPath $FilePath -Force -ErrorAction Stop
        }
        catch {
            # Only record errors if SystemErrors was passed in
            if ($SystemErrors) {
                $SystemErrors.Value.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "Remove-FileHC failed for '$FilePath': $_"
                    }
                )
            }
            else {
                Write-Warning "Remove-FileHC failed for '$FilePath': $_"
            }
        }
    }

    function New-HtmlCheckRow {
        param(
            [Parameter(Mandatory)]
            [object]$CheckObject
        )

        # Determine CSS class based on type (Error / Warning / Info)
        $cssClass = Get-HtmlClassProbTypeHC -Name $CheckObject.Type

        # HTML-encode dynamic fields
        $name = [System.Net.WebUtility]::HtmlEncode($CheckObject.Name)
        $desc = [System.Net.WebUtility]::HtmlEncode($CheckObject.Description)

        # Optional list of values
        $listHtml = Format-HtmlList -Value $CheckObject.Value

        # Output final row
        return @"
<tr>
    <td class="$cssClass"></td>
    <td colspan="7">
        <p class="probTitle">$name</p>
        <p>$desc</p>
        $listHtml
    </td>
</tr>
"@
    }

    function Get-HtmlClassProbTypeHC {
        [OutputType([string])]
        param (
            [Parameter(Mandatory)]
            [ValidateSet('FatalError', 'Warning', 'Information')]
            [string]$Name
        )

        switch ($Name) {
            'FatalError' { return 'probTypeError' }
            'Warning' { return 'probTypeWarning' }
            'Information' { return 'probTypeInfo' }
        }
    }

    function Validate-Settings {
        param(
            [object]$Settings,
            [object]$Matrix,
            [object]$Export,
            [object]$ServiceNow,
            [object]$MaxConcurrent
        )
        $errors = [System.Collections.Generic.List[object]]::new()

        function Add-Error {
            param(
                [string]$Type,
                [string]$Name,
                [string]$Description
            )
                
            $errors.Add(
                [PSCustomObject]@{ 
                    Type        = $Type 
                    Name        = $Name 
                    Description = $Description 
                }
            )
        }

        # ---------------------------
        # 1. Base Settings Validation
        # ---------------------------
        if (-not $Settings) { 
            Add-Error 'FatalError' 'Invalid configuration' "Property 'Settings' missing from JSON."
            return [PSCustomObject]@{ IsValid = $false; Errors = $errors; Settings = $null }
        }

        if ([string]::IsNullOrWhiteSpace($Settings.ScriptName)) {
            Add-Error 'Warning' 'Missing Script Name' "No 'Settings.ScriptName' found in JSON. A default name will be used."
            $Settings | Add-Member -NotePropertyName ScriptName -NotePropertyValue 'Default script name' -Force
        }

        if ($Settings.SaveLogFiles.Detailed -isnot [bool]) {
            Add-Error 'FatalError' 'Invalid type' 'Settings.SaveLogFiles.Detailed must be a boolean.'
        }

        if ($Settings.SaveInEventLog.Save -isnot [bool]) {
            Add-Error 'FatalError' 'Invalid type' 'Settings.SaveInEventLog.Save must be a boolean.'
        }

        if ([string]::IsNullOrWhiteSpace($Settings.SendMail.From)) {
            Add-Error 'FatalError' 'Invalid configuration' 'Settings.SendMail.From cannot be empty.'
        }

        if (-not $Settings.SendMail.To) {
            Add-Error 'FatalError' 'Invalid configuration' 'Settings.SendMail.To cannot be empty.'
        }
        elseif ($Settings.SendMail.To -isnot [array] -and $Settings.SendMail.To -isnot [string]) {
            Add-Error 'FatalError' 'Invalid type' 'Settings.SendMail.To must be an array or a string.'
        }
        
        if ([string]::IsNullOrWhiteSpace($Settings.SendMail.Body)) {
            Add-Error 'FatalError' 'Invalid configuration' 'Settings.SendMail.Body cannot be empty.'
        }
        
        if (-not $Settings.SendMail.Smtp.Port -or $Settings.SendMail.Smtp.Port -notmatch '^\d+$') {
            Add-Error 'FatalError' 'Invalid configuration' 'Settings.SendMail.Smtp.Port must be an integer.'
        }
        
        $validConnections = @('None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable')
        if ($Settings.SendMail.Smtp.ConnectionType -notin $validConnections) {
            Add-Error 'FatalError' 'Invalid configuration' "Settings.SendMail.Smtp.ConnectionType must be one of: $($validConnections -join ', ')."
        }

        # ---------------------------
        # 2. Matrix Validation
        # ---------------------------
        if ($Matrix) {
            if (-not (Test-Path -LiteralPath $Matrix.DefaultsFile -PathType Leaf)) {
                Add-Error 'FatalError' 'Invalid path' "Matrix.DefaultsFile '$($Matrix.DefaultsFile)' does not exist or is not a file."
            }
            
            # SCHEMA CHECK ONLY: Ensure the property is populated before the BEGIN block attempts to map it
            if ([string]::IsNullOrWhiteSpace($Matrix.FolderPath)) {
                Add-Error 'FatalError' 'Invalid configuration' 'Matrix.FolderPath cannot be empty.'
            }
            
            if ($Matrix.ExcludedSamAccountName -isnot [array]) {
                Add-Error 'FatalError' 'Invalid type' 'Matrix.ExcludedSamAccountName must be an array.'
            }
        }
        else {
            Add-Error 'FatalError' 'Invalid configuration' "Property 'Matrix' missing."
        }

        # ---------------------------
        # 3. MaxConcurrent Validation
        # ---------------------------
        if ($MaxConcurrent) {
            foreach ($prop in @('Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer')) {
                if (-not $MaxConcurrent.$prop -or $MaxConcurrent.$prop -notmatch '^\d+$') {
                    Add-Error 'FatalError' 'Invalid type' "MaxConcurrent.$prop must be an integer."
                }
            }
        }
        else {
            Add-Error 'FatalError' 'Invalid configuration' "Property 'MaxConcurrent' missing."
        }

        # ---------------------------
        # 4. Export & ServiceNow
        # ---------------------------
        if ($Export) {
            if (-not [string]::IsNullOrWhiteSpace($Export.PermissionsExcelFile) -and $Export.PermissionsExcelFile -notmatch '\.xlsx$') {
                Add-Error 'FatalError' 'Invalid path' 'Export.PermissionsExcelFile must end in .xlsx.'
            }
            if (-not [string]::IsNullOrWhiteSpace($Export.OverviewHtmlFile) -and $Export.OverviewHtmlFile -notmatch '\.html?$') {
                Add-Error 'FatalError' 'Invalid path' 'Export.OverviewHtmlFile must end in .html.'
            }
            if (-not [string]::IsNullOrWhiteSpace($Export.ServiceNowFormDataExcelFile)) {
                if ($Export.ServiceNowFormDataExcelFile -notmatch '\.xlsx$') {
                    Add-Error 'FatalError' 'Invalid path' 'Export.ServiceNowFormDataExcelFile must end in .xlsx.'
                }
                if (-not $ServiceNow) {
                    Add-Error 'FatalError' 'Invalid configuration' 'ServiceNow configuration object is required when Export.ServiceNowFormDataExcelFile is populated.'
                }
                else {
                    if ([string]::IsNullOrWhiteSpace($ServiceNow.CredentialsFilePath)) { 
                        Add-Error 'FatalError' 'Invalid configuration' 'ServiceNow.CredentialsFilePath is required.' 
                    }
                    if ([string]::IsNullOrWhiteSpace($ServiceNow.TableName)) { 
                        Add-Error 'FatalError' 'Invalid configuration' 'ServiceNow.TableName is required.' 
                    }
                    if ([string]::IsNullOrWhiteSpace($ServiceNow.Environment)) { 
                        Add-Error 'FatalError' 'Invalid configuration' 'ServiceNow.Environment is required.' 
                    }
                }
            }
        }

        return [PSCustomObject]@{
            Settings   = $Settings
            Matrix     = $Matrix
            Export     = $Export
            ServiceNow = $ServiceNow
            IsValid    = ($errors.Where({ $_.Type -eq 'FatalError' }).Count -eq 0)
            Errors     = $errors
        }
    }

    function Initialize-HtmlStructure {
        $style = @'
        <style type="text/css">
            a { color: black; text-decoration: underline; }
            a:hover { color: blue; }
            body { font-family:verdana; font-size:14px; background-color:white; }
            h1, h2, h3 { margin-bottom: 0; }
            p.italic { font-style: italic; font-size: 12px; }
            table { border-collapse:collapse; border:0px none; padding:3px; text-align:left; }
            td, th { border-collapse:collapse; border:1px none; padding:3px; text-align:left; }
            .matrixTable { border: 1px solid Black; border-collapse: separate; border-spacing: 0px 0.6em; width: 600px; }
            .matrixTitle { border: none; background-color: lightgrey; text-align: center; padding: 6px; }
            .matrixHeader { font-weight: normal; letter-spacing: 5pt; font-style: italic; }
            .matrixFileInfo { font-weight: normal; font-size: 12px; font-style: italic; text-align: center; }
            .legendTable { border-collapse: collapse; border: 1px solid Black; table-layout: fixed; }
            .legendTable td { text-align: center; }
            .probTitle { font-weight: bold; }
            .probTypeWarning { background-color: orange; }
            .probTextWarning { color: orange; font-weight: bold; }
            .probTypeError { background-color: red; }
            .probTextError { color: red; font-weight: bold; }
            .probTypeInfo { background-color: lightgrey; }
            table tbody tr td a { display: block; width: 100%; height: 100%; }
            .aboutTable th, .aboutTable td { color: rgb(143, 140, 140); font-weight: normal; }
            base { target="_blank" }
        </style>
'@
        $troubleshootingStyle = @'
        <style type="text/css">
            body { margin: 20px; }
        </style>
'@
        return @{
            Style                = $style
            TroubleshootingStyle = $troubleshootingStyle
            Templates            = @{
                SettingsHeader = '<th class="matrixHeader" colspan="8">Settings</th><tr><td></td><td>ID</td><td>ComputerName</td><td>Path</td><td>Action</td><td>Duration</td></tr>'
                LegendTable    = '<table class="legendTable"><tr><td class="probTypeError" style="width:150px;">Error</td><td class="probTypeWarning" style="width:150px;">Warning</td><td class="probTypeInfo" style="width:150px;">Information</td></tr></table>'
            }
        }
    }

    function Format-HtmlList {
        param([object]$Value)
        if (-not $Value) { return '' }
        if ($Value.Count -le 5 -and $Value -isnot [hashtable]) {
            $encodedItems = @($Value).ForEach(
                { "<li>$([System.Net.WebUtility]::HtmlEncode($_))</li>" }
            ) -join ''
            return "<ul>$encodedItems</ul>"
        }
        return '<p><i>Check JSON dump for multiple items.</i></p>'
    }

    function Process-MatrixObjects {
        param(
            [Parameter(Mandatory)][array]$ImportedMatrix,
            [Parameter(Mandatory)][object]$Html
        )

        #
        # Process each matrix item:
        #   - Generate its troubleshooting log
        #   - Attach TroubleshootingLogPath property
        #

        foreach (
            $matrixItem in
            $ImportedMatrix | Sort-Object { $_.File.Item.Name }
        ) {

            $logPath = $null

            try {
                $logPath = Write-MatrixTroubleshootingLog `
                    -Matrix $matrixItem `
                    -Html $Html
            }
            catch {
                Write-Warning "Failed to build troubleshooting log for '$($matrixItem.File.Item.Name)': $_"
            }

            #
            # Add or update TroubleshootingLogPath on the matrix item
            #
            $matrixItem |
            Add-Member -NotePropertyName TroubleshootingLogPath `
                -NotePropertyValue $logPath `
                -Force
        }

        return $ImportedMatrix
    }

    function Write-MatrixTroubleshootingLog {
        param(
            [Parameter(Mandatory)][object]$Matrix,
            [Parameter(Mandatory)][object]$Html
        )

        try {
            #
            # Validate log folder
            #
            $logFolder = $Matrix.File.LogFolder
            if (-not (Test-Path -LiteralPath $logFolder -PathType Container)) {
                return $null
            }

            #
            # File metadata (safe encoded)
            #
            $fileName = [System.Net.WebUtility]::HtmlEncode($Matrix.File.Item.Name)

            $modifiedBy = $Matrix.File.ExcelInfo.LastModifiedBy
            $modBy = if ($modifiedBy) {
                [System.Net.WebUtility]::HtmlEncode($modifiedBy.ToString().Trim())
            }
            else {
                'Unknown'
            }

            $modifiedTime = $Matrix.File.ExcelInfo.Modified
            $modDate = if ($modifiedTime -is [datetime]) {
                $modifiedTime.ToString('dd/MM/yyyy HH:mm:ss')
            }
            elseif ($modifiedTime) {
                [System.Net.WebUtility]::HtmlEncode("$modifiedTime")
            }
            else {
                'Unknown'
            }

            #
            # Function: Render a section (File / FormData / Permissions)
            #
            function New-SectionHtml {
                param(
                    [string]$SectionName,
                    [object]$Checks
                )

                if (-not $Checks) { return '' }

                $rows = ($Checks | ConvertTo-StructuredObjectHC | ForEach-Object {
                        New-HtmlCheckRow -CheckObject $_
                    }) -join ''

                return "<tr><th class='matrixHeader' colspan='8'>$SectionName</th></tr>$rows"
            }

            #
            # Build 3 main sections
            #
            $sectionsHtml = @(
                New-SectionHtml -SectionName 'File' -Checks $Matrix.File.Check
                New-SectionHtml -SectionName 'FormData' -Checks $Matrix.FormData.Check
                New-SectionHtml -SectionName 'Permissions' -Checks $Matrix.Permissions.Check
            ) -join ''

            #
            # Settings checks (if any)
            #
            $settingsHtml = ''

            if ($Matrix.Settings) {

                $settingsRows = foreach ($S in $Matrix.Settings | Sort-Object ID) {
                    if (-not $S.Check) { continue }

                    #
                    # Heading row for each setting entry
                    #
                    $encComp = [System.Net.WebUtility]::HtmlEncode($S.Import.ComputerName)
                    $encPath = [System.Net.WebUtility]::HtmlEncode($S.Import.Path)

                    "<tr><td colspan='8' style='background-color:#eee;'>
                    <b>Setting ID: $($S.ID)</b> ($encComp - $encPath)
                 </td></tr>" +

                    #
                    # Individual checks rendered via the shared row builder
                    #
                    (
                        $S.Check | ConvertTo-StructuredObjectHC | ForEach-Object {
                            New-HtmlCheckRow -CheckObject $_
                        }
                    ) -join ''
                }

                if ($settingsRows) {
                    $settingsHtml = "<tr><th class='matrixHeader' colspan='8'>Settings Checks</th></tr>$($settingsRows -join '')"
                }
            }

            #
            # Combine full table
            #
            $tableHtml = @"
<table class="matrixTable" style="width: 100%;">
    <tr><th colspan="8" class="matrixHeader">Troubleshooting details</th></tr>
    <tr><td colspan="8"><strong>Last change:</strong> $modBy @ $modDate</td></tr>
    $sectionsHtml
    $settingsHtml
</table>
<br>
$($Html.Templates.LegendTable)
"@

            #
            # Final HTML document
            #
            $fullHtml = @"
<!DOCTYPE html>
<html>
<head>
    $($Html.Style)
    $($Html.TroubleshootingStyle)
</head>
<body>
    <h1>Troubleshooting Log: $fileName</h1>
    $tableHtml
</body>
</html>
"@

            #
            # Write file
            #
            $filePath = Join-Path -Path $logFolder -ChildPath '00 - Troubleshooting Log.html'
            $fullHtml | Out-File -LiteralPath $filePath -Encoding UTF8 -Force

            return $filePath
        }
        catch {
            Write-Warning "Troubleshooting log failed for '$($Matrix.File.Item.Name)': $_"
            return $null
        }
    }

    function Build-MatrixEmailHtml {
        param(
            [Parameter(Mandatory)][array]$ImportedMatrix,
            [Parameter(Mandatory)][object]$Html
        )

        function New-SectionHtml {
            param([string]$Name, [object]$Checks)

            if (-not $Checks) { return '' }

            $rows = ($Checks | ConvertTo-StructuredObjectHC | ForEach-Object {
                    New-HtmlCheckRow -CheckObject $_
                }) -join ''

            return "<tr><th class='matrixHeader' colspan='8'>$Name</th></tr>$rows"
        }

        function New-SettingsTableHtml {
            param([object]$MatrixItem, [object]$Html)

            $fatalFile = @($MatrixItem.File.Check?.Type) -contains 'FatalError'
            $fatalPerms = @($MatrixItem.Permissions.Check?.Type) -contains 'FatalError'

            if (-not $MatrixItem.Settings -or $fatalFile -or $fatalPerms) {
                return ''    # Suppress table when File or Permissions have fatal errors
            }

            $rows = foreach ($S in $MatrixItem.Settings | Sort-Object ID) {

                $types = @($S.Check?.Type).Where({ $_ })
                if (-not $types) { continue }

                $class = if ($types -contains 'FatalError') { 'probTypeError' }
                elseif ($types -contains 'Warning') { 'probTypeWarning' }
                elseif ($types -contains 'Information') { 'probTypeInfo' }
                else { '' }

                $duration = if ($S.JobTime.Duration) {
                    '{0:00}:{1:00}:{2:00}' -f $S.JobTime.Duration.Hours,
                    $S.JobTime.Duration.Minutes,
                    $S.JobTime.Duration.Seconds
                }
                else {
                    'NA'
                }

                $link = $MatrixItem.TroubleshootingLogPath ?? '#'

                $encComp = [System.Net.WebUtility]::HtmlEncode($S.Import.ComputerName)
                $encPath = [System.Net.WebUtility]::HtmlEncode($S.Import.Path)
                $encAction = [System.Net.WebUtility]::HtmlEncode($S.Import.Action)

                "<tr>
                <td class='$class'></td>
                <td><a href='$link'>$($S.ID)</a></td>
                <td><a href='$link'>$encComp</a></td>
                <td><a href='$link'>$encPath</a></td>
                <td><a href='$link'>$encAction</a></td>
                <td><a href='$link'>$duration</a></td>
             </tr>"
            }

            if (-not $rows) { return '' }

            return $Html.Templates.SettingsHeader + ($rows -join '')
        }

        $resultHtml = ''

        foreach ($Item in $ImportedMatrix | Sort-Object { $_.File.Item.Name }) {

            # Build the 3 built‑in sections
            $sectionHtml = @(
                New-SectionHtml -Name 'File' -Checks $Item.File.Check
                New-SectionHtml -Name 'FormData' -Checks $Item.FormData.Check
                New-SectionHtml -Name 'Permissions' -Checks $Item.Permissions.Check
            ) -join ''

            # Build settings table (only if no fatal file/permissions)
            $settingsHtml = New-SettingsTableHtml -MatrixItem $Item -Html $Html

            # Metadata
            $encFileName = [System.Net.WebUtility]::HtmlEncode($Item.File.Item.Name)

            $modBy = [System.Net.WebUtility]::HtmlEncode(
                $Item.File.ExcelInfo.LastModifiedBy ??
                'Unknown'
            )

            $modDate = $Item.File.ExcelInfo.Modified
            if ($modDate -is [datetime]) {
                $modDate = $modDate.ToString('dd/MM/yyyy HH:mm:ss')
            }
            elseif ($modDate) {
                $modDate = [System.Net.WebUtility]::HtmlEncode("$modDate")
            }
            else {
                $modDate = 'Unknown'
            }

            # Assemble full table
            $resultHtml += @"
<table class="matrixTable">
    <tr>
        <th class="matrixTitle" colspan="8">
            <a href="$($Item.File.SaveFullName)">$encFileName</a>
        </th>
    </tr>
    <tr>
        <th class="matrixFileInfo" colspan="8">
            Last change: $modBy @ $modDate
        </th>
    </tr>
    $sectionHtml
    $settingsHtml
</table>
<br><br>
"@
        }

        return $resultHtml
    }

    function Build-AccessList {
        param(
            [array]$SamAccountNames, [hashtable]$AdObjectHash, [string]$FileName
        )
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($S in $SamAccountNames) {
            $adData = $AdObjectHash[$S]
            if (-not $adData?.adObject) { continue }
            
            if (-not $adData.adGroupMember) {
                $list.Add(
                    [PSCustomObject]@{ 
                        MatrixFileName       = $FileName
                        SamAccountName       = $S
                        Name                 = $adData.adObject.Name
                        Type                 = $adData.adObject.ObjectClass
                        MemberName           = $null 
                        MemberSamAccountName = $null 
                    }
                )
            }
            else {
                foreach ($member in $adData.adGroupMember) {
                    $list.Add(
                        [PSCustomObject]@{
                            MatrixFileName       = $FileName 
                            SamAccountName       = $S 
                            Name                 = $adData.adObject.Name
                            Type                 = $adData.adObject.ObjectClass
                            MemberName           = $member.Name 
                            MemberSamAccountName = $member.SamAccountName 
                        }
                    )
                }
            }
        }
        return $list
    }

    function Build-GroupManagerList {
        param([array]$SamAccountNames, [hashtable]$AdObjectHash, [hashtable]$GroupManagerHash, [string]$FileName)
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($S in $SamAccountNames) {
            $adData = $AdObjectHash[$S]
            if (-not $adData?.adObject -or $adData.adObject.ObjectClass -ne 'group') { continue }
            
            $managedBy = $adData.adObject.PSObject.Properties['ManagedBy']?.Value
            if ([string]::IsNullOrWhiteSpace($managedBy)) { continue }

            $gm = $GroupManagerHash[$managedBy]
            if (-not $gm?.adObject) { 
                $list.Add(
                    [PSCustomObject]@{ 
                        MatrixFileName    = $FileName 
                        GroupName         = $adData.adObject.Name 
                        ManagerName       = $null
                        ManagerType       = $null 
                        ManagerMemberName = $null 
                    }
                )
            }
            elseif (-not $gm.adGroupMember) { 
                $list.Add(
                    [PSCustomObject]@{ 
                        MatrixFileName    = $FileName
                        GroupName         = $adData.adObject.Name 
                        ManagerName       = $gm.adObject.Name 
                        ManagerType       = $gm.adObject.ObjectClass 
                        ManagerMemberName = $null 
                    }
                )
            }
            else { 
                foreach ($user in $gm.adGroupMember) { 
                    $list.Add(
                        [PSCustomObject]@{ 
                            MatrixFileName    = $FileName 
                            GroupName         = $adData.adObject.Name 
                            ManagerName       = $gm.adObject.Name
                            ManagerType       = $gm.adObject.ObjectClass
                            ManagerMemberName = $user.Name 
                        }
                    ) 
                } 
            }
        }
        return $list
    }

    function Build-ExportData {
        param(
            [Parameter(Mandatory)][array]$ImportedMatrix,
            [hashtable]$AdObjectHash,
            [hashtable]$GroupManagerHash
        )

        #
        # Prepare output structure
        #
        $export = @{
            AccessList    = [System.Collections.Generic.List[object]]::new()
            AdObjects     = [System.Collections.Generic.List[object]]::new()
            FormData      = [System.Collections.Generic.List[object]]::new()
            GroupManagers = [System.Collections.Generic.List[object]]::new()
        }

        #
        # Helper: Extract unique SAMs from a matrix
        #
        function Get-UniqueMatrixSamAccountNames {
            param([object]$MatrixItem)

            return $MatrixItem.Settings?.AdObjects?.Values |
            ForEach-Object { "$($_.SamAccountName)".Trim() } |
            Where-Object { $_ } |
            Sort-Object -Unique
        }

        #
        # Helper: Convert a single AdObject entry to export format
        #
        function Convert-AdObjectExport {
            param(
                [string]$MatrixName,
                [object]$Entry
            )

            return [PSCustomObject]@{
                MatrixFileName = $MatrixName
                SamAccountName = $Entry.SamAccountName
                GroupName      = $Entry.Converted.Begin
                SiteCode       = $Entry.Converted.Middle
                Name           = $Entry.Converted.End
            }
        }

        #
        # Main processing loop
        #
        foreach ($Matrix in $ImportedMatrix) {

            $matrixName = $Matrix.File.Item.BaseName
            $uniqueSams = Get-UniqueMatrixSamAccountNames -MatrixItem $Matrix

            #
            # 1. Access list
            #
            $access = Build-AccessList `
                -SamAccountNames $uniqueSams `
                -AdObjectHash $AdObjectHash `
                -FileName $matrixName

            if ($access) { $export.AccessList.AddRange($access) }

            #
            # 2. Group managers
            #
            $gm = Build-GroupManagerList `
                -SamAccountNames $uniqueSams `
                -AdObjectHash $AdObjectHash `
                -GroupManagerHash $GroupManagerHash `
                -FileName $matrixName

            if ($gm) { $export.GroupManagers.AddRange($gm) }

            #
            # 3. AD object (converted values)
            #
            if ($Matrix.Settings?.AdObjects) {

                $adConvertedObjects =
                $Matrix.Settings.AdObjects.GetEnumerator() |
                ForEach-Object {
                    Convert-AdObjectExport -MatrixName $matrixName -Entry $_.Value
                } |
                Group-Object SamAccountName |
                ForEach-Object { $_.Group[0] }   # ensure unique

                if ($adConvertedObjects) {
                    $export.AdObjects.AddRange($adConvertedObjects)
                }
            }

            #
            # 4. FormData (raw source)
            #
            if ($Matrix.FormData?.Import) {
                $export.FormData.AddRange($Matrix.FormData.Import)
            }
        }

        return $export
    }

    function Export-PermissionsFile {
        param(
            [Parameter(Mandatory)][object]$DataToExport,
            [Parameter(Mandatory)][string]$OutputPath,
            [Parameter(Mandatory)][string]$LogFolder,
            [Parameter(Mandatory)][ref]$SystemErrors
        )

        try {
            #
            # 1. Remove existing output files
            #
            Remove-FileHC -FilePath $OutputPath

            $logTempPath = Join-Path $LogFolder 'Permissions.xlsx'
            Remove-FileHC -FilePath $logTempPath

            #
            # 2. Export each collection into its corresponding worksheet
            #
            foreach ($entry in $DataToExport.GetEnumerator()) {

                if (-not $entry.Value) {
                    continue
                }

                $params = @{
                    Path          = $logTempPath
                    WorksheetName = $entry.Name
                    TableName     = $entry.Name
                    AutoSize      = $true
                    FreezeTopRow  = $true
                }

                try {
                    $entry.Value | Export-Excel @params
                }
                catch {
                    $SystemErrors.Value.Add(
                        [PSCustomObject]@{
                            DateTime = Get-Date
                            Message  = "Export-PermissionsFile failed for sheet '$($entry.Name)': $_"
                        }
                    )
                }
            }

            #
            # 3. Copy the final result to the target output path
            #
            if (Test-Path -LiteralPath $logTempPath -PathType Leaf) {
                Copy-Item -LiteralPath $logTempPath -Destination $OutputPath -Force
            }
        }
        catch {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Export-PermissionsFile failed: $_"
                }
            )
        }
    }

    function Upload-ServiceNowFormData {
        param(
            [Parameter(Mandatory)][string]$OutputPath,
            [Parameter(Mandatory)][object]$ServiceNowConfig,
            [Parameter(Mandatory)][object]$ScriptPathItem,
            [Parameter(Mandatory)][ref]$SystemErrors
        )

        try {
            #
            # 1. Validate required ServiceNow parameters
            #
            $credPath = $ServiceNowConfig.CredentialsFilePath
            $env = $ServiceNowConfig.Environment
            $table = $ServiceNowConfig.TableName

            if (-not $credPath -or -not $env -or -not $table) {
                return   # Silent skip 
            }

            #
            # 2. Execute uploader script
            #
            & $ScriptPathItem.UpdateServiceNow `
                -CredentialsFilePath $credPath `
                -Environment $env `
                -TableName $table `
                -FormDataExcelFilePath $OutputPath `
                -ExcelFileWorksheetName 'SnowFormData'
        }
        catch {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Upload-ServiceNowFormData failed: $_"
                }
            )
        }
    }

    function Export-ServiceNowFormData {
        param(
            [Parameter(Mandatory)][object]$DataToExport,
            [Parameter(Mandatory)][string]$OutputPath,
            [Parameter(Mandatory)][string]$ExportLogFolder,
            [Parameter(Mandatory)][ref]$SystemErrors
        )

        try {
            #
            # 1. Remove existing output file (if any)
            #
            Remove-FileHC -FilePath $OutputPath

            #
            # 2. Build a lookup of FormData keyed by MatrixFileName
            #
            $formDataHash = @{}
            foreach ($fd in $DataToExport.FormData) {
                if ($fd.MatrixFileName) {
                    $formDataHash[$fd.MatrixFileName] = $fd
                }
            }

            #
            # 3. Build ServiceNow export rows
            #
            $serviceNowRows = foreach ($adObj in $DataToExport.AdObjects) {

                $fd = $formDataHash[$adObj.MatrixFileName]

                if (
                    $fd -and
                    $fd.MatrixFormStatus -eq 'Enabled'
                ) {
                    [PSCustomObject]@{
                        u_matrixfilename        = $adObj.MatrixFileName
                        u_matrixfolderpath      = $fd.MatrixFolderPath
                        u_matrixcategoryname    = $fd.MatrixCategoryName
                        u_matrixsubcategoryname = $fd.MatrixSubCategoryName
                        u_matrixresponsible     = $fd.MatrixResponsible
                        u_adobjectname          = $adObj.SamAccountName
                    }
                }
            }

            #
            # 4. Nothing to export?
            #
            if (-not $serviceNowRows) {
                return $false
            }

            #
            # 5. Export to Excel (SnowFormData sheet)
            #
            $xlsxParams = @{
                Path          = $OutputPath
                WorksheetName = 'SnowFormData'
                TableName     = 'SnowFormData'
                AutoSize      = $true
                FreezeTopRow  = $true
            }

            $serviceNowRows | Export-Excel @xlsxParams

            #
            # 6. Copy file to export log folder
            #
            $logCopyPath = Join-Path $ExportLogFolder 'ServiceNowFormData.xlsx'

            if (Test-Path -LiteralPath $OutputPath) {
                Copy-Item -LiteralPath $OutputPath -Destination $logCopyPath -Force
                return $true
            }

            return $false
        }
        catch {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Export-ServiceNowFormData failed: $_"
                }
            )
            return $false
        }
    }

    function Export-OverviewHtml {
        param(
            [Parameter(Mandatory)][object]$DataToExport,
            [Parameter(Mandatory)][string]$OutputPath,
            [Parameter(Mandatory)][string]$ExportLogFolder,
            [Parameter(Mandatory)][ref]$SystemErrors
        )

        try {
            #
            # 1. Remove old output file
            #
            Remove-FileHC -FilePath $OutputPath

            #
            # 2. Build table rows
            #
            function New-OverviewRow {
                param([object]$FormData)

                $category = $FormData.MatrixCategoryName
                $subcat = $FormData.MatrixSubCategoryName
                $folderPath = $FormData.MatrixFolderDisplayName
                $filePath = $FormData.MatrixFilePath
                $fileName = $FormData.MatrixFileName

                # Build mailto: list safely
                $emails = ($FormData.MatrixResponsible -split ',') |
                ForEach-Object {
                    $trimmed = $_.Trim()
                    if ($trimmed) { "mailto:$trimmed$trimmed</a>" }
                } |
                Join-String -Separator ', '

                return @"
<tr>
    <td>$category</td>
    <td>$subcat</td>
    <td>$folderPath$folderPath</a></td>
    <td>$filePath$fileName</a></td>
    <td>$emails</td>
</tr>
"@
            }

            $rows = $DataToExport.FormData |
            Sort-Object MatrixCategoryName, MatrixSubCategoryName, MatrixFolderDisplayName |
            ForEach-Object { New-OverviewRow -FormData $_ } |
            Join-String

            #
            # 3. Build full HTML document
            #
            $html = @"
<html>
<head>
<style>
    body { font-family:Arial; }
    table { width:100%; border-collapse:collapse; }
    th, td { padding:10px; border-bottom:1px solid #ddd; text-align:left; }
</style>
</head>
<body>
<h1>Matrix files overview</h1>

<table>
    <tr>
        <th>Category</th>
        <th>Subcategory</th>
        <th>Folder</th>
        <th>Link to matrix</th>
        <th>Responsible</th>
    </tr>
    $rows
</table>

</body>
</html>
"@

            #
            # 4. Save HTML output
            #
            $html | Out-File -LiteralPath $OutputPath -Encoding UTF8 -Force

            #
            # 5. Copy to log folder
            #
            $logCopy = Join-Path $ExportLogFolder 'Overview.html'
            Copy-Item -LiteralPath $OutputPath -Destination $logCopy -Force
        }
        catch {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Export-OverviewHtml failed: $_"
                }
            )
        }
    }

    function Export-Files {
        param(
            [Parameter(Mandatory)][object]$DataToExport,
            [Parameter(Mandatory)][object]$ExportConfig,
            [Parameter(Mandatory)][object]$ServiceNowConfig,
            [Parameter(Mandatory)][string]$ExportLogFolder,
            [Parameter(Mandatory)][object]$ScriptPathItem,
            [Parameter(Mandatory)][ref]$SystemErrors
        )

        #
        # Return structure
        #
        $results = @{}


        #
        # Helper: run a task safely and track failures
        #
        function Invoke-Safe {
            param(
                [scriptblock]$Action,
                [string]$ErrorMessage
            )
            try { & $Action }
            catch { 
                $SystemErrors.Value.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "$ErrorMessage`: $_"
                    }
                )
                return $false
            }
            return $true
        }


        #
        # 1. Export Permissions Excel
        #
        if ($ExportConfig.PermissionsExcelFile) {

            $ok = Invoke-Safe `
                -ErrorMessage 'Export-PermissionsFile failed' `
                -Action {
                Export-PermissionsFile `
                    -DataToExport $DataToExport `
                    -OutputPath $ExportConfig.PermissionsExcelFile `
                    -LogFolder $ExportLogFolder `
                    -SystemErrors $SystemErrors
            }

            if ($ok -and (Test-Path -LiteralPath $ExportConfig.PermissionsExcelFile -PathType Leaf)) {
                $results['PermissionsExcelFile'] = $ExportConfig.PermissionsExcelFile
            }
        }


        #
        # 2. Export + Upload ServiceNow form data
        #
        if ($ExportConfig.ServiceNowFormDataExcelFile -and $DataToExport.FormData) {

            $hasData = Invoke-Safe `
                -ErrorMessage 'Export-ServiceNowFormData failed' `
                -Action {
                Export-ServiceNowFormData `
                    -DataToExport $DataToExport `
                    -OutputPath $ExportConfig.ServiceNowFormDataExcelFile `
                    -ExportLogFolder $ExportLogFolder `
                    -SystemErrors $SystemErrors
            }

            if ($hasData -and (Test-Path -LiteralPath $ExportConfig.ServiceNowFormDataExcelFile -PathType Leaf)) {

                Invoke-Safe `
                    -ErrorMessage 'Upload-ServiceNowFormData failed' `
                    -Action {
                    Upload-ServiceNowFormData `
                        -OutputPath $ExportConfig.ServiceNowFormDataExcelFile `
                        -ServiceNowConfig $ServiceNowConfig `
                        -ScriptPathItem $ScriptPathItem `
                        -SystemErrors $SystemErrors
                }

                $results['ServiceNowFormDataExcelFile'] = $ExportConfig.ServiceNowFormDataExcelFile
            }
        }


        #
        # 3. Export Overview HTML
        #
        if ($ExportConfig.OverviewHtmlFile -and $DataToExport.FormData) {

            $ok = Invoke-Safe `
                -ErrorMessage 'Export-OverviewHtml failed' `
                -Action {
                Export-OverviewHtml `
                    -DataToExport $DataToExport `
                    -OutputPath $ExportConfig.OverviewHtmlFile `
                    -ExportLogFolder $ExportLogFolder `
                    -SystemErrors $SystemErrors
            }

            if ($ok -and (Test-Path -LiteralPath $ExportConfig.OverviewHtmlFile -PathType Leaf)) {
                $results['OverviewHtmlFile'] = $ExportConfig.OverviewHtmlFile
            }
        }


        return $results
    }

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

    function Build-ErrorWarningTable {
        param(
            [Parameter(Mandatory)][object]$CounterData,
            [Parameter(Mandatory)][object]$SystemErrors
        )

        #
        # Helper: build a table row in a single consistent way
        #
        function New-ErrorRow {
            param(
                [string]$CssClass,
                [string]$Label,
                [int]   $Count
            )
            return "<tr class='$CssClass'><th>$Label</th><td>$Count</td></tr>"
        }

        $rows = @()

        #
        # 1. System errors
        #
        if ($SystemErrors.Count -gt 0) {
            $rows += New-ErrorRow -CssClass 'probTextError' -Label 'System errors' -Count $SystemErrors.Count
        }

        #
        # 2. Matrix errors (excluding system errors)
        #
        $matrixErrors = $CounterData.TotalErrors - $SystemErrors.Count
        if ($matrixErrors -gt 0) {
            $rows += New-ErrorRow -CssClass 'probTextError' -Label 'Matrix errors' -Count $matrixErrors
        }

        #
        # 3. Matrix warnings
        #
        if ($CounterData.TotalWarnings -gt 0) {
            $rows += New-ErrorRow -CssClass 'probTextWarning' -Label 'Matrix warnings' -Count $CounterData.TotalWarnings
        }

        #
        # 4. If no rows, return empty string
        #
        if (-not $rows) {
            return ''
        }

        #
        # 5. Wrap rows in the final table
        #
        return "<p><b>Detected issues:</b></p><table class='errorWarningTable'>" +
        ($rows -join '') +
        '</table>'
    }

    function Write-EventLogSafe {
        param(
            [Parameter(Mandatory)][object]$EventLogData,
            [Parameter(Mandatory)][string]$ScriptName,
            [Parameter(Mandatory)][object]$Settings,
            [Parameter(Mandatory)][ref]$SystemErrors
        )

        # Maximum message length allowed by Windows Event Log
        $maxMessageLength = 31000

        try {
            # Is event log writing enabled?
            $logName = Get-StringValueHC $Settings.SaveInEventLog.LogName

            if (-not ($Settings.SaveInEventLog.Save -and $logName)) {
                return
            }

            #
            # 1. Convert system errors into event entries
            #
            foreach ($err in $SystemErrors.Value) {
                $EventLogData.Add(
                    [PSCustomObject]@{
                        Message   = $err.Message
                        DateTime  = $err.DateTime
                        EntryType = 'Error'
                        EventID   = '2'
                    }
                )
            }

            #
            # 2. Add "Script ended" marker
            #
            $EventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script ended'
                    DateTime  = Get-Date
                    EntryType = 'Information'
                    EventID   = '199'
                }
            )

            #
            # 3. Truncate messages exceeding allowed length
            #
            foreach ($item in $EventLogData) {
                if ($item.Message -and $item.Message.Length -gt $maxMessageLength) {
                    $item.Message = $item.Message.Substring(0, $maxMessageLength) +
                    '... [TRUNCATED DUE TO EVENT LOG SIZE LIMITS]'
                }
            }

            #
            # 4. Write the events
            #
            Write-EventsToEventLogHC `
                -Source $ScriptName `
                -LogName $logName `
                -Events $EventLogData
        }
        catch {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Event log write failed: $_"
                }
            )
        }
    }

    function Cleanup-OldLogs {
        param(
            [Parameter(Mandatory)]
            [string]$LogFolder,
            [int]$RetentionDays,
            [Parameter(Mandatory)]
            [ref]$SystemErrors
        )

        #
        # No cleanup if retention is disabled or folder missing
        #
        if ($RetentionDays -le 0 -or -not $LogFolder) {
            return
        }

        try {
            if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
                return
            }

            $cutoff = (Get-Date).AddDays(-$RetentionDays)

            #
            # 1. Delete old files
            #
            Get-ChildItem -LiteralPath $LogFolder -Recurse -File -ErrorAction Stop |
            Where-Object { $_.CreationTime -lt $cutoff } |
            ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                }
                catch {
                    $SystemErrors.Value.Add(
                        [PSCustomObject]@{
                            DateTime = Get-Date
                            Message  = "Log cleanup failed to delete file '$($_.FullName)': $_"
                        }
                    )
                }
            }

            #
            # 2. Remove empty folders (bottom‑up)
            #
            Get-ChildItem -LiteralPath $LogFolder -Recurse -Directory -ErrorAction Stop |
            Sort-Object FullName -Descending | # ensures children deleted before parents
            ForEach-Object {
                if (-not $_.GetFileSystemInfos().Count) {
                    try {
                        Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                    }
                    catch {
                        $SystemErrors.Value.Add(
                            [PSCustomObject]@{
                                DateTime = Get-Date
                                Message  = "Log cleanup failed to remove empty folder '$($_.FullName)': $_"
                            }
                        )
                    }
                }
            }
        }
        catch {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Log cleanup failed: $_"
                }
            )
        }
    }

    function Write-SystemErrorLog {
        param(
            [Parameter(Mandatory)][object]$SystemErrors,
            [Parameter(Mandatory)][string]$LogFolder,
            [Parameter(Mandatory)][ref]$MailParams
        )
        if ($SystemErrors.Count -gt 0 -and (Test-Path -LiteralPath $LogFolder -PathType Container)) {
            $path = Join-Path (Get-DatedLogFolderPathHC) 'System errors log'
            $attachments = Out-LogFileHC -DataToExport $SystemErrors -PartialPath $path -FileExtensions '.json' -ErrorAction Ignore
            if ($attachments) {
                if (-not $MailParams.Value.Contains('Attachments')) { $MailParams.Value['Attachments'] = @() }
                $MailParams.Value['Attachments'] += $attachments
            }
        }
    }

    function Generate-MailRecipientList {
        param(
            [object]$Recipients,
            [array]$Defaults = @()
        )

        #
        # Normalize a single input value into an array of strings
        #
        function Normalize-MailEntry {
            param([object]$Value)

            if (-not $Value) { return @() }
            if ($Value -is [string]) { return , $Value.Trim() }
            if ($Value -is [array]) { return $Value | ForEach-Object { "$_".Trim() } }

            # Unsupported type → ignore silently (same as original loose behavior)
            return @()
        }

        #
        # Normalize primary list + defaults
        #
        $combined =
        (Normalize-MailEntry -Value $Recipients) +
        (Normalize-MailEntry -Value $Defaults)

        #
        # Remove empty values, trim, unique, sorted
        #
        return $combined |
        ForEach-Object { "$_".Trim() } |
        Where-Object { $_ } |
        Sort-Object -Unique
    }

    function Generate-MailSubject {
        param(
            [int]$MatrixCount,
            [object]$SystemErrors,
            [object]$Counter,
            [string]$CustomSubject
        )

        #
        # Helper: pluralize a word if needed
        #
        function Plural {
            param(
                [int]$Count,
                [string]$Word
            )
            return ($Count -eq 1) ? $Word : ($Word + 's')
        }

        #
        # Optional custom suffix
        #
        $suffix = if ($CustomSubject) { ", $CustomSubject" } else { '' }

        $subject = $null

        #
        # 1. If system errors exist → priority subject line
        #
        if ($SystemErrors.Count -gt 0) {

            $sysWord = Plural -Count $SystemErrors.Count -Word 'System Error'
            $matWord = Plural -Count $MatrixCount -Word 'matrix file'

            if ($MatrixCount -gt 0) {
                $subject = "$MatrixCount $matWord, $($SystemErrors.Count) $sysWord$suffix"
            }
            else {
                # No matrix, only system failures
                $subject = "$sysWord`: $($SystemErrors.Count) critical failure$(if ($SystemErrors.Count -ne 1) {'s'})$suffix"
            }
        }

        #
        # 2. Otherwise → Matrix subject with error and warning summary
        #
        else {
            $matWord = Plural -Count $MatrixCount -Word 'matrix file'

            $errPart = if ($Counter.TotalErrors) {
                ", $($Counter.TotalErrors) error$(if($Counter.TotalErrors-ne 1){'s'})" 
            }
            else { '' }

            $warnPart = if ($Counter.TotalWarnings) {
                ", $($Counter.TotalWarnings) warning$(if($Counter.TotalWarnings-ne 1){'s'})" 
            }
            else { '' }

            $subject = "$MatrixCount $matWord$errPart$warnPart$suffix"
        }

        #
        # 3. Sanitize for use as filenames (original behavior preserved)
        #
        return [string]::Join(
            '_',
            $subject.Split([System.IO.Path]::GetInvalidFileNameChars())
        )
    }

    function Generate-MailBodyHtml {
        param(
            [Parameter(Mandatory)][object]$Settings,
            [Parameter(Mandatory)][object]$Html,
            [Parameter()][object]$ExportedFiles,
            [Parameter()][string]$AttNote,
            [Parameter()][string]$DurStr,
            [Parameter()][datetime]$ScriptStartTime,
            [Parameter()][string]$LogFolder
        )

        #
        # Helper: Create exported file links
        #
        function New-ExportListHtml {
            param([object]$Files)

            if (-not $Files -or $Files.Count -eq 0) {
                return ''
            }

            $items = $Files.GetEnumerator() |
            ForEach-Object {
                "<li>$($_.Value)$($_.Key)</a></li>"
            }

            return "<p><b>Exported $($Files.Count) file$(if($Files.Count-ne 1){'s'}):</b></p><ul>$($items -join '')</ul>"
        }

        #
        # Helper: Build the metadata table
        #
        function New-MetadataTable {
            param(
                [datetime]$Start,
                [string]$Duration,
                [string]$LogFolder
            )

            $startStr = $Start.ToString('dd/MM/yyyy HH:mm (dddd)')
            $logHtml = if ($LogFolder) {
                "<tr><th>Log files</th><td>$LogFolderOpen log folder</a></td></tr>"
            }

            return @"
<table class="aboutTable">
    <tr><th>Start time</th><td>$startStr</td></tr>
    <tr><th>Duration</th><td>$Duration</td></tr>
    $logHtml
    <tr><th>Host</th><td>$($host.Name)</td></tr>
    <tr><th>Computer</th><td>$env:COMPUTERNAME</td></tr>
    <tr><th>Account</th><td>$($env:USERDNSDOMAIN)\$($env:USERNAME)</td></tr>
</table>
"@
        }

        #
        # Compose sections
        #
        $exportHtml = New-ExportListHtml -Files $ExportedFiles
        $metaTable = New-MetadataTable -Start $ScriptStartTime -Duration $DurStr -LogFolder $LogFolder

        #
        # Main HTML document
        #
        return @"
<!DOCTYPE html>
<html>
<head>
    $($Html.Style)
</head>
<body>

<h1>$($Settings.ScriptName)</h1>
<hr size="2" color="#06cc7a">

$($Settings.SendMail.Body)
$($Html.ErrorWarningTable)
$exportHtml
$($Html.MatrixTables)
$AttNote

<hr size="2" color="#06cc7a">

$metaTable

</body>
</html>
"@
    }

    function Build-MailParameters {
        param(
            [Parameter(Mandatory)][object]$Settings,
            [Parameter(Mandatory)][object]$Html,
            [Parameter()][object]$ExportedFiles,
            [Parameter()][object]$Counter,
            [Parameter()][object]$SystemErrors,
            [Parameter()][int]$MatrixCount,
            [Parameter(Mandatory)][hashtable]$ExistingMailParams,
            [Parameter()][array]$MailToDefaultsFile,
            [Parameter()][string]$LogFolder,
            [Parameter()][datetime]$ScriptStartTime
        )

        #
        # 1. Prepare base hashtable
        #
        $mail = $ExistingMailParams
        $sendMail = $Settings.SendMail
        $smtp = $sendMail.Smtp

        #
        # 2. Recipients
        #
        $mail.To = Generate-MailRecipientList `
            -Recipients $sendMail.To `
            -Defaults $MailToDefaultsFile

        if ($sendMail.Bcc) {
            $mail.Bcc = Generate-MailRecipientList -Recipients $sendMail.Bcc
        }

        #
        # 3. Basic metadata
        #
        $mail.From = Get-StringValueHC $sendMail.From
        $mail.FromDisplayName = Get-StringValueHC $sendMail.FromDisplayName
        $mail.SmtpServerName = Get-StringValueHC $smtp.ServerName
        $mail.SmtpPort = Get-StringValueHC $smtp.Port
        $mail.SmtpConnectionType = Get-StringValueHC $smtp.ConnectionType

        $mail.MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
        $mail.MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit

        #
        # 4. Credential (optional)
        #
        if ($smtp.UserName -and $smtp.Password) {
            $sec = ConvertTo-SecureString `
                -String (Get-StringValueHC $smtp.Password) `
                -AsPlainText -Force

            $mail.Credential = New-Object `
                System.Management.Automation.PSCredential `
            (Get-StringValueHC $smtp.UserName), $sec
        }

        #
        # 5. Subject line
        #
        $mail.Subject = Generate-MailSubject `
            -MatrixCount $MatrixCount `
            -SystemErrors $SystemErrors `
            -Counter $Counter `
            -CustomSubject $sendMail.Subject

        #
        # 6. Mail priority
        #
        if (
            $SystemErrors.Count -gt 0 -or
            $Counter.TotalErrors -gt 0 -or
            $Counter.TotalWarnings -gt 0
        ) {
            $mail.Priority = 'High'
        }

        #
        # 7. Build the mail body
        #
        $attachmentNote = if ($mail.Attachments) {
            '<p><i>* Check the attachment(s) for details</i></p>'
        }

        $durationString = $null
        if ($ScriptStartTime) {
            $ts = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
            $durationString = '{0:00}:{1:00}:{2:00}' -f $ts.Hours, $ts.Minutes, $ts.Seconds
        }

        $mail.Body = Generate-MailBodyHtml `
            -Settings $Settings `
            -Html $Html `
            -ExportedFiles $ExportedFiles `
            -AttNote $attachmentNote `
            -DurStr $durationString `
            -ScriptStartTime $ScriptStartTime `
            -LogFolder $LogFolder

        return $mail
    }

    function Send-MailSafe {
        param(
            [Parameter(Mandatory)][hashtable]$MailParams,
            [Parameter(Mandatory)][ref]$SystemErrors
        )
        try { Send-MailKitMessageHC @MailParams }
        catch {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date 
                    Message  = "Failed to send mail: $_" 
                }
            ) 
        }
    }

    function Save-MailBodyToLog {
        param(
            [Parameter(Mandatory)][hashtable]$MailParams,
            [Parameter(Mandatory)][string]$LogFolder,
            [Parameter(Mandatory)][ref]$SystemErrors
        )

        try {
            # No subject → no log file
            if (-not $MailParams.Subject) {
                return
            }

            # Ensure log folder exists
            if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
                return
            }

            # Build final file path
            $fileName = "Mail - $($MailParams.Subject).html"
            $fullPath = Join-Path (Get-DatedLogFolderPathHC) $fileName

            # Save HTML
            $MailParams.Body |
            Out-File -LiteralPath $fullPath -Encoding UTF8 -Force
        }
        catch {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed to save mail HTML: $_"
                }
            )
        }
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
        
        $isExportNeeded =
        $Export.ServiceNowFormDataExcelFile -or
        $Export.PermissionsExcelFile -or
        $Export.OverviewHtmlFile
    
        $dataToExport = $null
        if ($importedMatrix -and $isExportNeeded) {
            $dataToExport = Build-ExportData `
                -ImportedMatrix $importedMatrix `
                -AdObjectHash $adObjectHash `
                -GroupManagerHash $groupManagerHash
        }

        #
        # 5. EXPORT FILES
        #
        $exportedFiles = @{}
        if ($systemErrors.Count -eq 0 -and $dataToExport -and $isExportNeeded) {

            $exportLogFolderPath = ''
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