#Requires -Version 7
#Requires -Modules Pester, ImportExcel

BeforeAll {
    $testInputFile = @{
        MaxConcurrent          = @{
            Computers             = 1
            JobsPerRemoteComputer = 1
            FoldersPerMatrix      = 2
        }
        Matrix                 = @{
            FolderPath             = (New-Item 'TestDrive:/Matrix' -ItemType Directory).FullName
            DefaultsFile           = (New-Item 'TestDrive:/Default.xlsx' -ItemType File).FullName
            Archive                = $false
            ExcludedSamAccountName = @()
        }
        Export                 = @{
            ServiceNowFormDataExcelFile = $null
            OverviewHtmlFile            = $null
            PermissionsExcelFile        = $null
        }
        ServiceNow             = @{
            CredentialsFilePath = (New-Item 'TestDrive:/cred.json' -ItemType File).FullName
            Environment         = 'Test'
            TableName           = 'roles'
        }
        PSSessionConfiguration = 'PowerShell.7'
        Settings               = @{
            ScriptName     = 'Test (Brecht)'
            SendMail       = @{
                When         = 'Always'
                From         = 'm@example.com'
                To           = '007@example.com'
                Subject      = 'Email subject'
                Body         = 'Email body'
                Smtp         = @{
                    ServerName     = 'SMTP_SERVER'
                    Port           = 25
                    ConnectionType = 'StartTls'
                    UserName       = 'bob'
                    Password       = 'pass'
                }
                AssemblyPath = @{
                    MailKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                    MimeKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
                }
            }
            SaveLogFiles   = @{
                Detailed            = $true
                Where               = @{
                    Folder = (New-Item 'TestDrive:/log' -ItemType Directory).FullName
                }
                deleteLogsAfterDays = 30
            }
            SaveInEventLog = @{
                Save    = $true
                LogName = 'Scripts'
            }
        }
    }

    $testLogFolder = $testInputFile.Settings.SaveLogFiles.Where.Folder

    $testOutParams = @{
        FilePath = (New-Item 'TestDrive:/Test.json' -ItemType File).FullName
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ConfigurationJsonFile = $testOutParams.FilePath
        ScriptPath            = @{
            TestRequirementsFile = (New-Item 'TestDrive:/TestRequirements.ps1' -ItemType File).FullName
            SetPermissionFile    = (New-Item 'TestDrive:/SetPermissions.ps1' -ItemType File).FullName
            UpdateServiceNow     = (New-Item 'TestDrive:/UpdateServiceNow.ps1' -ItemType File).FullName
        }
    }

    $testCherwellFolder = New-Item 'TestDrive:/Export' -ItemType Directory

    #region Valid Excel files
    $testMatrix = @(
        [PSCustomObject]@{
            Path   = 'Path'
            ACL    = @{'Bob' = 'L' }
            Parent = $true
            Ignore = $false
        }
    )
    $testPermissions = @(
        [PSCustomObject]@{P1 = $null      ; P2 = 'bob' }
        [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
        [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
        [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
        [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
    )
    $testSettings = @(
        [PSCustomObject]@{
            Status       = 'Enabled'
            ComputerName = 'S1'
            Path         = 'E:\Department'
            Action       = 'Check'
        }
    )
    $testDefaultSettings = @(
        [PSCustomObject]@{
            ADObjectName = 'Bob'
            Permission   = 'L'
            MailTo       = 'bob@contoso.com'
        }
        [PSCustomObject]@{
            ADObjectName = 'Mike'
            Permission   = 'R'
        }
    )
    #endregion

    $testDefaultSettings |
    Export-Excel -Path $testInputFile.Matrix.DefaultsFile -WorksheetName 'Settings'

    $testSettingsParams = @{
        Path          = Join-Path $testInputFile.Matrix.FolderPath 'Matrix.xlsx'
        WorkSheetName = 'Settings'
    }
    $testPermissionsParams = @{
        Path          = $testSettingsParams.Path
        WorkSheetName = 'Permissions'
        NoHeader      = $true
    }

    $testInputFile | ConvertTo-Json -Depth 7 | Out-File @testOutParams

    function Compare-HashTableHC {
        param (
            [Parameter(Mandatory)]
            [hashtable]$ReferenceObject,
            [Parameter(Mandatory)]
            [hashtable]$DifferenceObject
        )

        (
            $ReferenceObject.GetEnumerator() |
            Sort-Object { $_.Key } | ConvertTo-Json
        ) |
        Should -BeExactly (
            $DifferenceObject.GetEnumerator() |
            Sort-Object { $_.Key } | ConvertTo-Json
        )
    }
    function Test-GetLogFileDataHC {
        param (
            [String]$FileNameRegex = '* - System errors log.json',
            [String]$LogFolderPath = $testInputFile.Settings.SaveLogFiles.Where.Folder
        )

        $testLogFile = Get-ChildItem -Path $LogFolderPath -File -Filter $FileNameRegex

        if ($testLogFile.count -eq 1) {
            Get-Content $testLogFile | ConvertFrom-Json
        }
        elseif (-not $testLogFile) {
            throw "No log file found in folder '$LogFolderPath' matching '$FileNameRegex'"
        }
        else {
            throw "Found multiple log files in folder '$LogFolderPath' matching '$FileNameRegex'"
        }
    }
    function Test-NewJsonFileHC {
        try {
            if (-not $testNewInputFile) {
                throw "Variable '$testNewInputFile' cannot be blank"
            }

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams
        }
        catch {
            throw "Failure in Test-NewJsonFileHC: $_"
        }
    }
    function Copy-ObjectHC {
        <#
        .SYNOPSIS
            Make a deep copy of an object using JSON serialization.

        .DESCRIPTION
            Uses ConvertTo-Json and ConvertFrom-Json to create an independent
            copy of an object. This method is generally effective for objects
            that can be represented in JSON format.

        .PARAMETER InputObject
            The object to copy.

        .EXAMPLE
            $newArray = Copy-ObjectHC -InputObject $originalArray
        #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [Object]$InputObject
        )

        $jsonString = $InputObject | ConvertTo-Json -Depth 100

        $deepCopy = $jsonString | ConvertFrom-Json -AsHashtable

        return $deepCopy
    }
    function Send-MailKitMessageHC {
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
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
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
    }

    Mock Invoke-Command
    Mock New-PSSession
    Mock Send-MailKitMessageHC
    Mock Test-MatrixPermissionsHC
    Mock Test-MatrixSettingHC
    Mock Wait-MaxRunningJobsHC
    Mock Write-EventLog
    Mock Test-FormDataHC
    Mock Get-AdUserPrincipalNameHC
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @(
        'ConfigurationJsonFile'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'create an error log file when' {
    AfterAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
    
        Test-NewJsonFileHC
    }
    It 'the log folder cannot be created' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SaveLogFiles.Where.Folder = 'x:\notExistingLocation'

        Test-NewJsonFileHC

        Mock Out-File

        .$testScript @testParams

        $LASTEXITCODE | Should -Be 1

        Should -Not -Invoke Out-File
    }
    Context 'the ImportFile' {
        It 'is not found' {
            Mock Out-File

            $testNewParams = $testParams.clone()
            $testNewParams.ConfigurationJsonFile = 'nonExisting.json'

            .$testScript @testNewParams

            $LASTEXITCODE | Should -Be 1

            Should -Not -Invoke Out-File
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'MaxConcurrent', 'Matrix', 'Export', 'ServiceNow', 
                'PSSessionConfiguration'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                Test-NewJsonFileHC

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                $testLogFileContent = Test-GetLogFileDataHC

                $testLogFileContent[0].Message |
                Should -BeLike "*Property '$_' not found*"
            }
            It 'MaxConcurrent.<_> not found' -ForEach @(
                'Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.MaxConcurrent.$_ = $null

                Test-NewJsonFileHC

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                $testLogFileContent = Test-GetLogFileDataHC

                $testLogFileContent[0].Message |
                Should -BeLike "*Property 'MaxConcurrent.$_' not found*"
            }
            It 'Matrix.<_> not found' -ForEach @(
                'FolderPath', 'DefaultsFile'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Matrix.$_ = $null

                Test-NewJsonFileHC

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                $testLogFileContent = Test-GetLogFileDataHC

                $testLogFileContent[0].Message |
                Should -BeLike "*Property 'Matrix.$_' not found*"
            }
            It 'Matrix.<_> not existing' -ForEach @(
                'FolderPath', 'DefaultsFile'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Matrix.$_ = 'x:\NotExisting'

                Test-NewJsonFileHC

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                $testLogFileContent = Test-GetLogFileDataHC

                $testLogFileContent[0].Message |
                Should -BeLike "*Matrix.$_ 'x:\NotExisting' not found: *"
            }
        }
    }
    Context 'a script file' {
        It '<_> is not found' -ForEach @(
            'TestRequirementsFile', 'SetPermissionFile', 'UpdateServiceNow'
        ) {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.ScriptPath.$_ = 'x:\NotExisting.ps1'
    
            $testNewInputFile = Copy-ObjectHC $testInputFile
            Test-NewJsonFileHC
    
            .$testScript @testNewParams
    
            $LASTEXITCODE | Should -Be 1
    
            $testLogFileContent = Test-GetLogFileDataHC
    
            $testLogFileContent[0].Message |
            Should -BeLike "*ScriptPath.$_ 'x:\NotExisting.ps1' not found*"
        }
    }
    Context 'the default settings file' {
        It "is missing worksheet 'Settings'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Matrix.DefaultsFile = (New-Item 'TestDrive:/Folder/DefaultWrong.xlsx' -ItemType File -Force).FullName
    
            Test-NewJsonFileHC

            '1' | Export-Excel -Path $testNewInputFile.Matrix.DefaultsFile -WorksheetName 'Sheet1'
    
            .$testScript @testParams

            $LASTEXITCODE | Should -Be 1
    
            $testLogFileContent = Test-GetLogFileDataHC
    
            $testLogFileContent[0].Message |
            Should -BeLike "*'$($testNewInputFile.Matrix.DefaultsFile)'* worksheet 'Settings' not found*"
        }
    
        $TestCases = @(
            @{
                Name         = "column header 'MailTo'"
                DefaultsFile = @(
                    [PSCustomObject]@{
                        ADObjectName = 'Bob'
                        Permission   = 'L'
                    }
                    [PSCustomObject]@{
                        ADObjectName = 'Mike'
                        Permission   = 'R'
                    }
                )
                errorMessage = "Column header 'MailTo' not found"
            }
            @{
                Name         = "column header 'ADObjectName'"
                DefaultsFile = @(
                    [PSCustomObject]@{
                        Permission = 'L'
                        MailTo     = 'Bob@mail.com'
                    }
                    [PSCustomObject]@{
                        Permission = 'R'
                    }
                )
                errorMessage = "Column header 'ADObjectName' not found"
            }
            @{
                Name         = "column header 'Permission'"
                DefaultsFile = @(
                    [PSCustomObject]@{
                        ADObjectName = 'Bob'
                        MailTo       = 'Bob@mail.com'
                    }
                    [PSCustomObject]@{
                        ADObjectName = 'Mike'
                    }
                )
                errorMessage = "Column header 'Permission' not found"
            }
            @{
                Name         = "'MailTo' addresses"
                DefaultsFile = @(
                    [PSCustomObject]@{
                        ADObjectName = 'Bob'
                        Permission   = 'L'
                        MailTo       = $null
                    }
                    [PSCustomObject]@{
                        ADObjectName = 'Mike'
                        Permission   = 'R'
                        MailTo       = ' '
                    }
                )
                errorMessage = 'No mail addresses found'
            }
        )
    
        It 'is missing <Name>' -ForEach $TestCases {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Matrix.DefaultsFile = (New-Item 'TestDrive:/Folder/DefaultWrong.xlsx' -ItemType File -Force).FullName
    
            Test-NewJsonFileHC
    
            $DefaultsFile | Export-Excel -Path $testNewInputFile.Matrix.DefaultsFile -WorksheetName Settings
    
            .$testScript @testParams

            $LASTEXITCODE | Should -Be 1
    
            $testLogFileContent = Test-GetLogFileDataHC
    
            $testLogFileContent[0].Message |
            Should -BeLike "*$($testNewInputFile.DefaultsFile)*$errorMessage*"
        }
    }
}
Describe 'a sub folder in the log folder' {
    BeforeAll {
        @(
            [PSCustomObject]@{
                Status       = $null
                ComputerName = 'S1'
                Path         = 'E:\Test'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @testPermissionsParams

        .$testScript @testParams
    }
    It "is created for each specific Excel file regardless of its 'Status'" {
        @(Get-ChildItem -Path $testLogFolder -Directory).Count |
        Should -BeExactly 1
    }
    It 'the Excel file is copied to the log folder' {
        $testMatrixLogFolder = Get-ChildItem -Path $testLogFolder -Directory

        @(Get-ChildItem -Path $testMatrixLogFolder.FullName -File -Filter '*.xlsx').Count | Should -BeExactly 1
    }
}
Describe "when 'Matrix.Archive' is true then" {
    BeforeAll {
        @(
            [PSCustomObject]@{
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @testPermissionsParams

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Matrix.Archive = $true

        Test-NewJsonFileHC

        .$testScript @testParams
    }
    It "a sub folder in the 'Matrix.FolderPath' named 'Archive' is created" {
        "$($testNewInputFile.Matrix.FolderPath)\Archive" | Should -Exist
    }
    It 'all matrix files are moved to the archive folder, even disabled ones' {
        $testSettingsParams.Path | Should -Not -Exist
        "$($testNewInputFile.Matrix.FolderPath)\Archive\Matrix.xlsx" | Should -Exist
    }
    It 'a matrix with the same name is overwritten in the archive folder' {
        $testFile = "$($testNewInputFile.Matrix.FolderPath)\Archive\Matrix.xlsx"
        $testFile | Remove-Item -Force -EA Ignore

        @(
            [PSCustomObject]@{
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel -Path $testFile -WorksheetName $testSettingsParams.WorkSheetName

        $testFile | Should -Exist

        @(
            [PSCustomObject]@{
                ComputerName = 'S2'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @testPermissionsParams

        .$testScript @testParams

        $testFile | Should -Exist
        $testSettingsParams.Path | Should -Not -Exist
        (Import-Excel -Path $testFile -WorksheetName Settings).ComputerName |
        Should -Be 'S2'
    }
    It 'multiple matrix files are moved to the archive folder' {
        Remove-Item -Path "$($testNewInputFile.Matrix.FolderPath)\Archive" -Recurse -EA Ignore
        1..5 | ForEach-Object {
            $FileName = "$($testNewInputFile.Matrix.FolderPath)\Matrix $_.xlsx"
            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'S1'
                    Path         = 'E:\Department'
                    GroupName    = 'G1'
                    SiteName     = 'S1'
                    SiteCode     = 'C1'
                    Action       = 'Check'
                }
            ) | Export-Excel -Path $FileName -WorksheetName Settings
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel -Path $FileName -WorksheetName Permissions -NoHeader
        }

        .$testScript @testParams

        (Get-ChildItem "$($testNewInputFile.Matrix.FolderPath)\Matrix*" -File).Count | Should -BeExactly 0
        (Get-ChildItem "$($testNewInputFile.Matrix.FolderPath)\Archive" -File).Count |
        Should -BeExactly 5
    }
}
Describe 'do not invoke the script to set permissions when' {
    It "there's only a default settings file in the 'Matrix.FolderPath' folder" {
        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
    It "there are only other file types than '.xlsx' in the 'Matrix.FolderPath' folder" {
        1 | Out-File "$($testInputFile.Matrix.FolderPath)\Wrong.txt"
        1 | Out-File "$($testInputFile.Matrix.FolderPath)\Wrong.csv"

        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
    It "there are only valid matrixes in subfolders of the 'Matrix.FolderPath' folder" {
        $Folder = (New-Item "$($testInputFile.Matrix.FolderPath)\Archive" -ItemType Directory -Force -EA Ignore).FullName
        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel -Path "$Folder/Matrix.xlsx" -WorksheetName Settings
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel -Path "$Folder/Matrix.xlsx" -WorksheetName Permissions -NoHeader

        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
    It "the 'Status' in the worksheet 'Settings' of the matrix file is not set to 'Enabled'" {
        @(
            [PSCustomObject]@{
                Status       = 'NOTEnabled'
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @testPermissionsParams

        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
}
Describe 'a FatalError object is registered' {
    AfterEach {
        $Error.Clear()
        Remove-Item -Path "$($testLogFolder)\*" -Recurse -Force -EA Ignore
        Remove-Item -Path "$($testInputFile.Matrix.FolderPath)\*" -Exclude 'Default.xlsx' -Recurse -Force -EA Ignore
    }
    Context "for the Excel 'File' when" {
        It "building the matrix with 'ConvertTo-MatrixAclHC' fails" {
            Mock ConvertTo-MatrixAclHC {
                throw 'Failed building the matrix'
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'S1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @testPermissionsParams

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Unknown error'
                Description = 'While checking the input and generating the matrix an error was reported.'
                Value       = 'Failed building the matrix'
            }.GetEnumerator().ForEach( 
                { $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value }
            )
        }
        It 'the worksheet Settings is not found' {
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @testPermissionsParams

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = "Failed importing the Excel workbook '$($testPermissionsParams.Path)' with worksheet 'Settings'*"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) |
                    Should -BeLike $_.Value
                })
        } -Skip
        It 'the worksheet Settings is empty' {
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @testPermissionsParams

            #region Add empty worksheet
            $pkg = New-Object OfficeOpenXml.ExcelPackage (Get-Item -Path $testSettingsParams.Path)
            $null = $pkg.Workbook.Worksheets.Add('Settings')
            $pkg.Save()
            $pkg.Dispose()
            #endregion

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = "Failed importing the Excel workbook '$($testSettingsParams.Path)' with worksheet 'Settings'*"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) |
                    Should -BeLike $_.Value
                })
        } -Skip
        It "the worksheet Permissions is not found when the 'Settings' sheet has 'Status' set to 'Enabled'" {
            $testSettings | Export-Excel @testSettingsParams

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = "Worksheet 'Permissions' not found"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value
                })
        } -Skip
        It "the worksheet Permissions is empty when the 'Settings' sheet has 'Status' set to 'Enabled'" {
            $testSettings | Export-Excel @testSettingsParams

            #region Add empty worksheet
            $pkg = New-Object OfficeOpenXml.ExcelPackage (Get-Item -Path $testSettingsParams.Path)
            $null = $pkg.Workbook.Worksheets.Add('Permissions')
            $pkg.Save()
            $pkg.Dispose()
            #endregion

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = "Worksheet 'Permissions' is empty"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value
                })
        }
    }
    Context "for the worksheet 'Permissions' when" {
        AfterAll {
            Mock Test-MatrixPermissionsHC
        }
        It "'Test-MatrixPermissionsHC' detects an input problem" {
            Mock Test-MatrixPermissionsHC {
                @{
                    Type = 'Warning'
                    Name = 'Matrix permission incorrect'
                }
            }

            $testSettings | Export-Excel @testSettingsParams
            $testPermissions | Export-Excel @testPermissionsParams

            .$testScript @testParams

            $ImportedMatrix.Permissions.Check.Name |
            Should -Contain 'Matrix permission incorrect'
        }
    }
    Context "for the worksheet 'Settings' when" {
        AfterAll {
            Mock Test-MatrixSettingHC
            Mock Test-ExpandedMatrixHC
        }
        It 'a duplicate ComputerName/Path combination is found' {
            $testProblem = @(
                [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Duplicate ComputerName/Path combination'
                    Description = "Every 'ComputerName' combined with a 'Path' needs to be unique over all the 'Settings' worksheets found in all the active matrix files."
                    Value       = @{
                        ('S1.' + $env:USERDNSDOMAIN) = 'E:\DUPLICATE'
                    }
                }
            )

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = $($testProblem.Value.Keys)
                    Path         = 'E:\Reports'
                    Action       = 'Check'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = $($testProblem.Value.Keys)
                    Path         = $($testProblem.Value.Values)
                    Action       = 'Check'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = $($testProblem.Value.Keys)
                    Path         = $($testProblem.Value.Values)
                    Action       = 'Fix'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'S3'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams

            $testPermissions | Export-Excel @testPermissionsParams

            .$testScript @testParams

            $toTest = @($ImportedMatrix.Settings.Where(
                    { $_.Import.Path -eq $testProblem.Value.Values }
                ))

            $toTest.Count | Should -BeExactly 2

            foreach ($testMatrix in $toTest) {
                $testCheck = $testMatrix.Check | Where-Object {
                    $_.Name -eq $testProblem.Name
                }
                $testCheck.Type | Should -Be $testProblem.Type
                $testCheck.Name | Should -Be $testProblem.Name
                $testCheck.Description | Should -Be $testProblem.Description
                $testCheck.Value.Name | Should -Be $testProblem.Value.Name
                $testCheck.Value.Value | Should -Be $testProblem.Value.Value
            }

        }
        It "'Test-MatrixSettingHC' detects an input problem" {
            $testProblem = @{
                Name = 'Matrix setting incorrect'
            }
            Mock Test-MatrixSettingHC {
                $testProblem
            }

            $testSettings | Export-Excel @testSettingsParams
            $testPermissions | Export-Excel @testPermissionsParams

            .$testScript @testParams

            $testProblem.Name |
            Should -Be ($ImportedMatrix.Settings.Check | Where-Object Name -EQ $testProblem.Name).Name
        }
        It "'Test-ExpandedMatrixHC' detects a problem" {
            Mock Test-MatrixSettingHC
            $testProblem = @{
                Name = 'Expansion incorrect'
            }
            Mock Test-ExpandedMatrixHC {
                $testProblem
            }
            Mock ConvertTo-MatrixAclHC {
                $testMatrix
            }

            Mock Get-ADObjectDetailHC { $true }

            $testSettings | Export-Excel @testSettingsParams
            $testPermissions | Export-Excel @testPermissionsParams

            .$testScript @testParams

            $testProblem.Name |
            Should -Be ($ImportedMatrix.Settings.Check | Where-Object Name -EQ $testProblem.Name).Name
        }
    }
}
Describe 'a Warning object is registered' {
    AfterEach {
        $Error.Clear()
        Remove-Item -Path "$($testLogFolder)\*" -Recurse -Force -EA Ignore
        Remove-Item -Path "$($testInputFile.Matrix.FolderPath)\*" -Exclude 'Default.xlsx' -Recurse -Force -EA Ignore
    }
    Context "for the Excel 'File' when" {
        It "the worksheet 'Settings' has no row with status 'Enabled'" {
            @(
                [PSCustomObject]@{
                    Status       = $null
                    ComputerName = 'A'
                    Path         = 'E:\Reports'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams
            $testPermissions | Export-Excel @testPermissionsParams

            .$testScript @testParams

            @{
                Type        = 'Warning'
                Name        = 'Matrix disabled'
                Description = 'Every Excel file needs at least one enabled matrix.'
                Value       = "The worksheet 'Settings' does not contain a row with 'Status' set to 'Enabled'."
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value
                })
        }
    }
}
Describe "each row in the worksheet 'settings'" {
    BeforeAll {
        Mock ConvertTo-MatrixADNamesHC { @{} }
        Mock ConvertTo-MatrixAclHC
        Mock Test-AdObjectsHC
        Mock Test-MatrixSettingHC

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'pc1'
                Path         = 'E:\Department'
                Action       = 'Check'
                GroupName    = 'A'
                SiteCode     = 'B'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'pc2'
                Path         = 'E:\Reports'
                Action       = 'Check'
                GroupName    = 'C'
                SiteCode     = 'D'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'pc3'
                Path         = 'E:\Finance'
                Action       = 'Check'
                GroupName    = 'E'
                SiteCode     = 'F'
            }
        ) | Export-Excel @testSettingsParams

        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'bob' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @testPermissionsParams

        .$testScript @testParams -EA ignore
    }
    It 'is tested for incorrect input' {
        Should -Invoke Test-MatrixSettingHC -Exactly 3 -Scope Describe
        @('pc1', 'pc2', 'pc3') | ForEach-Object {
            Should -Invoke Test-MatrixSettingHC -Exactly 1 -Scope Describe -ParameterFilter {
                $Setting.ComputerName -eq $_
            }
        }
    }
    Context 'creates a unique matrix with' {
        It 'complete SamAccountNames constructed from the header rows' {
            function testColumnHeaders {
                ($null -eq $ColumnHeaders[0].P1) -and
                ($ColumnHeaders[0].P2 -eq 'bob') -and
                ($ColumnHeaders[1].P1 -eq 'SiteCode') -and
                ($ColumnHeaders[1].P2 -eq 'SiteCode') -and
                ($ColumnHeaders[2].P1 -eq 'GroupName') -and
                ($ColumnHeaders[2].P2 -eq 'GroupName')
            }

            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 3 -Scope Describe
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Begin -eq 'A') -and ($Middle -eq 'B') -and
                (testColumnHeaders)
            }
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Begin -eq 'C') -and ($Middle -eq 'D') -and
                (testColumnHeaders)
            }
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Begin -eq 'E') -and ($Middle -eq 'F') -and
                (testColumnHeaders)
            }
        }
        It 'path and Acl' {
            Should -Invoke ConvertTo-MatrixAclHC -Exactly 3 -Scope Describe
        }
    }
}
Describe "the worksheet 'Permissions' is" {
    BeforeAll {
        Mock Test-MatrixPermissionsHC

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Department'
                Action       = 'Check'
                GroupName    = 'A'
                SiteCode     = 'B'
            }
        ) | Export-Excel @testSettingsParams

        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'bob' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @testPermissionsParams

        .$testScript @testParams
    }
    It 'tested for incorrect input' {
        Should -Invoke Test-MatrixPermissionsHC -Exactly 1 -Scope Describe
        Should -Invoke Test-MatrixPermissionsHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($null -eq $Permissions[0].P1) -and
            ($Permissions[0].P2 -eq 'bob') -and
            ($Permissions[1].P1 -eq 'SiteCode') -and
            ($Permissions[1].P2 -eq 'SiteCode') -and
            ($Permissions[2].P1 -eq 'GroupName') -and
            ($Permissions[2].P2 -eq 'GroupName') -and
            ($Permissions[3].P1 -eq 'Path') -and
            ($Permissions[3].P2 -eq 'L') -and
            ($Permissions[4].P1 -eq 'Folder') -and
            ($Permissions[4].P2 -eq 'W')
        }
    }
}
Describe 'the script that tests the remote computers for compliance' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        Mock Invoke-Command {
            'A'
        } -ParameterFilter {
            ($ComputerName -eq 'PC1') -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptPath.TestRequirementsFile)
        }
        Mock Invoke-Command {
            'B'
        } -ParameterFilter {
            ($ComputerName -eq 'PC2') -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptPath.TestRequirementsFile)
        }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Reports'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC2'
                Path         = 'E:\Finance'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = $null
                ComputerName = 'ignoredPc'
                Path         = 'E:\Finance'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams

        $testPermissions | Export-Excel @testPermissionsParams

        .$testScript @testParams
    }
    It "is not called for rows in the 'Settings' worksheets where Status is not Enabled" {
        Should -Not -Invoke Invoke-Command -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptPath.TestRequirementsFile) -and
            ($ComputerName -eq 'ignoredPc')
        }
    }
    It "is only called for unique ComputerNames in the 'Settings' worksheets" {
        Should -Invoke Invoke-Command -Times 2 -Exactly -Scope Describe -ParameterFilter {
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptPath.TestRequirementsFile)
        }
        @('PC1', 'PC2') | ForEach-Object {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ConfigurationName) -and
                ($FilePath -eq $testParams.ScriptPath.TestRequirementsFile) -and
                ($ComputerName -eq $_)
            }
        }
    }
    It 'saves the job result in Settings for each matrix' {
        @($ImportedMatrix.Settings.Where(
                {
                    ($_.Import.ComputerName -eq 'PC1') -and
                    ($_.Check -eq 'A') }
            )
        ).Count |
        Should -BeExactly 2

        @($ImportedMatrix.Settings.Where(
                {
                    ($_.Import.ComputerName -eq 'PC2') -and
                    ($_.Check -eq 'B') }
            )
        ).Count |
        Should -BeExactly 1
    }
}
Describe 'the script that sets the permissions on the remote computers' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        Mock Invoke-Command { 1 } -ParameterFilter {
            ($ComputerName -eq 'PC1') -and
            ($ArgumentList[0] -eq 'E:\Department') -and
            ($ArgumentList[1] -eq 'New') -and
            ($ArgumentList[2]) -and
            ($ArgumentList[3] -eq $testInputFile.MaxConcurrent.FoldersPerMatrix) -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptPath.SetPermissionFile)
        }
        Mock Invoke-Command { 2 } -ParameterFilter {
            ($ComputerName -eq 'PC1') -and
            ($ArgumentList[0] -eq 'E:\Reports') -and
            ($ArgumentList[1] -eq 'Fix') -and
            ($ArgumentList[2]) -and
            ($ArgumentList[3] -eq $testInputFile.MaxConcurrent.FoldersPerMatrix) -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptPath.SetPermissionFile)
        }
        Mock Invoke-Command { 3 } -ParameterFilter {
            ($ComputerName -eq 'PC2') -and
            ($ArgumentList[0] -eq 'E:\Finance') -and
            ($ArgumentList[1] -eq 'Check') -and
            ($ArgumentList[2]) -and
            ($ArgumentList[3] -eq $testInputFile.MaxConcurrent.FoldersPerMatrix) -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptPath.SetPermissionFile)
        }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Department'
                Action       = 'New'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Reports'
                Action       = 'Fix'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC2'
                Path         = 'E:\Finance'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = $null
                ComputerName = 'ignoredPc'
                Path         = 'E:\Finance'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams

        $testPermissions | Export-Excel @testPermissionsParams

        .$testScript @testParams
    }
    It "is not called for rows in the 'Settings' worksheets where Status is not Enabled" {
        Should -Not -Invoke Invoke-Command -Scope Describe -ParameterFilter {
            ($ComputerName -eq 'ignoredPc')
        }
    }
    It "is called for each row in the 'Settings' worksheets with Status Enabled" {
        Should -Invoke Invoke-Command -Times 3 -Exactly -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptPath.SetPermissionFile)
        }
        Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptPath.SetPermissionFile) -and
            ($ComputerName -eq 'PC1') -and
            ($ArgumentList[0] -eq 'E:\Department') -and
            ($ArgumentList[1] -eq 'New') -and
            ($ArgumentList[2] -ne $null) -and
            ($ArgumentList[3] -ne $null)
        }
        Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptPath.SetPermissionFile) -and
            ($ComputerName -eq 'PC1') -and
            ($ArgumentList[0] -eq 'E:\Reports') -and
            ($ArgumentList[1] -eq 'Fix') -and
            ($ArgumentList[2] -ne $null) -and
            ($ArgumentList[3] -ne $null)
        }
        Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptPath.SetPermissionFile) -and
            ($ComputerName -eq 'PC2') -and
            ($ArgumentList[0] -eq 'E:\Finance') -and
            ($ArgumentList[1] -eq 'Check') -and
            ($ArgumentList[2] -ne $null) -and
            ($ArgumentList[3] -ne $null)
        }
    }
    It 'saves the start/end/duration times for each job in the settings' {
        $ImportedMatrix.Settings.JobTime.Start | Should -HaveCount 3
        $ImportedMatrix.Settings.JobTime.End | Should -HaveCount 3
        $ImportedMatrix.Settings.JobTime.Duration | Should -HaveCount 3
    }
    It 'saves the job result in Settings for each matrix' {
        ($ImportedMatrix.Settings.Where( { ($_.ID -eq 1) })).Check |
        Should -Contain 1
        ($ImportedMatrix.Settings.Where( { ($_.ID -eq 2) })).Check |
        Should -Contain 2
        ($ImportedMatrix.Settings.Where( { ($_.ID -eq 3) })).Check |
        Should -Contain 3
    }
}
Describe 'an email is sent to the user in the default settings file' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Reports'
                Action       = 'Check'
                GroupName    = 'C'
                SiteCode     = 'D'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC2'
                Path         = 'E:\Finance'
                Action       = 'New'
                GroupName    = 'x'
                SiteCode     = 'x'
            }
        ) | Export-Excel @testSettingsParams

        $testPermissions | Export-Excel @testPermissionsParams

        .$testScript @testParams
    }
    It 'containing a summary per Settings row for executed matrixes' {
        Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($From -eq 'm@example.com') -and
            ($To[0] -eq '007@example.com') -and
            ($To[1] -eq 'bob@contoso.com') -and
            ($SmtpPort -eq 25) -and
            ($SmtpServerName -eq 'SMTP_SERVER') -and
            ($SmtpConnectionType -eq 'StartTls') -and
            ($Subject -eq '1 matrix file, Email subject') -and
            ($Body -like (
                '*<style type="text/css">*</style>
            <body>
            <h1>Test (Brecht)</h1>
            </body>*'
            ) -replace '[\r\n]+', '*'
            )
        }
    }
}
Describe 'export an Excel file with' {
    BeforeAll {
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                samAccountName = 'A B bob'
                adObject       = @{
                    ObjectClass    = 'user'
                    Name           = 'A B Bob'
                    SamAccountName = 'A B bob'
                    ManagedBy      = $null
                }
                adGroupMember  = $null
            }
            [PSCustomObject]@{
                samAccountName = 'movieStars'
                adObject       = @{
                    ObjectClass    = 'group'
                    Name           = 'Movie Stars'
                    SamAccountName = 'movieStars'
                    ManagedBy      = $null
                }
                adGroupMember  = $null
            }
            [PSCustomObject]@{
                samAccountName = 'starTrekCaptains'
                adObject       = @{
                    ObjectClass    = 'group'
                    SamAccountName = 'starTrekCaptains'
                    Name           = 'Star Trek Captains'
                    ManagedBy      = 'CN=CaptainManagers,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Jean Luc Picard'
                        SamAccountName = 'picard'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Ignored account'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
            [PSCustomObject]@{
                samAccountName = 'singers'
                adObject       = @{
                    ObjectClass    = 'group'
                    SamAccountName = 'singers'
                    Name           = 'Singers'
                    ManagedBy      = 'CN=SingerManagers,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Beyonce'
                        SamAccountName = 'queenb'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Ignored account'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'SamAccountName' }
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                DistinguishedName = 'CN=CaptainManagers,DC=contoso,DC=net'
                adObject          = @{
                    ObjectClass = 'group'
                    Name        = 'Captain Managers'
                }
                adGroupMember     = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Admiral Pike'
                        SamAccountName = 'pike'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Excluded user'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
            [PSCustomObject]@{
                DistinguishedName = 'CN=SingerManagers,DC=contoso,DC=net'
                adObject          = @{
                    ObjectClass = 'group'
                    Name        = 'Singer Managers'
                }
                adGroupMember     = $null
            }
        } -ParameterFilter { $Type -eq 'DistinguishedName' }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Reports'
                Action       = 'Check'
                GroupName    = 'A'
                SiteCode     = 'B'
            }
        ) | Export-Excel @testSettingsParams

        @(
            [PSCustomObject]@{
                P1 = $null      ; P2 = 'bob'       ; P3 = 'movieStars'; P4 = '' ; P5 = ''
            }
            [PSCustomObject]@{
                P1 = 'SiteCode' ; P2 = 'SiteCode'  ; P3 = ''; P4 = 'starTrekCaptains' ; P5 = ''
            }
            [PSCustomObject]@{
                P1 = 'GroupName'; P2 = 'GroupName' ; P3 = ''; P4 = '' ; P5 = 'Singers'
            }
            [PSCustomObject]@{
                P1 = 'Path'     ; P2 = 'L'         ; P3 = ''; P4 = '' ; P5 = ''
            }
            [PSCustomObject]@{
                P1 = 'Folder'   ; P2 = 'W'         ; P3 = ''; P4 = '' ; P5 = ''
            }
        ) | Export-Excel @testPermissionsParams

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Matrix.ExcludedSamAccountName = @('IgnoreMe')

        Test-NewJsonFileHC

        .$testScript @testParams

        $testMatrixFile = Get-ChildItem $testLogFolder -Filter '*Matrix.xlsx' -Recurse -File
    }
    Context "the worksheet 'AccessList'" {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    SamAccountName       = 'starTrekCaptains'
                    Name                 = 'Star Trek Captains'
                    Type                 = 'group'
                    MemberName           = 'Jean Luc Picard'
                    MemberSamAccountName = 'picard'
                }
                @{
                    SamAccountName       = 'A B bob'
                    Name                 = 'A B Bob'
                    Type                 = 'user'
                    MemberName           = $null
                    MemberSamAccountName = $null
                }
                @{
                    SamAccountName       = 'Singers'
                    Name                 = 'Singers'
                    Type                 = 'group'
                    MemberName           = 'Beyonce'
                    MemberSamAccountName = 'queenb'
                }
                @{
                    SamAccountName       = 'movieStars'
                    Name                 = 'Movie Stars'
                    Type                 = 'group'
                    MemberName           = $null
                    MemberSamAccountName = $null
                }
            )

            $actual = Import-Excel -Path $testMatrixFile.FullName -WorksheetName 'AccessList'
        }
        It 'added to the matrix log file' {
            $actual | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.SamAccountName -eq $testRow.SamAccountName
                }
                $actualRow.Name | Should -Be $testRow.Name
                $actualRow.Type | Should -BeLike $testRow.Type
                $actualRow.MemberName | Should -Be $testRow.MemberName
                $actualRow.MemberSamAccountName | Should -BeLike $testRow.MemberSamAccountName
            }
        }
    }
    Context "the worksheet 'GroupManagers'" {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    GroupName         = 'Star Trek Captains'
                    ManagerName       = 'Captain Managers'
                    ManagerType       = 'group'
                    ManagerMemberName = 'Admiral Pike'
                }
                @{
                    GroupName         = 'Singers'
                    ManagerName       = 'Singer Managers'
                    ManagerType       = 'group'
                    ManagerMemberName = $null
                }
                @{
                    GroupName         = 'Movie Stars'
                    ManagerName       = $null
                    ManagerType       = $null
                    ManagerMemberName = $null
                }
            )

            $actual = Import-Excel -Path $testMatrixFile.FullName -WorksheetName 'GroupManagers'
        }
        It 'added to the matrix log file' {
            $actual | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.GroupName -eq $testRow.GroupName
                }
                $actualRow.ManagerName | Should -Be $testRow.ManagerName
                $actualRow.ManagerType | Should -BeLike $testRow.ManagerType
                $actualRow.ManagerMemberName | Should -Be $testRow.ManagerMemberName
            }
        }
    }
}
Describe 'when a job fails' {
    Context 'the test requirements script' {
        BeforeAll {
            Mock Test-ExpandedMatrixHC
            Mock Invoke-Command { throw 'failure' } -ParameterFilter {
                ($ComputerName -eq 'PC1') -and
                ($ConfigurationName) -and
                ($FilePath -eq $testParams.ScriptPath.TestRequirementsFile)
            }
            Mock Invoke-Command { 'B' } -ParameterFilter {
                ($ComputerName -eq 'PC2') -and
                ($ConfigurationName) -and
                ($FilePath -eq $testParams.ScriptPath.TestRequirementsFile)
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC2'
                    Path         = 'E:\Reports'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams

            $testPermissions | Export-Excel @testPermissionsParams

            .$testScript @testParams
        }
        It 'the job error is saved in Settings for each matrix' {
            $actual = $ImportedMatrix.Settings.Where( { ($_.ID -eq 1) })
            $actual.Check.Type | Should -Be 'FatalError'
            $actual.Check.Value | Should -Be 'failure'

            $actual = $ImportedMatrix.Settings.Where( { ($_.ID -eq 2) })
            $actual.Check.Type | Should -Not -Be 'FatalError'
            $actual.Check.Value | Should -Not -Be 'failure'
        }
    }
    Context 'the set permissions script' {
        BeforeAll {
            Mock Test-ExpandedMatrixHC
            Mock Invoke-Command { 1 } -ParameterFilter {
                ($ConfigurationName) -and
                ($ArgumentList[0] -eq 'E:\Department') -and
                ($FilePath -eq $testParams.ScriptPath.SetPermissionFile)
            }
            Mock Invoke-Command { throw 'failure' } -ParameterFilter {
                ($ConfigurationName) -and
                ($ArgumentList[0] -eq 'E:\Reports') -and
                ($FilePath -eq $testParams.ScriptPath.SetPermissionFile)
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Reports'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams

            $testPermissions | Export-Excel @testPermissionsParams

            .$testScript @testParams
        }
        It 'the job error is saved in Settings for each matrix' {
            $actual = $ImportedMatrix.Settings.Where( { ($_.ID -eq 1) })
            $actual.Check.Type | Should -Not -Be 'FatalError'
            $actual.Check.Value | Should -Not -Be 'failure'

            $actual = $ImportedMatrix.Settings.Where( { ($_.ID -eq 2) })
            $actual.Check.Type | Should -Be 'FatalError'
            $actual.Check.Value | Should -Be 'failure'
        }
    }
}
Describe 'internal functions' {
    Context 'default permissions vs matrix permissions' {
        It 'add default permissions to the matrix' {
            Mock Test-ExpandedMatrixHC
            Mock ConvertTo-MatrixAclHC {
                @(
                    [PSCustomObject]@{
                        Path   = 'Path'
                        ACL    = @{'Mike' = 'L' }
                        Parent = $true
                        Ignore = $false
                    }
                )
            }
            Mock Get-DefaultAclHC {
                @{
                    'Bob' = 'R'
                }
            }
            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'test'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Mike' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = '' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = '' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @testPermissionsParams

            .$testScript @testParams

            $actual = ($ImportedMatrix.Settings.Matrix.Where(
                    { $_.Path -eq 'Path' })
            ).ACL

            $expected = @{
                'Bob'  = 'R'
                'Mike' = 'L'
            }

            Compare-HashTableHC $actual $expected
        }
        It 'do not add default permissions to the matrix ACL when the folder has no ACL' {
            Mock Test-ExpandedMatrixHC
            Mock ConvertTo-MatrixAclHC {
                @(
                    [PSCustomObject]@{
                        Path   = 'Path'
                        ACL    = @{}
                        Parent = $true
                        Ignore = $false
                    }
                )
            }
            Mock Get-DefaultAclHC {
                @{
                    'Bob' = 'R'
                }
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'test'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Mike' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = '' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = '' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = '' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'L' }
            ) | Export-Excel @testPermissionsParams

            .$testScript @testParams

            $actual = ($ImportedMatrix.Settings.Matrix.Where( {
                        $_.Path -eq 'Path' })).ACL

            $actual | Should -BeNullOrEmpty
        }
        It 'do not overwrite permissions to the matrix ACL when they are also in the default ACL' {
            Mock Test-ExpandedMatrixHC
            Mock ConvertTo-MatrixAclHC {
                @(
                    [PSCustomObject]@{
                        Path   = 'Path'
                        ACL    = @{
                            'Mike' = 'L'
                            'Bob'  = 'L'
                        }
                        Parent = $true
                        Ignore = $false
                    }
                )
            }
            Mock Get-DefaultAclHC {
                @{
                    'Bob' = 'R'
                }
            }
            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'test'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Mike' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = '' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = '' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @testPermissionsParams

            .$testScript @testParams

            $actual = ($ImportedMatrix.Settings.Matrix.Where( { $_.Path -eq 'Path' })).ACL

            $expected = @{
                'Bob'  = 'L'
                'Mike' = 'L'
            }

            Compare-HashTableHC $actual $expected
        }
    }
}
Describe 'when a FatalError occurs while executing the matrix' {
    AfterEach {
        $Error.Clear()
        Remove-Item -Path "$($testLogFolder)\*" -Recurse -Force -EA Ignore
        Remove-Item -Path "$($testInputFile.Matrix.FolderPath)\*" -Exclude 'Default.xlsx' -Recurse -Force -EA Ignore
    }
    It 'a detailed HTML log file is created for each settings row' {
        $testProblem = @{
            Type = 'FatalError'
            Name = 'Matrix setting incorrect'
        }
        Mock Test-MatrixSettingHC {
            $testProblem
        }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S2'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams
        $testPermissions | Export-Excel @testPermissionsParams

        .$testScript @testParams

        $testMatrixLogFolder = Get-ChildItem -Path $testLogFolder -Directory
        @(Get-ChildItem -Path $testMatrixLogFolder.FullName -File | Where-Object Extension -NE '.xlsx').Count | Should -BeExactly 2
    }
    It 'a TXT log file is created for each settings row when there are more than 5 elements in the value array' {
        $testProblem = @{
            Type        = 'FatalError'
            Name        = 'Matrix setting incorrect'
            Description = 'When things go south we need to report it. because it needs to get fixed'
            Value       = @('C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1')
        }
        Mock Test-MatrixSettingHC {
            $testProblem
        }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
                GroupName    = 'Group'
                SiteCode     = 'Site'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S2'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams
        $testPermissions | Export-Excel @testPermissionsParams

        .$testScript @testParams

        $testMatrixLogFolder = Get-ChildItem -Path $testLogFolder -Directory
        @(Get-ChildItem -Path $testMatrixLogFolder.FullName -File | Where-Object Extension -EQ '.txt').Count | Should -BeExactly 2
    }
    It 'an e-mail is send' {
        $testProblem = @{
            Type        = 'FatalError'
            Name        = 'Matrix setting incorrect'
            Description = 'When things go south we need to report it. because it needs to get fixed'
            Value       = @('C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1')
        }
        Mock Test-MatrixSettingHC {
            $testProblem
        }

        @(
            [PSCustomObject]@{Status = 'Enabled'
                ComputerName         = 'S1'
                Path                 = 'E:\Department'
                Action               = 'Check'
                GroupName            = 'Group'
                SiteCode             = 'Site'
            }
            [PSCustomObject]@{Status = 'Enabled'
                ComputerName         = 'S2'
                Path                 = 'E:\Department'
                Action               = 'Check'
            }
        ) | Export-Excel @testSettingsParams
        $testPermissions | Export-Excel @testPermissionsParams

        .$testScript @testParams

        Should -Invoke Send-MailKitMessageHC -Scope it -Times 1 -Exactly
    }
}
Describe 'when Export.ServiceNowFormDataExcelFile is used but' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Export.ServiceNowFormDataExcelFile = 'TestDrive:/snow.xlsx'

        Test-NewJsonFileHC
    }
    Context 'the Excel file is missing the sheet FormData' {
        BeforeAll {
            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams

            $testPermissions | Export-Excel @testPermissionsParams

            .$testScript @testParams
        }
        It 'a FatalError is registered for the file' {
            $actual = $ImportedMatrix.File.Check
            $actual.Type | Should -Contain 'FatalError'
            $actual.Name | Should -Contain "Worksheet 'FormData' not found"
        }
        It 'the permissions script is not executed' {
            Should -Not -Invoke Invoke-Command
        }
        It 'an email is sent to the user with the error' {
            Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Context -ParameterFilter {
                ($From -eq 'm@example.com') -and
                ($To[0] -eq '007@example.com') -and
                ($To[1] -eq 'bob@contoso.com') -and
                ($SmtpPort -eq 25) -and
                ($SmtpServerName -eq 'SMTP_SERVER') -and
                ($SmtpConnectionType -eq 'StartTls') -and
                ($Subject -eq '1 matrix file, 1 error, Email subject') -and
                ($Priority -eq 'High') -and
                ($Body -like "*Worksheet 'FormData' not found*")
            }
        }
    }
    Context 'the worksheet FormData contains incorrect data' {
        AfterAll {
            Mock Test-FormDataHC
            Mock Get-AdUserPrincipalNameHC
        }
        BeforeAll {
            Mock Test-ExpandedMatrixHC
            Mock Test-FormDataHC {
                @{
                    Type = 'FatalError'
                    Name = 'incorrect data'
                }
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams

            @(
                [PSCustomObject]@{
                    MatrixFormStatus = 'Enabled'
                }
            ) |
            Export-Excel -Path $testSettingsParams.Path -WorksheetName 'FormData'

            $testPermissions | Export-Excel @testPermissionsParams            

            .$testScript @testParams
        }
        It 'a FatalError is registered for the FormData sheet' {
            $actual = $ImportedMatrix.FormData.Check
            $actual.Type | Should -Contain 'FatalError'
            $actual.Name | Should -Contain 'incorrect data'
        }
        It 'the permissions script is not executed' {
            Should -Not -Invoke Invoke-Command
        }
        It 'an email is sent to the user with the error' {
            Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Context -ParameterFilter {
                ($From -eq 'm@example.com') -and
                ($To[0] -eq '007@example.com') -and
                ($To[1] -eq 'bob@contoso.com') -and
                ($SmtpPort -eq 25) -and
                ($SmtpServerName -eq 'SMTP_SERVER') -and
                ($SmtpConnectionType -eq 'StartTls') -and
                ($Subject -eq '1 matrix file, 1 error, Email subject') -and
                ($Priority -eq 'High') -and
                ($Body -like '*Errors*Warnings*FormData*') -and
                ($Body -like '*FormData*incorrect data*') -and
                ($Body -notlike '*Check the*overview*for details*')
            }
        }
    }
    Context 'the worksheet FormData has a non existing MatrixResponsible' {
        BeforeAll {
            Mock Test-ExpandedMatrixHC
            Mock Test-FormDataHC
            Mock Get-AdUserPrincipalNameHC {
                @{
                    UserPrincipalName = 'mike@contoso.com'
                    notFound          = 'bob@contoso.com'
                }
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @testSettingsParams

            @(
                [PSCustomObject]@{
                    MatrixFormStatus  = 'Enabled'
                    MatrixResponsible = 'mike@contoso.com, bob@contoso.com'
                }
            ) |
            Export-Excel -Path $testSettingsParams.Path -WorksheetName 'FormData'

            $testPermissions | Export-Excel @testPermissionsParams

            .$testScript @testParams
        }
        It 'a Warning is registered for the FormData sheet' {
            $actual = $ImportedMatrix.FormData.Check
            $actual.Type | Should -Contain 'Warning'
            $actual.Name | Should -Contain 'AD object not found'
            $actual.Value | Should -Contain 'bob@contoso.com'
        }
        It 'the permissions script is not executed' {
            Should -Not -Invoke Invoke-Command
        }
        It 'an email is sent to the user with the warning message' {
            Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Context -ParameterFilter {
                ($From -eq 'm@example.com') -and
                ($To[0] -eq '007@example.com') -and
                ($To[1] -eq 'bob@contoso.com') -and
                ($SmtpPort -eq 25) -and
                ($SmtpServerName -eq 'SMTP_SERVER') -and
                ($SmtpConnectionType -eq 'StartTls') -and
                ($Subject -eq '1 matrix file, 1 warning, Email subject') -and
                ($Priority -eq 'High') -and
                ($Body -like (
                    '*<th id="matrixHeader" colspan="8">FormData</th> <tr>
                    <td id="probTypeWarning"></td>
                    <p id="probTitle">AD object not found</p>
                    <p>The email address or SamAccountName is not found in the active directory. Multiple entries are supported with the comma ', ' separator.</p>*
                    <ul><li>bob@contoso.com</li></ul>*') -replace '[\r\n]+', '*'
                )
            }
        }
    }
} 
Describe 'when Export.ServiceNowFormDataExcelFile is used' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        Mock Get-AdUserPrincipalNameHC {
            @{
                UserPrincipalName = @('bob@contoso.com', 'mike@contoso.com')
                notFound          = $null
            }
        }
        Mock Test-FormDataHC
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                samAccountName = 'A B C'
                adObject       = @{
                    ObjectClass    = 'group'
                    Name           = 'A B C'
                    SamAccountName = 'A B c'
                    ManagedBy      = 'CN=CaptainManagers,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Jean Luc Picard'
                        SamAccountName = 'picard'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'SamAccountName' }
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                DistinguishedName = 'CN=CaptainManagers,DC=contoso,DC=net'
                adObject          = @{
                    ObjectClass = 'group'
                    Name        = 'Captain Managers'
                }
                adGroupMember     = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Admiral Pike'
                        SamAccountName = 'pike'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'DistinguishedName' }

        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'C' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @testPermissionsParams

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'SERVER1'
                GroupName    = 'A'
                SiteCode     = 'B'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams

        @(
            [PSCustomObject]@{
                MatrixFormStatus        = 'Enabled'
                MatrixCategoryName      = 'a'
                MatrixSubCategoryName   = 'b'
                MatrixResponsible       = 'c'
                MatrixFolderDisplayName = 'd'
                MatrixFolderPath        = 'e'
            }
        ) |
        Export-Excel -Path $testSettingsParams.Path -WorksheetName 'FormData'

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Export.ServiceNowFormDataExcelFile = 'TestDrive:/snow.xlsx'

        Test-NewJsonFileHC
        
        .$testScript @testParams

        $testSnowExcelLogFile = Get-ChildItem $testLogFolder -Recurse -File |
        Where-Object { $_.Name -like '* - Export - ServiceNowFormData.xlsx' }
    }
    Context 'the data in worksheet FormData' {
        It 'is verified to be correct' {
            Should -Invoke Test-FormDataHC -Exactly 1 -Scope Describe
        }
        It 'is retrieved regardless the MatrixFormStatus' {
            $actual = $ImportedMatrix.FormData.Import
            $actual | Should -Not -BeNullOrEmpty
        }
        It 'the MatrixResponsible is converted to UserPrincipalName' {
            Should -Invoke Get-AdUserPrincipalNameHC -Exactly 1 -Scope Describe
        }
    }
    Context 'the FormData is exported' {
        It 'to an Excel file in the Export folder' {
            $testNewInputFile.Export.ServiceNowFormDataExcelFile | 
            Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the log folder' {
            $testSnowExcelLogFile | Should -Not -BeNullOrEmpty
        }
        Context 'with the property' {
            BeforeAll {
                $actual = @{
                    logFolder    = @{
                        Excel = Import-Excel -Path $testSnowExcelLogFile -WorksheetName 'SnowFormData'
                    }
                    exportFolder = @{
                        Excel = Import-Excel -Path $testNewInputFile.Export.ServiceNowFormDataExcelFile -WorksheetName 'SnowFormData'
                    }
                }
            }
            It '<Name>' -ForEach @(
                @{ Name = 'u_adobjectname'; Value = 'A B C' }
                @{ Name = 'u_matrixresponsible'; Value = 'bob@contoso.com,mike@contoso.com' }
                @{ Name = 'u_matrixsubcategoryname'; Value = 'b' }
                @{ Name = 'u_matrixfolderpath'; Value = 'e' }
                @{ Name = 'u_matrixfilename'; Value = 'Matrix' }
                @{ Name = 'u_matrixcategoryname'; Value = 'a' }
            ) {
                $actual.exportFolder.Excel.$Name | Should -Be $Value
                $actual.logFolder.Excel.$Name | Should -Be $Value
            }
        }
    }
    It 'an email is sent to the user in the default settings file' {
        Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($From -eq 'm@example.com') -and
            ($To[0] -eq '007@example.com') -and
            ($To[1] -eq 'bob@contoso.com') -and
            ($SmtpPort -eq 25) -and
            ($SmtpServerName -eq 'SMTP_SERVER') -and
            ($SmtpConnectionType -eq 'StartTls') -and
            ($Subject -eq '1 matrix file, Email subject') -and
            ($Body -like '*<p><b>Exported 1 file:</b></p>*') -and
            ($Body -like '*Matrix results per file*')
        }
    }
}
Describe 'when Export.PermissionsExcelFile is used' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        Mock Get-AdUserPrincipalNameHC {
            @{
                UserPrincipalName = @('bob@contoso.com', 'mike@contoso.com')
                notFound          = $null
            }
        }
        Mock Test-FormDataHC
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                samAccountName = 'A B C'
                adObject       = @{
                    ObjectClass    = 'group'
                    Name           = 'A B C'
                    SamAccountName = 'A B c'
                    ManagedBy      = 'CN=CaptainManagers,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Jean Luc Picard'
                        SamAccountName = 'picard'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'SamAccountName' }
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                DistinguishedName = 'CN=CaptainManagers,DC=contoso,DC=net'
                adObject          = @{
                    ObjectClass = 'group'
                    Name        = 'Captain Managers'
                }
                adGroupMember     = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Admiral Pike'
                        SamAccountName = 'pike'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'DistinguishedName' }

        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'C' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @testPermissionsParams

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'SERVER1'
                GroupName    = 'A'
                SiteCode     = 'B'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams

        @(
            [PSCustomObject]@{
                MatrixFormStatus        = 'Enabled'
                MatrixCategoryName      = 'a'
                MatrixSubCategoryName   = 'b'
                MatrixResponsible       = 'c'
                MatrixFolderDisplayName = 'd'
                MatrixFolderPath        = 'e'
            }
        ) |
        Export-Excel -Path $testSettingsParams.Path -WorksheetName 'FormData'

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Export.PermissionsExcelFile = 'TestDrive:/permissions.xlsx'

        Test-NewJsonFileHC
        
        .$testScript @testParams

        $testPermissionsExcelLogFile = Get-ChildItem $testLogFolder -Recurse -File |
        Where-Object { $_.Name -like '* - Export - Permissions.xlsx' }
    }
    Context 'the AD object names are exported' {
        It 'to an Excel file in the Export folder' {
            $testNewInputFile.Export.PermissionsExcelFile | 
            Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the log folder' {
            $testPermissionsExcelLogFile | Should -Not -BeNullOrEmpty
        }
        Context 'with the property' {
            BeforeAll {
                $actual = @{
                    logFolder    = @{
                        Excel = Import-Excel -Path $testPermissionsExcelLogFile -WorksheetName 'AdObjects'
                    }
                    exportFolder = @{
                        Excel = Import-Excel -Path $testNewInputFile.Export.PermissionsExcelFile -WorksheetName 'AdObjects'
                    }
                }
            }
            It '<Name>' -ForEach @(
                @{ Name = 'MatrixFileName'; Value = 'Matrix' }
                @{ Name = 'SamAccountName'; Value = 'A B C' }
                @{ Name = 'GroupName'; Value = 'A' }
                @{ Name = 'SiteCode'; Value = 'B' }
                @{ Name = 'Name'; Value = 'C' }
            ) {
                $actual.exportFolder.Excel.$Name | Should -Be $Value
                $actual.logFolder.Excel.$Name | Should -Be $Value
            }
        }
    }
    Context 'the GroupManagers are exported' {
        It 'to an Excel file in the Export folder' {
            $testNewInputFile.Export.PermissionsExcelFile | 
            Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the log folder' {
            $testPermissionsExcelLogFile | Should -Not -BeNullOrEmpty
        }
        Context 'with the property' {
            BeforeAll {
                $actual = @{
                    logFolder    = @{
                        Excel = Import-Excel -Path $testPermissionsExcelLogFile -WorksheetName 'GroupManagers'
                    }
                    exportFolder = @{
                        Excel = Import-Excel -Path $testNewInputFile.Export.PermissionsExcelFile -WorksheetName 'GroupManagers'
                    }
                }
            }
            It '<Name>' -ForEach @(
                @{ Name = 'MatrixFileName'; Value = 'Matrix' }
                @{ Name = 'GroupName'; Value = 'A B C' }
                @{ Name = 'ManagerName'; Value = 'Captain Managers' }
                @{ Name = 'ManagerType'; Value = 'group' }
                @{ Name = 'ManagerMemberName'; Value = 'Admiral Pike' }
            ) {
                $actual.exportFolder.Excel.$Name | Should -Be $Value
                $actual.logFolder.Excel.$Name | Should -Be $Value
            }
        }
    }
    Context 'the AccessList are exported' {
        It 'to an Excel file in the Export folder' {
            $testNewInputFile.Export.PermissionsExcelFile | 
            Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the log folder' {
            $testPermissionsExcelLogFile | Should -Not -BeNullOrEmpty
        }
        Context 'with the property' {
            BeforeAll {
                $actual = @{
                    logFolder    = @{
                        Excel = Import-Excel -Path $testPermissionsExcelLogFile -WorksheetName 'AccessList'
                    }
                    exportFolder = @{
                        Excel = Import-Excel -Path $testNewInputFile.Export.PermissionsExcelFile -WorksheetName 'AccessList'
                    }
                }
            }
            It '<Name>' -ForEach @(
                @{ Name = 'MatrixFileName'; Value = 'Matrix' }
                @{ Name = 'SamAccountName'; Value = 'A B C' }
                @{ Name = 'Name'; Value = 'A B C' }
                @{ Name = 'Type'; Value = 'group' }
                @{ Name = 'MemberName'; Value = 'Jean Luc Picard' }
                @{ Name = 'MemberSamAccountName'; Value = 'picard' }
            ) {
                $actual.exportFolder.Excel.$Name | Should -Be $Value
                $actual.logFolder.Excel.$Name | Should -Be $Value
            }
        }
    }
    It 'an email is sent to the user in the default settings file' {
        Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($From -eq 'm@example.com') -and
            ($To[0] -eq '007@example.com') -and
            ($To[1] -eq 'bob@contoso.com') -and
            ($SmtpPort -eq 25) -and
            ($SmtpServerName -eq 'SMTP_SERVER') -and
            ($SmtpConnectionType -eq 'StartTls') -and
            ($Subject -eq '1 matrix file, Email subject') -and
            ($Body -like '*<p><b>Exported 1 file:</b></p>*') -and
            ($Body -like '*Matrix results per file*')
        }
    }
}
Describe 'when Export.OverviewHtmlFile is used' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        Mock Get-AdUserPrincipalNameHC {
            @{
                UserPrincipalName = @('bob@contoso.com', 'mike@contoso.com')
                notFound          = $null
            }
        }
        Mock Test-FormDataHC
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                samAccountName = 'A B C'
                adObject       = @{
                    ObjectClass    = 'group'
                    Name           = 'A B C'
                    SamAccountName = 'A B c'
                    ManagedBy      = 'CN=CaptainManagers,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Jean Luc Picard'
                        SamAccountName = 'picard'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'SamAccountName' }
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                DistinguishedName = 'CN=CaptainManagers,DC=contoso,DC=net'
                adObject          = @{
                    ObjectClass = 'group'
                    Name        = 'Captain Managers'
                }
                adGroupMember     = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Admiral Pike'
                        SamAccountName = 'pike'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'DistinguishedName' }

        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'C' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @testPermissionsParams

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'SERVER1'
                GroupName    = 'A'
                SiteCode     = 'B'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @testSettingsParams

        @(
            [PSCustomObject]@{
                MatrixFormStatus        = 'Enabled'
                MatrixCategoryName      = 'a'
                MatrixSubCategoryName   = 'b'
                MatrixResponsible       = 'c'
                MatrixFolderDisplayName = 'd'
                MatrixFolderPath        = 'e'
            }
        ) |
        Export-Excel -Path $testSettingsParams.Path -WorksheetName 'FormData'

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Export.OverviewHtmlFile = 'TestDrive:/permissions.html'

        Test-NewJsonFileHC
        
        .$testScript @testParams

        $testOverviewHtmlFile = Get-ChildItem $testLogFolder -Recurse -File |
        Where-Object { $_.Name -like '* - Export - Permissions.xlsx' }
    }
    Context 'the overview html file is exported is created in the' {
        It 'Export folder' {
            $testNewInputFile.Export.OverviewHtmlFile | 
            Should -Not -BeNullOrEmpty
        }
        It 'log folder' {
            $testOverviewHtmlFile | Should -Not -BeNullOrEmpty
        }
    } -Tag test
    It 'an email is sent to the user in the default settings file' {
        Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($From -eq 'm@example.com') -and
            ($To[0] -eq '007@example.com') -and
            ($To[1] -eq 'bob@contoso.com') -and
            ($SmtpPort -eq 25) -and
            ($SmtpServerName -eq 'SMTP_SERVER') -and
            ($SmtpConnectionType -eq 'StartTls') -and
            ($Subject -eq '1 matrix file, Email subject') -and
            ($Body -like '*<p><b>Exported 1 file:</b></p>*') -and
            ($Body -like '*Matrix results per file*')
        }
    }
}