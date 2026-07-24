#Requires -Version 7
#Requires -Modules Microsoft.Graph.Authentication

<#
    .SYNOPSIS
        Uploads a local file to a SharePoint Online document library,
        overwriting the file when it is already present.

    .DESCRIPTION
        Connects to Microsoft Graph with an Entra ID app registration using
        certificate based (app-only) authentication and uploads a single local
        file to a document library in a SharePoint Online site.

        The upload always overwrites: when a file with the same name already
        exists in the target folder its content is replaced (SharePoint keeps
        the previous content as a version, so nothing is lost). No error is
        raised for an existing file and no error is raised when the file is
        not there yet.

        Two upload strategies are used, chosen automatically on file size:

            - Files up to 4 MB are sent in a single PUT to the /content
              endpoint, which replaces existing content by default.
            - Larger files use a resumable upload session created with
              conflictBehavior 'replace', uploaded in sequential chunks.

        Any missing folders in FolderPath are created before the upload, so
        the library does not need to be prepared by hand.

        Security Feature: like the other Operations scripts, every string
        parameter supports dynamically fetching secrets from the host's
        environment variables to avoid hardcoding them. Simply prefix the
        value with 'ENV:' (e.g. -ClientId 'ENV:AZURE_CLIENT_ID').

        Required Entra ID application permission (admin consented):
        'Sites.ReadWrite.All', or 'Sites.Selected' with write access granted
        on the target site.

    .PARAMETER FilePath
        The absolute path to the local file to upload (e.g. the generated
        'Overview.html').

    .PARAMETER SiteUrl
        The full URL of the SharePoint Online site that hosts the document
        library (e.g. 'https://contoso.sharepoint.com/sites/IT').

    .PARAMETER DocumentLibraryName
        The display name of the target document library
        (e.g. 'Documents' or 'Shared Documents').

    .PARAMETER FolderPath
        (Optional) A folder path inside the document library, relative to its
        root (e.g. 'Reports/Permission matrix'). Folders that do not exist are
        created. When omitted the file is uploaded to the library root.

    .PARAMETER FileName
        (Optional) The name to give the file in SharePoint. Defaults to the
        file name of FilePath.

    .PARAMETER ClientId
        The application (client) ID of the Entra ID app registration.
        Supports 'ENV:AZURE_CLIENT_ID'.

    .PARAMETER TenantId
        The Entra ID tenant ID.
        Supports 'ENV:AZURE_TENANT_ID'.

    .PARAMETER CertificateThumbprint
        The thumbprint of the certificate, installed in the certificate store
        of the executing account, used for app-only authentication.
        Supports 'ENV:AZURE_POWERSHELL_CERTIFICATE_THUMBPRINT'.

    .PARAMETER MaxRetries
        The maximum number of attempts for each Graph call before giving up.
        Failed attempts are 3 seconds apart. (Default: 3, minimum 1)

    .PARAMETER ChunkSizeMB
        The size in MB of a single chunk in a resumable upload session. Must be
        a multiple of 320 KiB, which every allowed value here satisfies.
        Only used for files larger than 4 MB. (Default: 5)

    .EXAMPLE
        .\UploadToSharePoint.ps1 `
            -FilePath 'C:\reports\Overview.html' `
            -SiteUrl 'https://contoso.sharepoint.com/sites/IT' `
            -DocumentLibraryName 'Documents' `
            -FolderPath 'Reports/Permission matrix' `
            -ClientId 'ENV:AZURE_CLIENT_ID' `
            -TenantId 'ENV:AZURE_TENANT_ID' `
            -CertificateThumbprint 'ENV:AZURE_POWERSHELL_CERTIFICATE_THUMBPRINT'

        Uploads 'Overview.html' to the folder 'Reports/Permission matrix' in
        the 'Documents' library. The folder is created when it does not exist
        and the file is overwritten when it is already there.

    .EXAMPLE
        .\UploadToSharePoint.ps1 `
            -FilePath 'C:\reports\Overview.html' `
            -SiteUrl 'https://contoso.sharepoint.com/sites/IT' `
            -DocumentLibraryName 'Documents' `
            -FileName 'Permission matrix overview.html' `
            -ClientId $env:AZURE_CLIENT_ID `
            -TenantId $env:AZURE_TENANT_ID `
            -CertificateThumbprint $env:AZURE_POWERSHELL_CERTIFICATE_THUMBPRINT `
            -Verbose

        Uploads to the root of the library under a different name, with the
        credentials resolved by the caller instead of by this script.
#>

param (
    [Parameter(Mandatory)]
    [String]$FilePath,
    [Parameter(Mandatory)]
    [String]$SiteUrl,
    [Parameter(Mandatory)]
    [String]$DocumentLibraryName,
    [String]$FolderPath,
    [String]$FileName,
    [Parameter(Mandatory)]
    [String]$ClientId,
    [Parameter(Mandatory)]
    [String]$TenantId,
    [Parameter(Mandatory)]
    [String]$CertificateThumbprint,
    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxRetries = 3,
    [ValidateRange(1, 60)]
    [int]$ChunkSizeMB = 5
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
            Get-StringValueHC -Name 'ENV:AZURE_CLIENT_ID'

            # Output: the environment variable value of $ENV:AZURE_CLIENT_ID
            # or an error when the variable does not exist

        .EXAMPLE
            Get-StringValueHC -Name 'myClientId'

            # Output: myClientId
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

    function Invoke-WithRetryHC {
        <#
        .SYNOPSIS
            Run a Graph operation with bounded retries and return its result.

        .DESCRIPTION
            Executes the supplied script block, retrying on failure up to
            MaxRetries attempts with a 3 second pause *between* attempts (never
            after the final attempt). Exhausting the retries rethrows, because
            every call in this script is required for the upload to succeed:
            there is no useful "carry on without it" state.

            Returns whatever the script block returns.

        .PARAMETER Action
            The script block to execute.

        .PARAMETER Description
            A short lower-case description of the operation, used in warning
            and error messages ("Failed to <description>").

        .PARAMETER MaxRetries
            Maximum attempts before giving up and throwing.
        #>
        param (
            [Parameter(Mandatory)][scriptblock]$Action,
            [Parameter(Mandatory)][string]$Description,
            [Parameter(Mandatory)][int]$MaxRetries
        )

        $attempt = 0

        while ($true) {
            $attempt++

            try {
                return & $Action
            }
            catch {
                $errorMessage = $_

                if ($global:Error.Count -gt 0) {
                    $global:Error.RemoveAt(0)
                }

                # 4xx other than 429 means the request itself is wrong: 
                # wrong permissions, wrong site, wrong path. 
                # Retrying cannot fix it and only delays the report.
                if ("$errorMessage" -match 'accessDenied|itemNotFound|unauthenticated|invalidRequest') {
                    throw "Failed to $Description (not retryable): $errorMessage"
                }

                if ($attempt -ge $MaxRetries) {
                    throw "Failed to $Description after $MaxRetries attempts. Last error: $errorMessage"
                }

                Write-Warning "Attempt $attempt of $MaxRetries failed to $Description. Retrying in 3 seconds... Error: $errorMessage"

                Start-Sleep -Seconds 3
            }
        }
    }

    function ConvertTo-GraphPathHC {
        <#
        .SYNOPSIS
            Convert a library relative folder/file path to a Graph safe path.

        .DESCRIPTION
            Splits on both slash types, drops empty segments (so 'a//b/' and
            '\a\b' both become 'a/b') and URL encodes each segment
            individually. Encoding per segment rather than the whole string
            keeps the '/' separators intact while escaping spaces and other
            characters that are legal in SharePoint names but not in a URL.

            Returns an empty string when Path is empty.
        #>
        param (
            [String]$Path
        )

        if ([string]::IsNullOrWhiteSpace($Path)) {
            return ''
        }

        $segments = $Path.Split([char[]]('/', '\'), [StringSplitOptions]::RemoveEmptyEntries)

        ($segments | ForEach-Object {
            [System.Uri]::EscapeDataString($_)
        }) -join '/'
    }

    function Get-GraphSiteIdHC {
        <#
        .SYNOPSIS
            Resolve a SharePoint site URL to its Microsoft Graph site ID.

        .DESCRIPTION
            Graph addresses a site by hostname and server relative path
            ('contoso.sharepoint.com:/sites/IT'), not by its browser URL, so
            the URL is taken apart here. A root site ('https://contoso.
            sharepoint.com') has no path and is addressed by hostname alone.
        #>
        param (
            [Parameter(Mandatory)][string]$SiteUrl,
            [Parameter(Mandatory)][int]$MaxRetries
        )

        try {
            $uri = [System.Uri]$SiteUrl
            $sitePath = $uri.AbsolutePath.TrimEnd('/')
        }
        catch {
            throw "SharePoint site URL '$SiteUrl' is not a valid URL: $_"
        }

        $graphUri = if ($sitePath) {
            "https://graph.microsoft.com/v1.0/sites/$($uri.Host):$($sitePath)"
        }
        else {
            "https://graph.microsoft.com/v1.0/sites/$($uri.Host)"
        }

        Write-Verbose "Resolve SharePoint site '$SiteUrl'"

        $site = Invoke-WithRetryHC -MaxRetries $MaxRetries -Description "resolve SharePoint site '$SiteUrl'" -Action {
            Invoke-MgGraphRequest -Method 'GET' -Uri $graphUri -OutputType 'PSObject'
        }

        if (-not $site.id) {
            throw "SharePoint site '$SiteUrl' not found."
        }

        Write-Verbose "Site ID '$($site.id)'"

        $site.id
    }

    function Get-GraphDriveIdHC {
        <#
        .SYNOPSIS
            Find a document library on a site by its display name.

        .DESCRIPTION
            Each document library is exposed as a drive. The list is paged
            through in full before matching, so a site with many libraries
            still resolves correctly. Matching is case-insensitive.

            Note that the default library is named 'Documents' in Graph even
            though its URL says 'Shared Documents'; both names are accepted
            here to save a support call.
        #>
        param (
            [Parameter(Mandatory)][string]$SiteId,
            [Parameter(Mandatory)][string]$Name,
            [Parameter(Mandatory)][int]$MaxRetries
        )

        Write-Verbose "Find document library '$Name'"

        $drives = [System.Collections.Generic.List[object]]::new()
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives"

        while ($uri) {
            $requestUri = $uri

            $response = Invoke-WithRetryHC -MaxRetries $MaxRetries -Description "list document libraries of site '$SiteId'" -Action {
                Invoke-MgGraphRequest -Method 'GET' -Uri $requestUri -OutputType 'PSObject'
            }

            if ($response.value) {
                $drives.AddRange(@($response.value))
            }

            $uri = $response.'@odata.nextLink'
        }

        $acceptedName = @($Name)

        if ($Name -eq 'Shared Documents') { $acceptedName += 'Documents' }
        if ($Name -eq 'Documents') { $acceptedName += 'Shared Documents' }

        $drive = $drives | Where-Object { $acceptedName -contains $_.name } | Select-Object -First 1

        if (-not $drive) {
            # An empty list is not the same failure as a list without a match:
            # Graph returns zero drives when the caller may see the site but not
            # its content, so an empty result means missing permissions rather
            # than a wrong name. Fall back to the default library endpoint,
            # which answers in some cases where /drives does not.
            if ($drives.Count -eq 0) {
                Write-Verbose "No libraries returned for site '$SiteId', trying the default library"

                try {
                    $defaultDrive = Invoke-WithRetryHC -MaxRetries $MaxRetries -Description "get the default document library of site '$SiteId'" -Action {
                        Invoke-MgGraphRequest -Method 'GET' -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drive" -OutputType 'PSObject'
                    }

                    if ($defaultDrive.id -and ($acceptedName -contains $defaultDrive.name)) {
                        Write-Verbose "Document library ID '$($defaultDrive.id)' (via the default library)"
                        return $defaultDrive.id
                    }
                }
                catch {
                    if ($global:Error.Count -gt 0) { $global:Error.RemoveAt(0) }
                }

                throw "No document libraries are visible to this application on site '$SiteId'. The site itself resolved, so this is an access problem rather than a wrong library name: verify that the app registration has an admin-consented 'Sites.ReadWrite.All' role, or a 'Sites.Selected' grant with the 'write' role on this specific site."
            }

            throw "Document library '$Name' not found on site '$SiteId'. Available libraries: $(($drives.name | Sort-Object) -join ', ')"
        }

        Write-Verbose "Document library ID '$($drive.id)'"

        $drive.id
    }

    function New-GraphFolderPathHC {
        <#
        .SYNOPSIS
            Ensure every folder in a library relative path exists.

        .DESCRIPTION
            Walks the path one segment at a time. For each segment the children
            of the parent folder are listed and, when no folder with that name
            is there, it is created.

            Existence is checked by listing children rather than by requesting
            the item and catching a 404, because that keeps this function from
            depending on the exact shape of a Graph error object, which differs
            between SDK versions.

            A 'nameAlreadyExists' error on creation is swallowed: it means a
            parallel run won the race, which is the outcome we wanted anyway.
        #>
        param (
            [Parameter(Mandatory)][string]$DriveId,
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][int]$MaxRetries
        )

        $segments = $Path.Split([char[]]('/', '\'), [StringSplitOptions]::RemoveEmptyEntries)

        $parentPath = ''

        foreach ($segment in $segments) {
            $encodedParent = ConvertTo-GraphPathHC -Path $parentPath

            $childrenUri = if ($encodedParent) {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($encodedParent):/children"
            }
            else {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children"
            }

            $currentPath = if ($parentPath) { "$parentPath/$segment" } else { $segment }

            $children = Invoke-WithRetryHC -MaxRetries $MaxRetries -Description "list the content of folder '$parentPath'" -Action {
                Invoke-MgGraphRequest -Method 'GET' -Uri $childrenUri -OutputType 'PSObject'
            }

            $existingFolder = $children.value | Where-Object {
                ($_.name -eq $segment) -and ($null -ne $_.folder)
            } | Select-Object -First 1

            if ($existingFolder) {
                Write-Verbose "Folder '$currentPath' already exists"
            }
            else {
                Write-Verbose "Create folder '$currentPath'"

                $body = @{
                    'name'                              = $segment
                    'folder'                            = @{}
                    '@microsoft.graph.conflictBehavior' = 'fail'
                }

                try {
                    $null = Invoke-WithRetryHC -MaxRetries $MaxRetries -Description "create folder '$currentPath'" -Action {
                        Invoke-MgGraphRequest -Method 'POST' -Uri $childrenUri -Body ($body | ConvertTo-Json) -ContentType 'application/json' -OutputType 'PSObject'
                    }
                }
                catch {
                    if ("$_" -notmatch 'nameAlreadyExists') {
                        throw
                    }

                    if ($global:Error.Count -gt 0) {
                        $global:Error.RemoveAt(0)
                    }

                    Write-Verbose "Folder '$currentPath' was created by another process"
                }
            }

            $parentPath = $currentPath
        }
    }

    function Send-GraphFileHC {
        <#
        .SYNOPSIS
            Upload a local file to a drive path, overwriting what is there.

        .DESCRIPTION
            Files up to 4 MB are sent in one PUT to the /content endpoint,
            which replaces the content of an existing item by default and
            creates the item when it is not there yet.

            Larger files go through a resumable upload session created with
            conflictBehavior 'replace'. The chunks are PUT to the session's
            pre-authenticated upload URL with Invoke-RestMethod instead of
            Invoke-MgGraphRequest, because that URL carries its own
            authorisation and must not be sent a bearer token.

            Returns the resulting driveItem.
        #>
        param (
            [Parameter(Mandatory)][string]$DriveId,
            [Parameter(Mandatory)][string]$TargetPath,
            [Parameter(Mandatory)][string]$LocalFilePath,
            [Parameter(Mandatory)][int]$MaxRetries,
            [Parameter(Mandatory)][int]$ChunkSizeMB
        )

        $file = Get-Item -LiteralPath $LocalFilePath
        $encodedTargetPath = ConvertTo-GraphPathHC -Path $TargetPath

        $simpleUploadLimit = 4MB

        if ($file.Length -le $simpleUploadLimit) {
            #region Single request upload
            Write-Verbose "Upload $([math]::Round($file.Length / 1KB, 1)) KB in a single request"

            $uploadUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($encodedTargetPath):/content"

            return Invoke-WithRetryHC -MaxRetries $MaxRetries -Description "upload file to '$TargetPath'" -Action {
                $params = @{
                    Method        = 'PUT'
                    Uri           = $uploadUri
                    InputFilePath = $LocalFilePath
                    ContentType   = 'text/html'
                    OutputType    = 'PSObject'
                }
                Invoke-MgGraphRequest @params
            }
            #endregion
        }

        #region Resumable upload session
        Write-Verbose "Upload $([math]::Round($file.Length / 1MB, 1)) MB in chunks of $ChunkSizeMB MB"

        $sessionUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($encodedTargetPath):/createUploadSession"

        $sessionBody = @{
            'item' = @{
                '@microsoft.graph.conflictBehavior' = 'replace'
            }
        }

        $session = Invoke-WithRetryHC -MaxRetries $MaxRetries -Description "create an upload session for '$TargetPath'" -Action {
            Invoke-MgGraphRequest -Method 'POST' -Uri $sessionUri -Body ($sessionBody | ConvertTo-Json -Depth 3) -ContentType 'application/json' -OutputType 'PSObject'
        }

        if (-not $session.uploadUrl) {
            throw "Failed to create an upload session for '$TargetPath': no upload URL returned."
        }

        $uploadUrl = $session.uploadUrl

        # 320 KiB is the Graph chunk granularity; every chunk except the last
        # must be a multiple of it.
        $chunkSize = $ChunkSizeMB * 320 * 1024 * [math]::Floor(1MB / (320 * 1024))

        if ($chunkSize -le 0) { $chunkSize = 320 * 1024 }

        $result = $null
        $stream = $null

        try {
            $stream = [System.IO.File]::OpenRead($file.FullName)

            $buffer = [byte[]]::new($chunkSize)
            $position = 0

            while ($position -lt $file.Length) {
                $bytesRead = $stream.Read($buffer, 0, $buffer.Length)

                if ($bytesRead -le 0) { break }

                # A short final chunk needs its own buffer. $buffer[0..n] would
                # return Object[] rather than byte[], leaving the request body to
                # rely on coercion; Array.Copy keeps it a genuine byte[].
                $chunk = if ($bytesRead -eq $buffer.Length) {
                    $buffer
                }
                else {
                    $lastChunk = [byte[]]::new($bytesRead)
                    [Array]::Copy($buffer, 0, $lastChunk, 0, $bytesRead)
                    $lastChunk
                }

                $rangeStart = $position
                $rangeEnd = $position + $bytesRead - 1

                $headers = @{
                    'Content-Range' = "bytes $rangeStart-$rangeEnd/$($file.Length)"
                }

                $description = "upload bytes $rangeStart-$rangeEnd of '$TargetPath'"

                Write-Verbose $description

                $result = Invoke-WithRetryHC -MaxRetries $MaxRetries -Description $description -Action {
                    $params = @{
                        Method      = 'PUT'
                        Uri         = $uploadUrl
                        Headers     = $headers
                        Body        = $chunk
                        ContentType = 'application/octet-stream'
                    }
                    Invoke-RestMethod @params
                }

                $position += $bytesRead
            }
        }
        finally {
            if ($stream) { $stream.Dispose() }

            # A half finished session would keep a lock on the target item.
            if ($position -lt $file.Length) {
                try {
                    Invoke-RestMethod -Method 'DELETE' -Uri $uploadUrl -EA Ignore | Out-Null
                }
                catch {
                    if ($global:Error.Count -gt 0) {
                        $global:Error.RemoveAt(0)
                    }
                }
            }
        }

        $result
        #endregion
    }

    $ErrorActionPreference = 'Stop'

    try {
        #region Test the file to upload
        Write-Verbose "Test file to upload '$FilePath'"

        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
            throw "File '$FilePath' not found."
        }

        if (-not $FileName) {
            $FileName = Split-Path -Path $FilePath -Leaf
        }
        #endregion

        #region Get the connection details
        Write-Verbose 'Get the Microsoft Graph connection details'

        $graphClientId = Get-StringValueHC -Name $ClientId
        $graphTenantId = Get-StringValueHC -Name $TenantId
        $graphThumbprint = Get-StringValueHC -Name $CertificateThumbprint

        @{
            ClientId              = $graphClientId
            TenantId              = $graphTenantId
            CertificateThumbprint = $graphThumbprint
        }.GetEnumerator().where(
            { -not $_.Value }
        ).foreach(
            { throw "Parameter '$($_.Key)' resolved to an empty value." }
        )
        #endregion
    }
    catch {
        throw "SharePoint upload of file '$FilePath': $_"
    }

    #region Connect to MS Graph
    # Reuse the session when the process is already connected as the same
    # application to the same tenant, so repeated calls in one run do not
    # trigger a new authentication each time.
    $mgContext = Get-MgContext

    if (
        $mgContext -and
        ($mgContext.ClientId -eq $graphClientId) -and
        ($mgContext.TenantId -eq $graphTenantId)
    ) {
        Write-Verbose 'Already connected to MS Graph'
    }
    else {
        Write-Verbose 'Connect to MS Graph'

        try {
            $params = @{
                ClientId              = $graphClientId
                TenantId              = $graphTenantId
                CertificateThumbprint = $graphThumbprint
                NoWelcome             = $true
            }
            Connect-MgGraph @params
        }
        catch {
            $errorMessage = $_

            if ($global:Error.Count -gt 0) {
                $global:Error.RemoveAt(0)
            }

            throw "Failed to connect to MS Graph with ClientId '$graphClientId' TenantId '$graphTenantId' CertificateThumbprint '$graphThumbprint': $errorMessage"
        }
    }
    #endregion
}

process {
    #region Resolve the target document library
    $siteId = Get-GraphSiteIdHC -SiteUrl $SiteUrl -MaxRetries $MaxRetries

    $driveId = Get-GraphDriveIdHC -SiteId $siteId -Name $DocumentLibraryName -MaxRetries $MaxRetries
    #endregion

    #region Create the folder structure when needed
    if ($FolderPath) {
        New-GraphFolderPathHC -DriveId $driveId -Path $FolderPath -MaxRetries $MaxRetries
    }
    #endregion

    #region Upload the file, overwriting an existing one
    $targetPath = if ($FolderPath) {
        '{0}/{1}' -f $FolderPath.Trim([char[]]('/', '\')), $FileName
    }
    else {
        $FileName
    }

    Write-Verbose "Upload file '$FilePath' to '$SiteUrl' library '$DocumentLibraryName' path '$targetPath'"

    $params = @{
        DriveId       = $driveId
        TargetPath    = $targetPath
        LocalFilePath = $FilePath
        MaxRetries    = $MaxRetries
        ChunkSizeMB   = $ChunkSizeMB
    }
    $uploadedItem = Send-GraphFileHC @params

    Write-Verbose "Upload complete: '$($uploadedItem.webUrl)'"
    #endregion

    [PSCustomObject]@{
        Name    = $uploadedItem.name
        WebUrl  = $uploadedItem.webUrl
        Id      = $uploadedItem.id
        DriveId = $driveId
        SiteId  = $siteId
    }
}