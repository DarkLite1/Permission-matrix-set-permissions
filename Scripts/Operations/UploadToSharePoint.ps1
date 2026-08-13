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
        the previous content as a version, so nothing is lost). Neither an
        existing file nor a missing one raises an error.

        Two upload strategies are chosen automatically on file size:
            - Up to 4 MB: a single PUT to the /content endpoint, which replaces
              existing content by default.
            - Larger: a resumable upload session with conflictBehavior
              'replace', uploaded in sequential chunks.

        Missing folders in FolderPath are created before the upload, so the
        library does not need to be prepared by hand.

        Required Entra ID application permission (admin consented):
        'Sites.ReadWrite.All', or 'Sites.Selected' with write access granted on
        the target site.

    .PARAMETER FolderPath
        Optional folder path inside the document library, relative to its root
        (e.g. 'Reports/Permission matrix'). Missing folders are created. When
        omitted the file lands in the library root.

    .PARAMETER FileName
        Optional name to give the file in SharePoint. Defaults to the file name
        of FilePath.

    .PARAMETER ClientId
        Application (client) ID of the Entra ID app registration.

    .PARAMETER CertificateThumbprint
        Thumbprint of the certificate, installed in the certificate store of the
        executing account, used for app-only authentication.

    .PARAMETER MaxRetries
        Maximum attempts per Graph call before giving up. (Default: 3)

    .PARAMETER ChunkSizeMB
        Size in MB of a single chunk in a resumable upload session. Must be a
        multiple of 320 KiB, which every allowed value satisfies. Only used for
        files larger than 4 MB. (Default: 5)

    .NOTES
        Like the other Operations scripts, EVERY string parameter accepts an
        'ENV:' prefix to read the value from the host's environment variables
        instead of hardcoding it (e.g. -ClientId 'ENV:AZURE_CLIENT_ID').

    .EXAMPLE
        .\UploadToSharePoint.ps1 `
            -FilePath 'C:\reports\Overview.html' `
            -SiteUrl 'https://contoso.sharepoint.com/sites/IT' `
            -DocumentLibraryName 'Documents' `
            -FolderPath 'Reports/Permission matrix' `
            -ClientId 'ENV:AZURE_CLIENT_ID' `
            -TenantId 'ENV:AZURE_TENANT_ID' `
            -CertificateThumbprint 'ENV:AZURE_POWERSHELL_CERTIFICATE_THUMBPRINT'

        Uploads to 'Reports/Permission matrix' in the 'Documents' library,
        creating the folder when absent and overwriting an existing file.
        Credentials are resolved by the script from environment variables.

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

        Uploads to the library root under a different name, with the credentials
        resolved by the caller instead of by this script.
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
            A Name starting with 'ENV:' (case-insensitive) is treated as an
            environment variable name; its value is returned, or the function
            throws when the variable does not exist. Anything else is returned
            unchanged, and a blank Name returns $null.

        .EXAMPLE
            Get-StringValueHC -Name 'ENV:AZURE_CLIENT_ID'

            Returns the value of $env:AZURE_CLIENT_ID, or throws when it is not
            set.
        #>
        param (
            [String]$Name
        )

        if ([string]::IsNullOrWhiteSpace($Name)) {
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

    function Test-IsRetryableErrorHC {
        <#
        .SYNOPSIS
            Decide whether a failed Graph call is worth attempting again.

        .DESCRIPTION
            Client errors mean the request itself is wrong: the app lacks the
            permission, the site does not exist, the path is malformed. Repeating
            an identical request cannot change any of that, so retrying only
            delays the report by MaxRetries * the backoff.

            Throttling (429) and server-side failures (5xx) are the opposite:
            the request was fine and a later attempt has a real chance.

            Returns $true when the error is worth retrying.
        #>
        param (
            [Parameter(Mandatory)][string]$ErrorText
        )

        # 429 is a client error but IS retryable, so it is checked first.
        if ($ErrorText -match '\b429\b|activityLimitReached|TooManyRequests|throttl') {
            return $true
        }

        $nonRetryable = @(
            'accessDenied'          # no permission on the site or library
            'itemNotFound'          # wrong site, library, or path
            'unauthenticated'       # token rejected
            'invalidRequest'        # malformed request
            'invalidRange'          # bad Content-Range on a chunk
            'nameAlreadyExists'     # handled by the caller, never retried here
            'malwareDetected'
            'quotaLimitReached'
            'resourceModified'
            '\b40[0-9]\b'           # any other 4xx
        )

        foreach ($pattern in $nonRetryable) {
            if ($ErrorText -match $pattern) {
                return $false
            }
        }

        # Unknown failures (network blips, 5xx, transient DNS) get the benefit of
        # the doubt: retrying them is cheap and often works.
        $true
    }

    function Get-RetryAfterSecondsHC {
        <#
        .SYNOPSIS
            Read a Retry-After hint out of a Graph error, if it carries one.

        .DESCRIPTION
            When Graph throttles it returns a Retry-After header saying how long
            to wait. Ignoring it and retrying sooner makes the throttling worse,
            so it is honoured when present.

            The header is read from the exception's response where the SDK
            surfaces it, and otherwise scraped from the message text. Returns 0
            when no hint is available, leaving the caller to fall back to its own
            backoff.
        #>
        param (
            $ErrorRecord
        )

        try {
            $response = $ErrorRecord.Exception.Response

            if ($response -and $response.Headers) {
                $retryAfter = $response.Headers.RetryAfter

                if ($retryAfter.Delta.TotalSeconds -gt 0) {
                    return [int][math]::Ceiling($retryAfter.Delta.TotalSeconds)
                }
                if ($retryAfter.Date) {
                    $seconds = ($retryAfter.Date - (Get-Date)).TotalSeconds
                    if ($seconds -gt 0) { return [int][math]::Ceiling($seconds) }
                }
            }
        }
        catch {
            # The exception shape varies by SDK version. Failing to read a hint
            # is not itself a problem: fall through to scraping the message and
            # then to the caller's own backoff.
        }

        if ("$ErrorRecord" -match 'Retry-After[:\s]+(\d+)') {
            return [int]$Matches[1]
        }

        0
    }

    function Invoke-WithRetryHC {
        <#
        .SYNOPSIS
            Run a Graph operation with bounded retries and return its result.

        .DESCRIPTION
            Executes the supplied script block, retrying up to MaxRetries
            attempts. Exhausting the retries rethrows, because every call in
            this script is required for the upload to succeed: there is no
            useful "carry on without it" state.

            Three things decide what happens after a failure:

            - Errors that cannot succeed on a repeat (403, 404, malformed
              requests) throw immediately. Retrying a missing Sites.Selected
              grant three times just delays an accurate error by nine seconds.
            - A Retry-After hint from a throttling response is honoured exactly,
              because retrying sooner deepens the throttling.
            - Everything else backs off exponentially (RetryDelaySeconds, then
              doubling, capped at 60s) rather than hammering at a fixed
              interval.

            Pauses happen only *between* attempts, never after the last one.

        .PARAMETER Description
            A short lower-case description of the operation, used in warning and
            error messages ("Failed to <description>").

        .PARAMETER RetryDelaySeconds
            Base delay for the first retry. Each subsequent wait doubles it.
        #>
        param (
            [Parameter(Mandatory)][scriptblock]$Action,
            [Parameter(Mandatory)][string]$Description,
            [Parameter(Mandatory)][int]$MaxRetries,
            [int]$RetryDelaySeconds = 3
        )

        $attempt = 0
        $maxDelaySeconds = 60

        while ($true) {
            $attempt++

            try {
                return & $Action
            }
            catch {
                $errorRecord = $_
                $errorMessage = "$_"

                if (-not (Test-IsRetryableErrorHC -ErrorText $errorMessage)) {
                    throw "Failed to $Description and the error will not be resolved by retrying: $errorMessage"
                }

                if ($attempt -ge $MaxRetries) {
                    throw "Failed to $Description after $MaxRetries attempts. Last error: $errorMessage"
                }

                $retryAfter = Get-RetryAfterSecondsHC -ErrorRecord $errorRecord

                $delay = if ($retryAfter -gt 0) {
                    [math]::Min($retryAfter, $maxDelaySeconds)
                }
                else {
                    [math]::Min($RetryDelaySeconds * [math]::Pow(2, $attempt - 1), $maxDelaySeconds)
                }

                $delay = [int]$delay

                Write-Warning "Attempt $attempt of $MaxRetries failed to $Description. Retrying in $delay seconds... Error: $errorMessage"

                Start-Sleep -Seconds $delay
            }
        }
    }

    function ConvertTo-GraphPathHC {
        <#
        .SYNOPSIS
            Convert a library relative folder/file path to a Graph safe path.

        .DESCRIPTION
            Splits on both slash types and drops empty segments, so 'a//b/' and
            '\a\b' both become 'a/b'. Each segment is URL encoded individually:
            encoding per segment rather than the whole string keeps the '/'
            separators intact while escaping spaces and other characters that
            are legal in SharePoint names but not in a URL.

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

    function Get-ContentTypeHC {
        <#
        .SYNOPSIS
            Map a file extension to a MIME type for the upload request.

        .DESCRIPTION
            The single-request upload has to declare a content type. Deriving it
            from the extension keeps the script honest about what it is sending:
            the parameter is FilePath, not HtmlFilePath, so it will not always be
            the overview HTML.

            Unknown extensions fall back to application/octet-stream, which
            SharePoint accepts and then classifies from the file name anyway.
        #>
        param (
            [Parameter(Mandatory)][string]$Path
        )

        switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
            '.html' { 'text/html'; break }
            '.htm' { 'text/html'; break }
            '.csv' { 'text/csv'; break }
            '.txt' { 'text/plain'; break }
            '.json' { 'application/json'; break }
            '.xml' { 'application/xml'; break }
            '.pdf' { 'application/pdf'; break }
            '.xlsx' { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; break }
            '.xlsm' { 'application/vnd.ms-excel.sheet.macroEnabled.12'; break }
            '.docx' { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'; break }
            '.pptx' { 'application/vnd.openxmlformats-officedocument.presentationml.presentation'; break }
            '.zip' { 'application/zip'; break }
            '.png' { 'image/png'; break }
            '.jpg' { 'image/jpeg'; break }
            '.jpeg' { 'image/jpeg'; break }
            default { 'application/octet-stream' }
        }
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
                    # The fallback is best effort. Whatever it failed with, the
                    # permissions message below is the more useful thing to
                    # report, so the original error is deliberately dropped.
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
            $contentType = Get-ContentTypeHC -Path $LocalFilePath

            return Invoke-WithRetryHC -MaxRetries $MaxRetries -Description "upload file to '$TargetPath'" -Action {
                $params = @{
                    Method        = 'PUT'
                    Uri           = $uploadUri
                    InputFilePath = $LocalFilePath
                    ContentType   = $contentType
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

        # Graph requires every chunk except the last to be a multiple of 320 KiB.
        # ChunkSizeMB is therefore rounded DOWN to the nearest legal boundary.
        # Sizes that divide evenly are exact (5 MB is 16 x 320 KiB); others are
        # not (1 MB becomes 0.94 MB, being 3 x 320 KiB rather than 3.2).
        $chunkGranularity = 320KB
        $chunkSize = [math]::Max(
            1, [math]::Floor(($ChunkSizeMB * 1MB) / $chunkGranularity)
        ) * $chunkGranularity

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
                    # Abandoning the session is a courtesy to SharePoint, not a
                    # requirement. The upload failure is already on its way up;
                    # a failed cleanup must not replace it.
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

        # SharePoint rejects these outright. Catching it here produces a clear
        # message naming the offending characters instead of a Graph
        # invalidRequest after the connection and site lookup have already run.
        $illegalCharacters = '"', '*', ':', '<', '>', '?', '/', '\', '|'

        $foundIllegal = @($illegalCharacters.Where({ $FileName.Contains($_) }))

        if ($foundIllegal) {
            throw "FileName '$FileName' contains characters SharePoint does not allow: $($foundIllegal -join ' ')"
        }

        if ($FileName.EndsWith('.') -or $FileName.StartsWith('.') -or $FileName.Trim() -ne $FileName) {
            throw "FileName '$FileName' cannot start or end with a period or whitespace."
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

            # The message ends up in SystemErrors and therefore in the summary
            # mail. ClientId and TenantId identify which app failed and are safe
            # to include; the credential is deliberately not echoed back.
            throw "Failed to connect to MS Graph with ClientId '$graphClientId' TenantId '$graphTenantId': $errorMessage"
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