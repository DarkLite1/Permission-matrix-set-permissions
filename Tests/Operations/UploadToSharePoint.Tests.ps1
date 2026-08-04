#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

<#
    Tests for Operations\UploadToSharePoint.ps1

    Approach:
        UploadToSharePoint.ps1 is a *script* (param + begin/process), invoked with
        the call operator. Its helpers (Get-StringValueHC, Invoke-WithRetryHC,
        ConvertTo-GraphPathHC, Get-GraphSiteIdHC, Get-GraphDriveIdHC,
        New-GraphFolderPathHC, Send-GraphFileHC) live inside begin{} so they can't
        be dot-sourced and unit-tested in isolation; instead the script is run end
        to end and its dependencies are mocked.

        Everything the script touches over the network goes through four commands,
        all of which are mocked:
            - Get-MgContext          session reuse check
            - Connect-MgGraph        certificate based app-only sign in
            - Invoke-MgGraphRequest  every Graph call (site, drives, folders,
                                     small uploads, upload sessions)
            - Invoke-RestMethod      chunk PUTs to the pre-authenticated upload
                                     URL of a resumable session, which must NOT
                                     carry a bearer token and so deliberately does
                                     not use Invoke-MgGraphRequest

        Because those calls happen in the script (not inside an imported module),
        test-scope mocks intercept them without -ModuleName.

        Start-Sleep is mocked so the retry tests don't wait. Write-Warning is
        mocked so retry warnings don't clutter the output.

        Files to upload are real files under TestDrive.

    IMPORTANT - literals inside scriptblocks:
        Pester invokes -MockWith and -ParameterFilter scriptblocks in a different
        scope from the It body. Variables defined in BeforeAll are NOT reliably
        visible there: they silently evaluate to $null, which produces mock
        responses with empty fields and filters that never match. Every value
        inside a Mock or a ParameterFilter below is therefore written out in full.
        Only two things are allowed to cross that boundary:
            - helper functions defined in BeforeAll (function lookup does walk the
              scope chain, and the sibling UpdateServiceNow tests rely on this)
            - $script: counters initialised in the It body and incremented in the
              mock, used ONLY to branch inside the mock (first call returns page
              one, second returns page two). Their value cannot be read back in
              the It body: assert call counts with Should -Invoke instead.

        The test data is deliberately short and readable ('site-guid',
        'b!documents-drive-id') precisely so those literals stay legible.

    Fixtures:
        Site URL      https://contoso.sharepoint.com/sites/IT
        Site ID       contoso.sharepoint.com,site-guid,web-guid
        Drive ID      b!documents-drive-id
        Client ID     client-id-1
        Tenant ID     tenant-id-1
        Thumbprint    THUMBPRINT1

    Routing the Graph mocks:
        Invoke-MgGraphRequest is mocked several times over, each with a
        -ParameterFilter matching one endpoint shape. The filters are mutually
        exclusive, so definition order does not matter:

            .../sites/<something>              (no /drive) GET  -> the site
            .../drives                                     GET  -> library list
            .../drive                                      GET  -> default library
            .../children                                   GET  -> folder contents
            .../children                                   POST -> create folder
            ...:/content                                   PUT  -> small upload
            .../createUploadSession                        POST -> session

    Microsoft.Graph.Authentication must be installed: the script #Requires it, and
    Pester can only mock commands that exist.

    Note on $Uri being a [System.Uri], not a string:
        Invoke-MgGraphRequest declares -Uri as [System.Uri], so PowerShell coerces
        the string at binding time and Pester records a Uri object. Two
        consequences for the filters below:
            -eq    converts the right operand to [Uri] and compares Uri to Uri.
                   Uri equality NORMALISES percent-encoding, so an assertion
                   written with -eq passes whether or not the script escaped the
                   path. Fine for asserting which endpoint was called; useless
                   for asserting that escaping happened.
            -match stringifies the left operand with ToString(), which returns the
                   UNESCAPED form. A pattern containing '%20' can therefore never
                   match, even though the request really was escaped.
        Any assertion that is specifically about percent-encoding must use
        $Uri.AbsoluteUri, which preserves it. Everything else can use $Uri.

    Note on the retry contract:
        Invoke-WithRetryHC no longer retries blindly. Three behaviours are pinned
        by the tests in the 'retries' Context:
            - errors a repeat cannot fix (accessDenied, itemNotFound, 4xx) throw
              on the first attempt with no sleep at all
            - 429 / throttling IS retried, despite being a client error
            - a Retry-After hint overrides the exponential schedule (3, 6, 12,
              capped at 60)
        Start-Sleep is mocked, so these run instantly and can be asserted on by
        the -Seconds value.

    Note on -Exactly:
        Every Should -Invoke below uses -Exactly. In Pester 5, `-Times N` without
        `-Exactly` means "at least N", so `-Times 0` ("at least 0") never fails and
        `-Times 2` would still pass if the command were called 3 times. -Exactly
        turns these into the strict counts the tests actually intend.

    Note on the drive-lookup patch:
        The Context 'no libraries are visible' block expects the improved error
        message and the /drive fallback added after the Sites.Selected incident.
        If that patch was not applied, delete that Context; the rest is unaffected.
#>

BeforeAll {
    $script:ScriptPath = "$PSScriptRoot\..\..\Scripts\Operations\UploadToSharePoint.ps1"

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Script under test not found: '$ScriptPath'. Adjust the path resolution for this test's location."
    }

    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
        throw "Module 'Microsoft.Graph.Authentication' is required to run these tests. Install it with: Install-Module Microsoft.Graph.Authentication"
    }
    Import-Module 'Microsoft.Graph.Authentication' -ErrorAction Stop

    function New-SmallFile {
        param(
            [Parameter(Mandatory)][string]$Path,
            [string]$Content = '<html><body>overview</body></html>'
        )
        Set-Content -LiteralPath $Path -Value $Content -Encoding UTF8
        $Path
    }

    # A file comfortably over the 4 MB single-request threshold, so the script
    # takes the resumable upload session branch. Zero bytes are fine: nothing
    # inspects the content.
    function New-LargeFile {
        param(
            [Parameter(Mandatory)][string]$Path,
            [int]$SizeInBytes = 4718592   # 4.5 MB
        )
        [System.IO.File]::WriteAllBytes($Path, [byte[]]::new($SizeInBytes))
        $Path
    }

    # Helper functions ARE reachable from inside a MockWith block, so these carry
    # the fixture values as literal defaults and keep the mock bodies short.
    function New-Site {
        param([string]$Id = 'contoso.sharepoint.com,site-guid,web-guid')
        [PSCustomObject]@{
            id          = $Id
            displayName = 'IT'
            webUrl      = 'https://contoso.sharepoint.com/sites/IT'
        }
    }

    function New-Drive {
        param(
            [Parameter(Mandatory)][string]$Name,
            [string]$Id = 'b!documents-drive-id'
        )
        [PSCustomObject]@{ name = $Name; id = $Id }
    }

    function New-DriveResponse {
        param(
            [object[]]$Drives = @(),
            [string]$NextLink
        )
        $response = [PSCustomObject]@{ value = $Drives }

        if ($NextLink) {
            $response | Add-Member -NotePropertyName '@odata.nextLink' -NotePropertyValue $NextLink
        }

        $response
    }

    function New-DefaultDriveResponse {
        [PSCustomObject]@{ id = 'b!documents-drive-id'; name = 'Documents' }
    }

    function New-Folder {
        param([Parameter(Mandatory)][string]$Name)
        [PSCustomObject]@{ name = $Name; folder = [PSCustomObject]@{ childCount = 0 } }
    }

    function New-FileItem {
        param([Parameter(Mandatory)][string]$Name)
        [PSCustomObject]@{ name = $Name; file = [PSCustomObject]@{ mimeType = 'text/html' } }
    }

    function New-UploadedItem {
        param([string]$Name = 'Overview.html')
        [PSCustomObject]@{
            name            = $Name
            id              = 'item-id-1'
            webUrl          = "https://contoso.sharepoint.com/sites/IT/Shared%20Documents/$Name"
            parentReference = [PSCustomObject]@{ driveId = 'b!documents-drive-id' }
        }
    }
}

Describe 'UploadToSharePoint.ps1' {
    BeforeEach {
        Remove-Item (Join-Path $TestDrive '*') -Recurse -Force -ErrorAction Ignore

        $overviewFile = New-SmallFile -Path (Join-Path $TestDrive 'Overview.html')

        $params = @{
            FilePath              = $overviewFile
            SiteUrl               = 'https://contoso.sharepoint.com/sites/IT'
            DocumentLibraryName   = 'Documents'
            ClientId              = 'client-id-1'
            TenantId              = 'tenant-id-1'
            CertificateThumbprint = 'THUMBPRINT1'
        }

        # Default: no existing Graph session, so the script connects.
        Mock Get-MgContext {}
        Mock Connect-MgGraph {}
        Mock Start-Sleep {}
        Mock Write-Warning {}

        #region Graph endpoint routing
        # The site itself: a GET whose URI has no /drive and no /children segment.
        Mock Invoke-MgGraphRequest -ParameterFilter {
            $Method -eq 'GET' -and $Uri -notmatch '/drives?(/|$)' -and $Uri -notmatch '/children$'
        } -MockWith {
            New-Site
        }

        # The document libraries of the site: one page, one matching library.
        Mock Invoke-MgGraphRequest -ParameterFilter {
            $Method -eq 'GET' -and $Uri -match '/drives$'
        } -MockWith {
            New-DriveResponse -Drives @(
                New-Drive -Name 'Documents' -Id 'b!documents-drive-id'
                New-Drive -Name 'Site Assets' -Id 'b!site-assets-drive-id'
            )
        }

        # The default document library, used only by the empty-list fallback.
        Mock Invoke-MgGraphRequest -ParameterFilter {
            $Method -eq 'GET' -and $Uri -match '/drive$'
        } -MockWith {
            New-DefaultDriveResponse
        }

        # Folder contents: empty by default, so any requested folder is created.
        Mock Invoke-MgGraphRequest -ParameterFilter {
            $Method -eq 'GET' -and $Uri -match '/children$'
        } -MockWith {
            [PSCustomObject]@{ value = @() }
        }

        Mock Invoke-MgGraphRequest -ParameterFilter {
            $Method -eq 'POST' -and $Uri -match '/children$'
        } -MockWith {
            [PSCustomObject]@{ id = 'new-folder-id' }
        }

        Mock Invoke-MgGraphRequest -ParameterFilter {
            $Method -eq 'PUT' -and $Uri -match ':/content$'
        } -MockWith {
            New-UploadedItem
        }

        Mock Invoke-MgGraphRequest -ParameterFilter {
            $Method -eq 'POST' -and $Uri -match '/createUploadSession$'
        } -MockWith {
            [PSCustomObject]@{ uploadUrl = 'https://upload.contoso.example/session/abc123' }
        }

        # Chunk PUTs of a resumable session.
        Mock Invoke-RestMethod { New-UploadedItem }
        #endregion
    }

    AfterEach {
        Remove-Item Env:\SP_TEST_CLIENT_ID -ErrorAction Ignore
        Remove-Item Env:\SP_TEST_TENANT_ID -ErrorAction Ignore
        Remove-Item Env:\SP_TEST_THUMBPRINT -ErrorAction Ignore
        Remove-Item Env:\SP_MISSING_VAR -ErrorAction Ignore
    }

    Context 'parameter validation' {
        It 'rejects a MaxRetries below 1' {
            $params.MaxRetries = 0

            { & $ScriptPath @params } | Should -Throw

            Should -Invoke Connect-MgGraph -Exactly -Times 0
        }

        It 'rejects a ChunkSizeMB of <ChunkSizeMB>' -TestCases @(
            @{ ChunkSizeMB = 0 }
            @{ ChunkSizeMB = 61 }
        ) {
            param($ChunkSizeMB)

            $params.ChunkSizeMB = $ChunkSizeMB

            { & $ScriptPath @params } | Should -Throw

            Should -Invoke Connect-MgGraph -Exactly -Times 0
        }

        It 'throws when the file to upload does not exist' {
            $params.FilePath = Join-Path $TestDrive 'no-such-file.html'

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*not found*'

            Should -Invoke Connect-MgGraph -Exactly -Times 0
            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0
        }

        It 'throws when the file to upload is a folder rather than a file' {
            $folder = Join-Path $TestDrive 'a-folder'
            $null = New-Item -Path $folder -ItemType Directory
            $params.FilePath = $folder

            { & $ScriptPath @params } | Should -Throw

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0
        }
    }

    Context 'credential resolution' {
        It 'passes a plain credential value straight through' {
            & $ScriptPath @params

            Should -Invoke Connect-MgGraph -Exactly -Times 1 -ParameterFilter {
                $ClientId -eq 'client-id-1' -and
                $TenantId -eq 'tenant-id-1' -and
                $CertificateThumbprint -eq 'THUMBPRINT1'
            }
        }

        It 'resolves an ENV: value from the environment' {
            $env:SP_TEST_CLIENT_ID = 'client-from-env'
            $env:SP_TEST_TENANT_ID = 'tenant-from-env'
            $env:SP_TEST_THUMBPRINT = 'thumbprint-from-env'

            $params.ClientId = 'ENV:SP_TEST_CLIENT_ID'
            $params.TenantId = 'ENV:SP_TEST_TENANT_ID'
            $params.CertificateThumbprint = 'ENV:SP_TEST_THUMBPRINT'

            & $ScriptPath @params

            Should -Invoke Connect-MgGraph -Exactly -Times 1 -ParameterFilter {
                $ClientId -eq 'client-from-env' -and
                $TenantId -eq 'tenant-from-env' -and
                $CertificateThumbprint -eq 'thumbprint-from-env'
            }
        }

        It 'resolves ENV: case insensitively' {
            $env:SP_TEST_CLIENT_ID = 'client-from-env'
            $params.ClientId = 'env:SP_TEST_CLIENT_ID'

            & $ScriptPath @params

            Should -Invoke Connect-MgGraph -Exactly -Times 1 -ParameterFilter {
                $ClientId -eq 'client-from-env'
            }
        }

        It 'throws when an ENV: value points to a missing variable' {
            $params.ClientId = 'ENV:SP_MISSING_VAR'

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage "*Environment variable 'SP_MISSING_VAR' not found*"

            Should -Invoke Connect-MgGraph -Exactly -Times 0
        }
    }

    Context 'connecting to MS Graph' {
        It 'connects when there is no existing session' {
            & $ScriptPath @params

            Should -Invoke Connect-MgGraph -Exactly -Times 1
        }

        It 'reuses an existing session for the same app and tenant' {
            Mock Get-MgContext {
                [PSCustomObject]@{ ClientId = 'client-id-1'; TenantId = 'tenant-id-1' }
            }

            & $ScriptPath @params

            Should -Invoke Connect-MgGraph -Exactly -Times 0
            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            }
        }

        It 'reconnects when the existing session belongs to a different app' {
            Mock Get-MgContext {
                [PSCustomObject]@{ ClientId = 'another-client'; TenantId = 'tenant-id-1' }
            }

            & $ScriptPath @params

            Should -Invoke Connect-MgGraph -Exactly -Times 1
        }

        It 'reconnects when the existing session belongs to a different tenant' {
            Mock Get-MgContext {
                [PSCustomObject]@{ ClientId = 'client-id-1'; TenantId = 'another-tenant' }
            }

            & $ScriptPath @params

            Should -Invoke Connect-MgGraph -Exactly -Times 1
        }

        It 'connects without a welcome banner' {
            & $ScriptPath @params

            Should -Invoke Connect-MgGraph -Exactly -Times 1 -ParameterFilter {
                $NoWelcome -eq $true
            }
        }

        It 'throws a clear error when the connection fails' {
            Mock Connect-MgGraph { throw 'certificate not found' }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*Failed to connect to MS Graph*'

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0
        }
    }

    Context 'resolving the site' {
        It 'addresses a team site by hostname and server relative path' {
            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Uri -eq 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/IT'
            }
        }

        It 'addresses a root site by hostname alone' {
            $params.SiteUrl = 'https://contoso.sharepoint.com'

            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Uri -eq 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com'
            }
        }

        It 'ignores a trailing slash on the site URL' {
            $params.SiteUrl = 'https://contoso.sharepoint.com/sites/IT/'

            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Uri -eq 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/IT'
            }
        }

        It 'throws when the site cannot be resolved' {
            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -notmatch '/drives?(/|$)' -and $Uri -notmatch '/children$'
            } -MockWith {
                [PSCustomObject]@{ }
            }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*not found*'
        }

        It 'throws when the site URL is not a valid URL' {
            # A scheme-less value constructs as a relative Uri without throwing;
            # it is the AbsolutePath call that fails. Both are inside the same
            # try, so the friendly message must come back rather than a raw
            # InvalidOperationException.
            $params.SiteUrl = 'contoso.sharepoint.com/sites/IT'   # no scheme

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*not a valid URL*'
        }
    }

    Context 'finding the document library' {
        It 'matches the library by its display name' {
            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Uri -eq 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,site-guid,web-guid/drives'
            }
            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and
                $Uri -eq 'https://graph.microsoft.com/v1.0/drives/b!documents-drive-id/root:/Overview.html:/content'
            }
        }

        It "accepts 'Shared Documents' when the library is called 'Documents'" {
            $params.DocumentLibraryName = 'Shared Documents'

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -match '/drives$'
            } -MockWith {
                New-DriveResponse -Drives @(New-Drive -Name 'Documents' -Id 'b!documents-drive-id')
            }

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match '/drives/b!documents-drive-id/'
            }
        }

        It "accepts 'Documents' when the library is called 'Shared Documents'" {
            $params.DocumentLibraryName = 'Documents'

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -match '/drives$'
            } -MockWith {
                New-DriveResponse -Drives @(New-Drive -Name 'Shared Documents' -Id 'b!documents-drive-id')
            }

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match '/drives/b!documents-drive-id/'
            }
        }

        It 'follows @odata.nextLink until every page is read' {
            $script:driveCallCount = 0

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -match '/drives'
            } -MockWith {
                $script:driveCallCount = ([int]$script:driveCallCount) + 1

                if ($script:driveCallCount -eq 1) {
                    New-DriveResponse -Drives @(New-Drive -Name 'Site Assets' -Id 'b!site-assets-drive-id') -NextLink 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,site-guid,web-guid/drives?page=2'
                }
                else {
                    New-DriveResponse -Drives @(New-Drive -Name 'Documents' -Id 'b!documents-drive-id')
                }
            }

            { & $ScriptPath @params } | Should -Not -Throw

            # The counter drives the branching inside the mock only; it cannot be
            # read back here, so the call count is asserted through Pester.
            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 2 -ParameterFilter {
                $Method -eq 'GET' -and $Uri -match '/drives'
            }
        }

        It 'throws and lists what is available when the name does not match' {
            $params.DocumentLibraryName = 'Reports Library'

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -match '/drives$'
            } -MockWith {
                New-DriveResponse -Drives @(
                    New-Drive -Name 'Site Assets' -Id 'b!site-assets-drive-id'
                    New-Drive -Name 'Style Library' -Id 'b!style-library-drive-id'
                )
            }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*Site Assets*'
        }
    }

    Context 'no libraries are visible (drive-lookup patch)' {
        BeforeEach {
            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -match '/drives$'
            } -MockWith {
                New-DriveResponse -Drives @()
            }
        }

        It 'falls back to the default library when the list comes back empty' {
            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'GET' -and
                $Uri -eq 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,site-guid,web-guid/drive'
            }
        }

        It 'points at permissions rather than the library name when the fallback also fails' {
            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -match '/drive$'
            } -MockWith {
                throw 'accessDenied'
            }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*Sites.Selected*'
        }
    }

    Context 'creating the folder structure' {
        It 'does not touch folders when no FolderPath is given' {
            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0 -ParameterFilter {
                $Method -eq 'POST' -and $Uri -match '/children$'
            }
        }

        It 'creates every missing segment of the path' {
            $params.FolderPath = 'Reports/Permission matrix'

            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 2 -ParameterFilter {
                $Method -eq 'POST' -and $Uri -match '/children$'
            }
        }

        It 'skips segments that already exist' {
            $params.FolderPath = 'Reports'

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -match '/children$'
            } -MockWith {
                [PSCustomObject]@{ value = @(New-Folder -Name 'Reports') }
            }

            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0 -ParameterFilter {
                $Method -eq 'POST' -and $Uri -match '/children$'
            }
        }

        It 'creates the folder when a file of the same name exists' {
            # A file named 'Reports' must not be mistaken for the folder.
            $params.FolderPath = 'Reports'

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -match '/children$'
            } -MockWith {
                [PSCustomObject]@{ value = @(New-FileItem -Name 'Reports') }
            }

            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'POST' -and $Uri -match '/children$'
            }
        }

        It 'tolerates a folder created concurrently by another run' {
            $params.FolderPath = 'Reports'

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'POST' -and $Uri -match '/children$'
            } -MockWith {
                throw 'nameAlreadyExists'
            }

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            }
        }

        It 'rethrows a folder creation error that is not a name collision' {
            $params.FolderPath = 'Reports'

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'POST' -and $Uri -match '/children$'
            } -MockWith {
                throw 'quotaLimitReached'
            }

            { & $ScriptPath @params } | Should -Throw
        }

        It 'normalises mixed and duplicated separators' {
            $params.FolderPath = '\Reports//Permission matrix\'

            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 2 -ParameterFilter {
                $Method -eq 'POST' -and $Uri -match '/children$'
            }
        }
    }

    Context 'uploading a small file' {
        It 'sends a single PUT to the content endpoint' {
            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and
                $Uri -eq 'https://graph.microsoft.com/v1.0/drives/b!documents-drive-id/root:/Overview.html:/content' -and
                $InputFilePath -match 'Overview\.html$'
            }

            Should -Invoke Invoke-RestMethod -Exactly -Times 0
        }

        It 'defaults the target name to the local file name' {
            $params.FilePath = New-SmallFile -Path (Join-Path $TestDrive 'Something else.html')

            & $ScriptPath @params

            # $Uri is bound as [System.Uri], and its ToString() returns the
            # UNESCAPED form, so '%20' would never match. AbsoluteUri keeps the
            # escaping, which is the whole point of this assertion.
            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri.AbsoluteUri -match 'Something%20else\.html'
            }
        }

        It 'uses FileName when one is given' {
            $params.FileName = 'Permission matrix overview.html'

            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri.AbsoluteUri -match 'Permission%20matrix%20overview\.html'
            }
        }

        It 'URL encodes each path segment but keeps the separators' {
            $params.FolderPath = 'Reports/Permission matrix'

            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and
                $Uri -eq 'https://graph.microsoft.com/v1.0/drives/b!documents-drive-id/root:/Reports/Permission%20matrix/Overview.html:/content'
            }
        }

        It 'returns the uploaded item details' {
            $result = & $ScriptPath @params

            $result.Name | Should -BeExactly 'Overview.html'
            $result.Id | Should -BeExactly 'item-id-1'
            $result.DriveId | Should -BeExactly 'b!documents-drive-id'
            $result.SiteId | Should -BeExactly 'contoso.sharepoint.com,site-guid,web-guid'
            $result.WebUrl | Should -Not -BeNullOrEmpty
        }

        It 'overwrites without complaint when the file is already there' {
            # Overwriting is the default behaviour of PUT to /content, so the
            # script must not branch on existence: no DELETE before uploading.
            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0 -ParameterFilter {
                $Method -eq 'DELETE'
            }
        }
    }

    Context 'uploading a large file' {
        BeforeEach {
            $params.FilePath = New-LargeFile -Path (Join-Path $TestDrive 'Big.html')
            $params.ChunkSizeMB = 1
        }

        It 'creates a resumable upload session that replaces on conflict' {
            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'POST' -and
                $Uri -eq 'https://graph.microsoft.com/v1.0/drives/b!documents-drive-id/root:/Big.html:/createUploadSession' -and
                $Body -match 'replace'
            }
        }

        It 'does not use the single request endpoint' {
            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            }
        }

        It 'sends the chunks to the pre-authenticated upload URL' {
            & $ScriptPath @params

            # 4.5 MB (4718592 bytes) in chunks of 983040 bytes (1 * 320 KiB * 3)
            # is 5 requests: four full chunks and a short final one.
            Should -Invoke Invoke-RestMethod -Exactly -Times 5 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -eq 'https://upload.contoso.example/session/abc123'
            }
        }

        It 'sends a Content-Range header describing each chunk' {
            & $ScriptPath @params

            Should -Invoke Invoke-RestMethod -Exactly -Times 1 -ParameterFilter {
                $Headers['Content-Range'] -eq 'bytes 0-983039/4718592'
            }
            Should -Invoke Invoke-RestMethod -Exactly -Times 1 -ParameterFilter {
                $Headers['Content-Range'] -eq 'bytes 3932160-4718591/4718592'
            }
        }

        It 'throws when the session has no upload URL' {
            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'POST' -and $Uri -match '/createUploadSession$'
            } -MockWith {
                [PSCustomObject]@{ }
            }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*no upload URL*'

            Should -Invoke Invoke-RestMethod -Exactly -Times 0
        }

        It 'abandons the session when a chunk fails, so no lock is left behind' {
            Mock Invoke-RestMethod -ParameterFilter { $Method -eq 'PUT' } -MockWith {
                throw 'network error'
            }

            { & $ScriptPath @params } | Should -Throw

            # The chunk is retried MaxRetries (default 3) times before the
            # session is abandoned exactly once.
            Should -Invoke Invoke-RestMethod -Exactly -Times 3 -ParameterFilter {
                $Method -eq 'PUT'
            }
            Should -Invoke Invoke-RestMethod -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'DELETE' -and $Uri -eq 'https://upload.contoso.example/session/abc123'
            }
        }
    }

    Context 'retries' {
        It 'retries a failing call and succeeds on a later attempt' {
            $script:uploadAttempt = 0

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            } -MockWith {
                $script:uploadAttempt = ([int]$script:uploadAttempt) + 1

                if ($script:uploadAttempt -lt 3) { throw 'transient 503' }

                New-UploadedItem
            }

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 3 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            }
            Should -Invoke Start-Sleep -Exactly -Times 2
            Should -Invoke Write-Warning -Exactly -Times 2
        }

        It 'gives up after MaxRetries attempts' {
            $params.MaxRetries = 3

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            } -MockWith {
                throw 'permanent 500'
            }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*after 3 attempts*'

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 3 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            }
            # Three attempts means two pauses: never sleep after the last one.
            Should -Invoke Start-Sleep -Exactly -Times 2
        }

        It 'does not retry when MaxRetries is 1' {
            $params.MaxRetries = 1

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            } -MockWith {
                throw 'permanent 500'
            }

            { & $ScriptPath @params } | Should -Throw

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            }
            Should -Invoke Start-Sleep -Exactly -Times 0
        }

        It 'retries the site lookup as well as the upload' {
            $params.MaxRetries = 2
            $script:siteAttempt = 0

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'GET' -and $Uri -notmatch '/drives?(/|$)' -and $Uri -notmatch '/children$'
            } -MockWith {
                $script:siteAttempt = ([int]$script:siteAttempt) + 1

                if ($script:siteAttempt -lt 2) { throw 'transient 503' }

                New-Site
            }

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 2 -ParameterFilter {
                $Method -eq 'GET' -and $Uri -notmatch '/drives?(/|$)' -and $Uri -notmatch '/children$'
            }
            Should -Invoke Start-Sleep -Exactly -Times 1
        }

        It 'does not retry <Error>, which a repeat cannot fix' -TestCases @(
            @{ Error = 'accessDenied' }
            @{ Error = 'itemNotFound' }
            @{ Error = 'unauthenticated' }
            @{ Error = 'invalidRequest' }
            @{ Error = 'HTTP 403 Forbidden' }
        ) {
            param($Error)

            # A missing Sites.Selected grant used to cost three attempts and two
            # backoff waits before reporting an error that was never going to
            # clear. It now fails on the first attempt.
            $params.MaxRetries = 3

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            } -MockWith {
                throw $Error
            }

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*will not be resolved by retrying*'

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            }
            Should -Invoke Start-Sleep -Exactly -Times 0
        }

        It 'still retries throttling, even though 429 is a client error' {
            $params.MaxRetries = 3
            $script:throttleAttempt = 0

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            } -MockWith {
                $script:throttleAttempt = ([int]$script:throttleAttempt) + 1

                if ($script:throttleAttempt -lt 2) { throw 'HTTP 429 activityLimitReached' }

                New-UploadedItem
            }

            { & $ScriptPath @params } | Should -Not -Throw

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 2 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            }
        }

        It 'honours a Retry-After hint instead of its own backoff' {
            $params.MaxRetries = 3
            $script:retryAfterAttempt = 0

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            } -MockWith {
                $script:retryAfterAttempt = ([int]$script:retryAfterAttempt) + 1

                if ($script:retryAfterAttempt -lt 2) {
                    throw 'HTTP 429 TooManyRequests Retry-After: 17'
                }

                New-UploadedItem
            }

            & $ScriptPath @params

            # Retrying sooner than Graph asked deepens the throttling, so the
            # hint wins over the exponential schedule (which would say 3).
            Should -Invoke Start-Sleep -Exactly -Times 1 -ParameterFilter {
                $Seconds -eq 17
            }
        }

        It 'backs off exponentially when there is no Retry-After hint' {
            $params.MaxRetries = 4

            Mock Invoke-MgGraphRequest -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            } -MockWith {
                throw 'transient 503'
            }

            { & $ScriptPath @params } | Should -Throw

            # Four attempts, three waits: 3, 6, 12.
            Should -Invoke Start-Sleep -Exactly -Times 1 -ParameterFilter { $Seconds -eq 3 }
            Should -Invoke Start-Sleep -Exactly -Times 1 -ParameterFilter { $Seconds -eq 6 }
            Should -Invoke Start-Sleep -Exactly -Times 1 -ParameterFilter { $Seconds -eq 12 }
        }
    }

    Context 'target file name validation' {
        It 'rejects a FileName containing <Character>, which SharePoint forbids' -TestCases @(
            @{ Character = '*'; FileName = 'Over*view.html' }
            @{ Character = ':'; FileName = 'Over:view.html' }
            @{ Character = '<'; FileName = 'Over<view.html' }
            @{ Character = '|'; FileName = 'Over|view.html' }
            @{ Character = '?'; FileName = 'Over?view.html' }
        ) {
            param($FileName)

            # Caught before connecting, so the message names the character
            # instead of arriving as a Graph invalidRequest after the site and
            # library have already been resolved.
            $params.FileName = $FileName

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*SharePoint does not allow*'

            Should -Invoke Connect-MgGraph -Exactly -Times 0
        }

        It 'rejects a FileName that ends with a period' {
            $params.FileName = 'Overview.html.'

            { & $ScriptPath @params } |
                Should -Throw -ExpectedMessage '*period or whitespace*'
        }

        It 'accepts a FileName with spaces and other legal punctuation' {
            $params.FileName = 'Permission matrix overview (2026-07).html'

            { & $ScriptPath @params } | Should -Not -Throw
        }
    }

    Context 'content type' {
        It 'declares <Expected> when uploading a <Extension> file' -TestCases @(
            @{ Extension = '.html'; Expected = 'text/html' }
            @{ Extension = '.csv'; Expected = 'text/csv' }
            @{ Extension = '.json'; Expected = 'application/json' }
            @{ Extension = '.bin'; Expected = 'application/octet-stream' }
        ) {
            param($Extension, $Expected)

            # The parameter is FilePath, not HtmlFilePath: hardcoding text/html
            # mislabels anything else this gets pointed at.
            $params.FilePath = New-SmallFile -Path (Join-Path $TestDrive "Report$Extension")

            & $ScriptPath @params

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and $ContentType -eq $Expected
            }
        }
    }

    Context 'happy path' {
        It 'connects, resolves, creates the folder and uploads exactly once each' {
            $params.FolderPath = 'Reports'

            $result = & $ScriptPath @params

            Should -Invoke Connect-MgGraph -Exactly -Times 1
            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Uri -eq 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/IT'
            }
            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Uri -eq 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,site-guid,web-guid/drives'
            }
            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'POST' -and $Uri -match '/children$'
            }
            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
                $Method -eq 'PUT' -and
                $Uri -eq 'https://graph.microsoft.com/v1.0/drives/b!documents-drive-id/root:/Reports/Overview.html:/content'
            }
            Should -Invoke Start-Sleep -Exactly -Times 0
            Should -Invoke Write-Warning -Exactly -Times 0

            $result.Name | Should -BeExactly 'Overview.html'
        }

        It 'is repeatable: a second run behaves identically' {
            & $ScriptPath @params
            $second = & $ScriptPath @params

            $second.Name | Should -BeExactly 'Overview.html'

            Should -Invoke Invoke-MgGraphRequest -Exactly -Times 2 -ParameterFilter {
                $Method -eq 'PUT' -and $Uri -match ':/content$'
            }
        }
    }
}