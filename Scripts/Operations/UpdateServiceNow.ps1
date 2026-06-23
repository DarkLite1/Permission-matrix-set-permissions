#Requires -Version 7
#Requires -Modules ImportExcel, ServiceNow

<#
    .SYNOPSIS
        Synchronizes Permission Matrix form data with a specified ServiceNow
        table.

    .DESCRIPTION
        This script connects to the ServiceNow API to populate the catalog
        tables used by end-users to request access to specific folders.

        It performs a "differential sync" (non-destructive): instead of dropping
        and re-uploading the table, it compares the records produced by the
        Permission Matrix pipeline against the records currently in the target
        ServiceNow table and reconciles the differences:

            - A source row that is NOT present in the table is created with
              u_active = TRUE.
            - A table row that is NOT present in the source collection is marked
              u_active = FALSE (rows already inactive are left untouched).
            - A source row that matches an existing table row which was
              previously deactivated is reactivated (u_active = TRUE).

        Rows are matched by comparing ALL source fields (every column in the
        worksheet) except the managed u_active flag and ServiceNow's own system
        fields (sys_*). Because matching is on the full field set, a record that
        changed in any field no longer matches its old row: the old row is
        deactivated and the changed record is added as a new, active row.

        Security Feature: The JSON credentials file supports dynamically
        fetching secrets from the host's environment variables to avoid
        hardcoding passwords. Simply prefix the JSON value with 'ENV:'
        (e.g., "Password": "ENV:SNOW_SECURE_PASS").

    .PARAMETER CredentialsFilePath
        The absolute path to the .JSON file containing the ServiceNow logon
        credentials.

    .PARAMETER Environment
        The specific ServiceNow environment/node to target (e.g., 'Test',
        'Prod'). This key must exist inside the Credentials JSON file.

    .PARAMETER FormDataExcelFilePath
        The absolute path to the compiled Excel file containing the structured
        records to be uploaded.

    .PARAMETER ExcelFileWorksheetName
        The specific worksheet inside the FormDataExcelFilePath to read the
        records from (usually 'FormData').

    .PARAMETER TableName
        The exact physical name of the target table in ServiceNow
        (e.g., 'u_permission_matrix_roles').

    .PARAMETER MaxRetries
        The maximum number of retry attempts for network/API timeouts when
        creating, reactivating or deactivating records. (Default: 3, minimum 1)

    .PARAMETER ChunkSize
        The number of existing records to fetch per chunk when reading the target
        table for comparison. The table is read up front in chunks of this size
        using -First/-Skip paging. Each chunk is retried independently, so a
        transient network failure only costs a single chunk. (Default: 200)

    .EXAMPLE
        .\Update-ServiceNow.ps1 `
            -CredentialsFilePath 'C:\Secrets\ServiceNow.json' `
            -Environment 'Prod' `
            -FormDataExcelFilePath 'C:\Matrix\Exports\FormData.xlsx' `
            -ExcelFileWorksheetName 'FormData' `
            -TableName 'u_custom_matrix_roles' `
            -MaxRetries 5
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
    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxRetries = 3,
    [ValidateRange(1, 10000)]
    [int]$ChunkSize = 200
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

    function ConvertTo-ComparableStringHC {
        <#
        .SYNOPSIS
            Normalise a value to the single canonical string used BOTH for
            comparison and for upload, so the two can never disagree.

        .DESCRIPTION
            Import-Excel infers cell types, so an identifier such as the AD name
            '158774' arrives as a number, not text. If that number is uploaded as
            a JSON number, ServiceNow stores it as '158774.0', which then never
            matches the '158774' this script computes when comparing - so the row
            is deactivated and recreated on every single run.

            Coercing to a trimmed string here, and dropping the fractional part of
            whole numbers, makes the value we STORE and the value we COMPARE
            identical: '158774' in both places. NULL becomes an empty string.
        #>
        param (
            $Value
        )

        if ($null -eq $Value) { return '' }

        if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
            $number = [double]$Value

            # Whole number (e.g. an AD name read as 158774.0) -> '158774', never
            # '158774.0'. The range guard keeps the [long] cast safe.
            if (
                [double]::IsFinite($number) -and
                [math]::Truncate($number) -eq $number -and
                [math]::Abs($number) -lt 9e18
            ) {
                return ([long]$number).ToString([System.Globalization.CultureInfo]::InvariantCulture)
            }

            return $number.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        }

        return ([string]$Value).Trim()
    }

    function Get-RecordComparisonKeyHC {
        <#
        .SYNOPSIS
            Build a stable comparison key for a record from a fixed field list.

        .DESCRIPTION
            Concatenates "<field>=<value>" pairs (in the order the fields are
            supplied) into a single string. Each value is run through
            ConvertTo-ComparableStringHC (NULL -> '', whitespace trimmed, whole
            numbers rendered without a trailing '.0'), the same normalisation used
            when building the upload payload so stored and compared values agree.
            The field name is included so that, e.g., a value of 'a' in one field
            is never confused with the same value in another field. The pieces are
            joined with the Unit Separator control character, which will not occur
            in normal field data.

            Comparison is case-sensitive and works on the raw (flat) field value.
            Pass the field list pre-sorted so both sides of the comparison build
            their key in the same order.

        .PARAMETER Record
            A PSCustomObject (a table row) or a Hashtable (a source row). Both
            support '$Record.$fieldName' access.

        .PARAMETER Field
            The ordered list of field names to include in the key.
        #>
        param (
            [Parameter(Mandatory)]$Record,
            [Parameter(Mandatory)][string[]]$Field
        )

        $separator = [char]0x1F  # Unit Separator - absent from real field data

        $parts = foreach ($name in $Field) {
            "$name=$(ConvertTo-ComparableStringHC -Value $Record.$name)"
        }

        $parts -join $separator
    }

    function Test-RecordActiveHC {
        <#
        .SYNOPSIS
            Interpret a u_active value as a boolean.

        .DESCRIPTION
            ServiceNow may return the flag as an actual boolean or as a string
            ('true'/'false', '1'/'0', 'yes'/'no') depending on the table and
            session configuration. This normalises all of those to $true/$false.
            A missing/NULL value is treated as inactive.
        #>
        param (
            $Value
        )

        if ($null -eq $Value) { return $false }
        if ($Value -is [bool]) { return $Value }

        $normalized = ([string]$Value).Trim().ToLowerInvariant()

        return (@('true', '1', 'yes', 'y') -contains $normalized)
    }

    function Invoke-WithRetryHC {
        <#
        .SYNOPSIS
            Run a ServiceNow operation with bounded retries.

        .DESCRIPTION
            Executes the supplied script block, retrying on failure up to
            MaxRetries attempts with a 3 second pause *between* attempts (never
            after the final attempt).

            - With -Critical, exhausting the retries rethrows a 'CRITICAL FAILURE'
              error (aborting the script). Used for record creation, where a
              missing record is not acceptable.
            - Without -Critical, exhausting the retries writes a warning and
              returns $false so the caller can continue. Used for (de)activation,
              which is best-effort and self-heals on the next run.

            Returns $true on success, $false on a non-critical exhaustion.
        #>
        param (
            [Parameter(Mandatory)][scriptblock]$Action,
            [Parameter(Mandatory)][string]$Description,
            [Parameter(Mandatory)][int]$MaxRetries,
            [switch]$Critical
        )

        $attempt = 0

        while ($attempt -lt $MaxRetries) {
            $attempt++

            try {
                & $Action
                return $true
            }
            catch {
                if ($attempt -ge $MaxRetries) {
                    if ($Critical) {
                        throw "CRITICAL FAILURE: Could not $Description after $MaxRetries attempts. Last error: $_"
                    }

                    Write-Warning "Failed to $Description after $MaxRetries attempts; continuing. Last error: $_"
                    return $false
                }

                Write-Warning "Attempt $attempt of $MaxRetries failed to $Description. Retrying in 3 seconds... Error: $_"

                Start-Sleep -Seconds 3
            }
        }
    }

    function Get-TableRecordsInChunksHC {
        <#
        .SYNOPSIS
            Read an entire ServiceNow table up front, in retryable chunks.

        .DESCRIPTION
            Reads every row of the table before any comparison or write happens,
            so a flaky network can only spoil the read phase (which is safe to
            retry) and never leaves the table half-reconciled.

            The table is paged with -First/-Skip: each chunk requests the next
            $ChunkSize rows at an increasing -Skip offset. The read does not mutate
            the table, so the offset stays stable across chunks (rows do not shift
            underneath it).

            Each chunk is fetched independently and retried up to MaxRetries with a
            3 second pause between attempts, so a transient failure costs at most
            one chunk. A chunk that still cannot be read after MaxRetries throws -
            an incomplete read must not be treated as "these rows are gone".

            Returns every row (an empty sequence when the table is empty).

        .PARAMETER Table
            The physical ServiceNow table name to read.

        .PARAMETER ChunkSize
            The maximum number of rows to request per chunk.

        .PARAMETER MaxRetries
            Maximum attempts per chunk before giving up and throwing.
        #>
        param (
            [Parameter(Mandatory)][string]$Table,
            [Parameter(Mandatory)][int]$ChunkSize,
            [Parameter(Mandatory)][int]$MaxRetries
        )

        $allRecords = [System.Collections.Generic.List[object]]::new()
        $skip = 0
        $chunkNumber = 0

        while ($true) {
            $chunkNumber++

            $attempt = 0
            $chunk = $null

            while ($true) {
                $attempt++

                try {
                    $chunk = @(
                        Get-ServiceNowRecord -Table $Table -First $ChunkSize -Skip $skip -Verbose:$false
                    )
                    break
                }
                catch {
                    if ($attempt -ge $MaxRetries) {
                        throw "Failed to retrieve records (chunk $chunkNumber) from ServiceNow table '$Table' after $MaxRetries attempts: $_"
                    }

                    Write-Warning "Attempt $attempt of $MaxRetries failed to read chunk $chunkNumber from '$Table'. Retrying in 3 seconds... Error: $_"

                    Start-Sleep -Seconds 3
                }
            }

            if ($chunk.Count -eq 0) {
                break
            }

            $allRecords.AddRange($chunk)

            # A short chunk means we have reached the end of the table.
            if ($chunk.Count -lt $ChunkSize) {
                break
            }

            $skip += $chunk.Count
        }

        $allRecords
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
        $recordsToUpload = @(Import-Excel @params)
    }
    catch {
        throw "Failed to import records to upload from file '$FormDataExcelFilePath' with worksheet name '$ExcelFileWorksheetName': $_"
    }
    #endregion

    if ($recordsToUpload) {
        #region Determine the fields to compare on
        # "All fields" means every column in the source data, except the active
        # flag we manage and ServiceNow's own system fields (sys_*). Sorted so the
        # comparison key is built in the same order for the source and the table.
        $comparisonField = @(
            $recordsToUpload[0].PSObject.Properties.Name | Where-Object {
                $_ -ne 'u_active' -and $_ -notlike 'sys_*'
            }
        ) | Sort-Object

        if (-not $comparisonField) {
            throw "No comparable fields found in '$FormDataExcelFilePath' worksheet '$ExcelFileWorksheetName'."
        }
        #endregion

        # Values written to the custom 'u_active' field. If your table stores this
        # as a string/choice rather than a true/false boolean, change these to
        # 'true' / 'false'.
        $activeTrueValue = $true
        $activeFalseValue = $false

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

        #region Read all existing records from the ServiceNow table (chunked)
        # The whole table (active AND inactive rows) is read up front, in chunks,
        # so a flaky network only affects the read phase - which is retried and
        # safe to retry - rather than leaving the table half-reconciled. Paging is
        # by -First/-Skip; the read does not mutate the table, so the offset stays
        # stable. See Get-TableRecordsInChunksHC for details.
        Write-Verbose "Read existing records from ServiceNow table '$TableName' in chunks of $ChunkSize"

        $readParams = @{
            Table      = $TableName
            ChunkSize  = $ChunkSize
            MaxRetries = $MaxRetries
        }
        $existingRecords = @(Get-TableRecordsInChunksHC @readParams)

        Write-Verbose "Read $($existingRecords.Count) existing record(s) from '$TableName'"
        #endregion

        #region Index existing records by their comparison key
        # key -> list of table rows. A list (not a single row) because the table
        # can, in principle, hold more than one row with the same field values.
        $existingByKey = @{}

        foreach ($existingRecord in $existingRecords) {
            $key = Get-RecordComparisonKeyHC -Record $existingRecord -Field $comparisonField

            if (-not $existingByKey.ContainsKey($key)) {
                $existingByKey[$key] = [System.Collections.Generic.List[object]]::new()
            }

            $existingByKey[$key].Add($existingRecord)
        }
        #endregion

        $matchedSysId = [System.Collections.Generic.HashSet[string]]::new()

        $createParams = @{ Table = $TableName; Verbose = $false }
        $updateParams = @{ Table = $TableName; Verbose = $false }

        $createdCount = 0
        $reactivatedCount = 0
        $deactivatedCount = 0

        #region Add new rows and reactivate matched-but-inactive rows
        # Build each upload payload as a hashtable of STRING values, using the same
        # normalisation as the comparison key (ConvertTo-ComparableStringHC). This
        # is what keeps the value we store and the value we compare identical: a
        # numeric-looking identifier such as an AD name '158774' is sent as the
        # string "158774" and stored verbatim, rather than going out as the JSON
        # number 158774.0 (which ServiceNow stores as "158774.0" and which then
        # never matches on the next run). u_active is managed below; system fields
        # are never uploaded.
        $recordsToUploadHashTable = foreach ($sourceRecord in $recordsToUpload) {
            $payload = @{}

            foreach ($property in $sourceRecord.PSObject.Properties) {
                if ($property.Name -eq 'u_active' -or $property.Name -like 'sys_*') {
                    continue
                }

                $payload[$property.Name] = ConvertTo-ComparableStringHC -Value $property.Value
            }

            $payload
        }

        $currentRecordCount = 0
        $totalRecordCount = @($recordsToUploadHashTable).Count

        foreach ($record in $recordsToUploadHashTable) {
            $currentRecordCount++

            $key = Get-RecordComparisonKeyHC -Record $record -Field $comparisonField

            if ($existingByKey.ContainsKey($key)) {
                #region Already present - keep it, reactivate if it was switched off
                foreach ($match in $existingByKey[$key]) {
                    [void]$matchedSysId.Add([string]$match.sys_id)

                    if (-not (Test-RecordActiveHC $match.u_active)) {
                        $sysId = [string]$match.sys_id
                        $description = "reactivate record '$sysId' for matrix '$($record.u_matrixfilename)' AD object '$($record.u_adobjectname)'"

                        Write-Verbose "($currentRecordCount/$totalRecordCount) $description"

                        $reactivated = Invoke-WithRetryHC -MaxRetries $MaxRetries -Description $description -Action {
                            Update-ServiceNowRecord @updateParams -ID $sysId -Values @{ u_active = $activeTrueValue }
                        }

                        if ($reactivated) { $reactivatedCount++ }
                    }
                }
                #endregion
            }
            else {
                #region Not present - create it, switched on
                $record['u_active'] = $activeTrueValue
                $description = "create record for matrix '$($record.u_matrixfilename)' AD object '$($record.u_adobjectname)'"

                Write-Verbose "($currentRecordCount/$totalRecordCount) $description"

                # Creation is critical: an exhausted retry rethrows and aborts.
                Invoke-WithRetryHC -MaxRetries $MaxRetries -Description $description -Critical -Action {
                    $record | New-ServiceNowRecord @createParams
                } | Out-Null

                $createdCount++
                #endregion
            }
        }
        #endregion

        #region Deactivate table rows that are no longer in the collection
        # Anything not matched by a source row and still active is switched off.
        # Rows already inactive are left untouched (no redundant API calls).
        foreach ($existingRecord in $existingRecords) {
            if ($matchedSysId.Contains([string]$existingRecord.sys_id)) {
                continue
            }

            if (-not (Test-RecordActiveHC $existingRecord.u_active)) {
                continue
            }

            $sysId = [string]$existingRecord.sys_id
            $description = "deactivate record '$sysId'"

            Write-Verbose $description

            # Deactivation is best-effort: a failure warns and continues.
            $deactivated = Invoke-WithRetryHC -MaxRetries $MaxRetries -Description $description -Action {
                Update-ServiceNowRecord @updateParams -ID $sysId -Values @{ u_active = $activeFalseValue }
            }

            if ($deactivated) { $deactivatedCount++ }
        }
        #endregion

        Write-Verbose "Sync complete for '$TableName': $createdCount created, $reactivatedCount reactivated, $deactivatedCount deactivated (of $($existingRecords.Count) existing record(s))."
    }
}