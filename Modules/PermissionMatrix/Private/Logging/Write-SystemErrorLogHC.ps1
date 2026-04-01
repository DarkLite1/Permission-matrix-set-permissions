function Write-SystemErrorLogHC {
    <#
        Creates JSON log file of system errors and attaches to email params.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$SystemErrors,
        [Parameter(Mandatory)][string]$LogFolder,
        [Parameter(Mandatory)][ref]$MailParams,
        [datetime]$ScriptStartTime = (Get-Date),
        [object]$JsonFileItem = @{ BaseName = 'MatrixConfig' }
    )

    if ($SystemErrors.Count -eq 0) { return }
    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return }

    $datedFolder = Get-DatedLogFolderPathHC `
        -LogFolder $LogFolder `
        -ScriptStartTime $ScriptStartTime `
        -JsonFileItem $JsonFileItem

    $partial = Join-Path $datedFolder 'SystemErrors'

    $attachments = Out-LogFileHC `
        -DataToExport $SystemErrors `
        -PartialPath $partial `
        -FileExtensions '.json' `
        -ErrorAction Ignore

    if ($attachments) {
        if (-not $MailParams.Value.ContainsKey('Attachments')) {
            $MailParams.Value['Attachments'] = @()
        }
        $MailParams.Value['Attachments'] += $attachments
    }
}

