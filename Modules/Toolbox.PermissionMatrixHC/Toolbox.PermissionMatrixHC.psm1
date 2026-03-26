# ---------------------------------------------------------------------
# Load PUBLIC functions
# ---------------------------------------------------------------------
$publicFolder = Join-Path $PSScriptRoot 'Public'
if (Test-Path $publicFolder) {
    Get-ChildItem -Path $publicFolder -Filter *.ps1 | ForEach-Object {
        Write-Verbose "Load '$_'"
        . $_.FullName
    }
}

# ---------------------------------------------------------------------
# Load PRIVATE functions
# ---------------------------------------------------------------------
$privateFolder = Join-Path $PSScriptRoot 'Private'
if (Test-Path $privateFolder) {
    Get-ChildItem -Path $privateFolder -Filter *.ps1 | ForEach-Object {
        Write-Verbose "Load '$_'"
        . $_.FullName
    }
}

# Export-ModuleMember -Function * -Alias *


Export-ModuleMember -Function @(
    'Invoke-PermissionMatrix',
    'Export-PermissionMatrix'
)
