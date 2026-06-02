#region Load private functions
$privateFolder = Join-Path -Path $PSScriptRoot -ChildPath 'Private'

if (Test-Path -LiteralPath $privateFolder -PathType Container) {
    $privateFiles = Get-ChildItem -LiteralPath $privateFolder -Filter '*.ps1' -Recurse -File
    
    foreach ($file in $privateFiles) {
        Write-Verbose "Loading Private Function: $($file.Name)"
        . $file.FullName
    }
}
#endregion

#region Load public functions and export
$publicFolder = Join-Path -Path $PSScriptRoot -ChildPath 'Public'
$functionsToExport = [System.Collections.Generic.List[string]]::new()

if (Test-Path -LiteralPath $publicFolder -PathType Container) {
    $publicFiles = Get-ChildItem -LiteralPath $publicFolder -Filter '*.ps1' -Recurse -File
    
    foreach ($file in $publicFiles) {
        Write-Verbose "Loading Public Function: $($file.Name)"
        . $file.FullName
        
        # Add the script's name (without the .ps1 extension) to our export list
        $functionsToExport.Add($file.BaseName)
    }
}

# Automatically export everything loaded from the Public folder
if ($functionsToExport.Count -gt 0) {
    Export-ModuleMember -Function $functionsToExport.ToArray()
}
#endregion