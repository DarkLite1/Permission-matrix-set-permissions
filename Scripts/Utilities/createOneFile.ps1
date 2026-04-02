param (
    [string]$SourceFolderPath = (Resolve-Path "$PSScriptRoot\..\..").Path,
    [string]$OutputFile = 'AllCode.txt', 
    [HashTable]$Ignored = @{
        Folders = @(
            '.vscode',
            '.git',
            'Tests',
            'legacy'
        )
        Files   = @(
            '.gitignore',
            '.prettierrc.json'
        )
    }
)

function Convert-ToNativePath {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [string]$Path
    )
    begin {
        $Separator = [System.IO.Path]::DirectorySeparatorChar
    }
    process {
        $Path.Replace('\', $Separator).Replace('/', $Separator)
    }
}

# 1. Resolve Output file to a full absolute path
$OutputFilePath = Join-Path -Path $SourceFolderPath -ChildPath $OutputFile

# 2. Transform the relative ignore lists into full native paths
$ignoredFolders = @(
    $Ignored.Folders | ForEach-Object { 
        Join-Path -Path $SourceFolderPath -ChildPath $_ | Convert-ToNativePath 
    }
)

$ignoredFiles = @(
    $Ignored.Files | ForEach-Object { 
        Join-Path -Path $SourceFolderPath -ChildPath $_ | Convert-ToNativePath 
    }
)

# 3. Dynamically add the script itself and the output file to the ignore list
$ignoredFiles += $PSCommandPath | Convert-ToNativePath
$ignoredFiles += $OutputFilePath | Convert-ToNativePath 

# Initialize or clear the output file
New-Item -ItemType File -Path $OutputFilePath -Force | Out-Null

# Setup the traversal queue
$DirectoriesToProcess = [System.Collections.Generic.List[string]]::new()
$DirectoriesToProcess.Add($SourceFolderPath)

while ($DirectoriesToProcess.Count -gt 0) {
    
    $CurrentPath = $DirectoriesToProcess[0]
    $DirectoriesToProcess.RemoveAt(0)
    
    $Children = Get-ChildItem -Path $CurrentPath

    foreach ($Child in $Children) {
        
        if ($Child.PSIsContainer) {
            # Check if the folder is in the full-path ignore list
            if ($ignoredFolders -contains $Child.FullName) {
                Write-Verbose "Folder ignored: $($Child.FullName)"
                continue
            }
            
            # If not ignored, add it to the queue
            $DirectoriesToProcess.Add($Child.FullName)
        } 
        else {
            # Check if the file is in the full-path ignore list (which now includes the script & output file)
            if ($ignoredFiles -contains $Child.FullName) {
                Write-Verbose "File ignored: $($Child.FullName)"
                continue
            }

            # Safely calculate the relative path by cutting off the base path length
            $relativePath = $Child.FullName.Substring($SourceFolderPath.Length).TrimStart('\/')
            $fileContent = Get-Content -Path $Child.FullName -Raw # -Raw reads it as one big string for speed

            Write-Verbose "Adding file content: $relativePath"

            $fileContentToAdd = @"

-------------------------------------------------------------------------------
FILE: $relativePath
-------------------------------------------------------------------------------

$fileContent
"@
            # Append the formatted content and a trailing newline
            ($fileContentToAdd, "`n") | Add-Content -Path $OutputFilePath -Force
        }
    }
}

Write-Host "Success! Consolidated file created at: $OutputFilePath" -ForegroundColor Cyan