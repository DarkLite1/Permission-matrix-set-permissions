# Import public functions first
Get-ChildItem "$PSScriptRoot/Public/*.ps1" | ForEach-Object {
    . $_.FullName
}

# Import private internals
Get-ChildItem "$PSScriptRoot/Private/*.ps1" | ForEach-Object {
    . $_.FullName
}