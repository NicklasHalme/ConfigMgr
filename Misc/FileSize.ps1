# The script lists all files in the specified directory and its subdirectories, 
# filtering out files with a size of 0 bytes. It then outputs the full path and size of each file found.
# Replace "C:\Your\Directory\Path" with the path of the directory you want to check.
Get-ChildItem -Path "C:\Your\Directory\Path" -Recurse | Where-Object { $_.Length -eq 0 } | ForEach-Object {
    Write-Output "File: $($_.FullName) - Size: $($_.Length) bytes"
}
