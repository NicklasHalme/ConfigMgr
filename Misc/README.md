Powershell.exe -Command "iex (gc MyFile.ps1 -raw)"
powershell.exe -command "Invoke-Expression (Get-Content MyFile.ps1 -raw)"