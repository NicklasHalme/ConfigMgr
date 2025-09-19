Lets you execute dynamically generated code or commands contained in a string. This can be very powerful but also potentially risky if used improperly.
Powershell.exe -Command "iex (gc MyFile.ps1 -raw)"
powershell.exe -command "Invoke-Expression (Get-Content MyFile.ps1 -raw)"