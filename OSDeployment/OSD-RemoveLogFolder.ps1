#requires -version 3
<# #Requires -RunAsAdministrator
.DESCRIPTION
  Removes the previous deployment logs for the current computer

.COPYRIGHT 
    The MIT License (MIT)

.LICENSEURI 
    https://opensource.org/licenses/MIT

.PARAMETER <SLShare>
    Path to the network shared folder in which the deployment logs are stored at the end of the deployment process.
.PARAMETER <OSDComputerName>
    The computer name of the current computer.

.INPUTS
  Paramaeters form command line

.OUTPUTS
  None

.NOTES
  Version:        1
  Author:         Nicklas Halme, Microsoft
  Creation Date:  22.11.2018
  Purpose/Change: Initial script development

  Version:        2
  Author:         FirstName LastName
  Creation Date:  dd.mm.yyyy
  Purpose/Change: improved...
  
.EXAMPLE
  powershell.exe -noprofile -executionpolicy bypass -force -file .\RemoveLogFolder.ps1 -SLShare %SLShare% -OSDComputerName %OSDComputerName%
#>
Param
(
    [parameter(Mandatory=$true)]
    [String[]]
    $SLShare,
    [parameter(Mandatory=$true)]
    [String[]]
    $OSDComputerName
)

    Remove-Item -Recurse -Force -Path "$SLShare\$OSDComputerName" -ErrorAction SilentlyContinue
