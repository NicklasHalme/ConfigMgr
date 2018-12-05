<#
#requires -version 3
#Requires -RunAsAdministrator
.DESCRIPTION
  This script will get the information from the Mobile broadband interface, also called wireless wide area network (WWAN) service.
  The collected information is then written to MIF -file that the SCCM will then collect to HW inventory.

.COPYRIGHT 
    The MIT License (MIT)

.LICENSEURI 
    https://opensource.org/licenses/MIT

.PARAMETER <Parameter_Name>
    None

.INPUTS
  None

.OUTPUTS
  Log file stored in path specified by the variable $Logpath. Please see variable section.
  MIF -file is created to %windir%\CCM\Inventory\Noidmifs\MobileBroadband.mif

.NOTES
  Documentation: https://contosoniku.wordpress.com/2016/12/06/a-step-by-step-guide-to-extend-hardware-inventory-in-sccm-for-collecting-sim-iccid/

  Version:        1
  Author:         Nicklas Halme
  Creation Date:  15.03.2016
  Purpose/Change: Initial script development
  
  Version:        2
  Author:         Nicklas Halme
  Creation Date:  05.12.2018
  Change:         - Changed the logic how the information is processed based on Todd Wilkolak and Clayton's feedback
                  - Modified the Logging -function
                  - Removed the Parse -function
   
.EXAMPLE
  None
#>

#----------------------------------------------------------------------------------------------------------
#
#                                         Global Editable Variables
#
#---------------------------------------------------------------------------------------------------------- 

#Log File Info
$LogPath = "$env:SystemRoot\Logs"
$LogFileName = "$env:Computername-MobileBroadband_$(get-date -format `"yyyy-MM-dd`").log"
$LogFile = Join-Path -Path $LogPath -ChildPath $LogFileName

#----------------------------------------------------------------------------------------------------------
#
#                                          Function Definition
#
#---------------------------------------------------------------------------------------------------------- 
#
#--------------------------------------
# Write to log file
#--------------------------------------
function Write-LogEntry($msg,$ForegroundColor) {
  if ($ForegroundColor -eq $null) { $ForegroundColor = [System.ConsoleColor]::White }
  If ((Test-Path $LogPath) -eq $False) { new-item -Path $LogPath -ItemType Directory }
  $LogDate = Get-Date -Format "yyyy-MM-dd HH.MM.ss"
  Write-Host "$LogDate : $msg" -ForegroundColor $ForegroundColor
  "$LogDate : $msg" | Out-File $LogFile -Append
}

#----------------------------------------------------------------------------------------------------------------
#
#                               Main Execution 
#
#----------------------------------------------------------------------------------------------------------------
Write-LogEntry "******** Script Started ********" -ForegroundColor cyan
Write-LogEntry "Logging to: $LogFile"

Write-LogEntry "Get the content from the 'netsh mbn show interface' command."
$get_MBInfo = cmd /c "netsh mbn show interface"

Write-LogEntry "Read the content..."
If ($get_mbinfo -like "*There is no Mobile Broadband interface*") {
    Write-LogEntry "There is no Mobile Broadband interface. Set the inventory status..." -ForegroundColor Yellow
    $DeviceID = "No Mobile Broadband interface"
    $ICCID = "No Mobile Broadband interface"
}
If ($get_mbinfo -like "*Mobile Broadband Service (wwansvc) is not running.*") {
    Write-LogEntry "wwansvc service is not running. Set the inventory status..." -ForegroundColor Yellow
    $DeviceID = "wwansvc is not running"
    $ICCID = "wwansvc is not running"
}
# Read the second line (NOTE: line numbering start from 0) of the content
If ($get_mbinfo[1] -like "*There is 1 interface on the system:*") {
    Write-LogEntry "There is 1 interface on the system..." -ForegroundColor Green
    # Read info from specific line and split the line with ':'-delimeter. Trim spaces from the beginning and end of a string.
    Write-LogEntry "Get Provider Name: Read info from specific line (16) and split the line with ':'-delimeter. Trim spaces from the beginning and end of a string."
    $ProviderName = $get_mbinfo[16].Split(":")[1].Trim()
    
    Write-LogEntry "Get the content from the 'netsh mbn show read i=*' command."
    $get_iccid = cmd /c "netsh mbn show read i=*"
    Write-LogEntry "Get SIM ICC ID: Read info from specific line (5) and split the line with ':'-delimeter. Trim spaces from the beginning and end of a string." -ForegroundColor Green
    $ICCID = $get_iccid[5].Split(":")[1].Trim()
}

Write-LogEntry "Create the inventory results to MIF -file (MobileBroadband.mif) and save it to %Windir%\CCM\Inventory\Noidmifs -folder"
# Create the inventory results to MIF -file and save it to %Windir%\CCM\Inventory\Noidmifs -folder
$outfile="$env:windir\CCM\Inventory\Noidmifs\MobileBroadband.mif"
If (Test-Path $outfile) {
    Write-LogEntry "Previous file exist. Let's remove it!"
    Remove-Item $outfile -Force
}
'Start Component' | Out-File $outfile -Encoding default -Append
'Name = "System_InformationInventory"' | Out-File $outfile -Encoding default -Append
'Start Group' | Out-File $outfile -Encoding default -Append
'Name = "MobileBroadband_Information"' | Out-File $outfile -Encoding default -Append
'ID = 1' | Out-File $outfile -Encoding default -Append
'Class = "SIM_ICCID"' | Out-File $outfile -Encoding default -Append
'Start Attribute' | Out-File $outfile -Encoding default -Append
'Name = "SIM_ICCID"' | Out-File $outfile -Encoding default -Append
'ID = 1' | Out-File $outfile -Encoding default -Append
'Type = String(80)' | Out-File $outfile -Encoding default -Append
'Value = "'+$ICCID+'"' | Out-File $outfile -Encoding default -Append
'End Attribute' | Out-File $outfile -Encoding default -Append
'Start Attribute' | Out-File $outfile -Encoding default -Append
'Name = "Provider_name"' | Out-File $outfile -Encoding default -Append
'ID = 2' | Out-File $outfile -Encoding default -Append
'Type = String(80)' | Out-File $outfile -Encoding default -Append
'Value = "'+$ProviderName+'"' | Out-File $outfile -Encoding default -Append
'End Attribute' | Out-File $outfile -Encoding default -Append
'End Group' | Out-File $outfile -Encoding default -Append
'End Component' | Out-File $outfile -Encoding default -Append
Write-LogEntry "MIF -file (MobileBroadband.mif) created to %Windir%\CCM\Inventory\Noidmifs -folder. Bye!"
Write-LogEntry "******** Script Ended ********" -ForegroundColor cyan
