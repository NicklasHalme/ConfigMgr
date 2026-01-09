#----------------------------------------------------------------------------------------------------------
#
#                                          Function Definition
#
#---------------------------------------------------------------------------------------------------------- 
#
$LogPath = $PSScriptRoot
$LogFileName = $MyInvocation.MyCommand.Name.log
$LogFile = Join-Path -Path $LogPath -ChildPath $LogFileName
#--------------------------------------
# Write to log file
#--------------------------------------
function Write-Log($msg,$ForegroundColor) {
  if ($null -eq $ForegroundColor) { $ForegroundColor = [System.ConsoleColor]::White }
  If ((Test-Path $LogPath) -eq $False) { new-item -Path $LogPath -ItemType Directory }
  $LogDate = Get-Date -Format "yyyy-MM-dd HH.MM.ss"
  Write-Host "$LogDate : $msg" -ForegroundColor $ForegroundColor
  "$LogDate : $msg" | Out-File $LogFile -Append
}
#----------------------------------------------------------------------------------------------------------
# Name of the software to check
$softwareNames = @("ProPlus2021Volume - en-us", "ProPlus2021Volume - fi-fi")
# Silent uninstall parameters
$silentUninstall = "DisplayLevel=False forceappshutdown=true"

# Check if software is installed
foreach ($softwareName in $softwareNames){
    $software = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.PSChildName -match $softwareName }

    # Uninstall the software if it is installed
    if ($null -ne $software) {
        # Software is installed, uninstall it
        try {
            Write-Log "Uninstalling $softwareName..."
            & $software.UninstallString + " " + $silentUninstall
            Write-Log "$softwareName has been uninstalled."
        }
        catch {
            Write-Log "Failed to uninstall $softwareName. Error: $_"
        }
    }
    else {
        # Software is not installed
        Write-Log "$softwareName is not installed."
    }
}
