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