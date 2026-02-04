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
#                                          Main Code
#----------------------------------------------------------------------------------------------------------
Write-Log "`n==================== Starting Script: $($MyInvocation.MyCommand.Name) ====================" -ForegroundColor Yellow

# This will reset the Microsoft Store cache and may resolve minor issues.
    Write-Log "`n[Resetting Microsoft Store...]" -ForegroundColor Cyan
    Try {
    Write-Log "Resetting Microsoft Store..." -ForegroundColor Cyan
    wsreset.exe
    Write-Log "Microsoft Store reset command executed." -ForegroundColor Green
    }
    Catch {
        Write-Log "An error occurred while resetting Microsoft Store: $_" -ForegroundColor Red
    }

    # This will re-register the Microsoft Store app.
    Write-Host "`n[Restore the Microsoft Store App...]" -ForegroundColor Cyan
    try {
        Get-AppXPackage *WindowsStore* -AllUsers | ForEach-Object {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml"}
        Write-Log "Microsoft Store restored successfully." -ForegroundColor Green
    }
    catch {
        Write-Log "ERROR: An error occurred while resetting Microsoft Store: $_" -ForegroundColor Red
    }
Write-Log "`n==================== Script Completed: $($MyInvocation.MyCommand.Name) ====================" -ForegroundColor Yellow