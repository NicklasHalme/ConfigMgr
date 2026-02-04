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


<#
.SYNOPSIS
    Microsoft Store Repair Menu Script
.DESCRIPTION
    Provides a menu to run different repair and maintenance functions.
    Includes input validation and error handling.
#>

# --- Repair Functions ---

function Reset-MSStore {
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
}

function Restore-MSStore {
    # This will re-register the Microsoft Store app.
    Write-Host "`n[Restore the Microsoft Store App...]" -ForegroundColor Cyan
    try {
        Get-AppXPackage *WindowsStore* -AllUsers | ForEach-Object {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml"}
        Write-Host "Microsoft Store restored successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: An error occurred while resetting Microsoft Store: $_" -ForegroundColor Red
    }
}
<#
function Check-Disk {
    Write-Host "`n[Checking Disk for Errors...]" -ForegroundColor Cyan
    try {
        chkdsk C: /F /R
        Write-Host "A restart may be required to complete the disk check." -ForegroundColor Yellow
    }
    catch {
        Write-Host "Error running CHKDSK: $_" -ForegroundColor Red
    }
}

function Clear-TempFiles {
    Write-Host "`n[Clearing Temporary Files...]" -ForegroundColor Cyan
    try {
        Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Temporary files cleared." -ForegroundColor Green
    }
    catch {
        Write-Host "Error clearing temp files: $_" -ForegroundColor Red
    }
}
#>

#----------------------------------------------------------------------------------------------------------
#                                          Main Code
#----------------------------------------------------------------------------------------------------------
# --- Menu Loop ---
do {
    Clear-Host
    Write-Host "=== Microsoft Store Repair Menu ===" -ForegroundColor Yellow
    Write-Host "1. Reset Microsoft Store (wsreset.exe)"
    Write-Host "2. Re-register the Microsoft Store app"
    Write-Host "3. Exit"
    Write-Host "==========================="

    $choice = Read-Host "Enter your choice (1-3)"

    switch ($choice) {
        '1' { Repair-SystemFiles }
        '2' { Restore-MSStore }
        '3' { Write-Host "Exiting..." -ForegroundColor Green }
        default { Write-Host "Invalid choice. Please enter a number between 1 and 3." -ForegroundColor Red }
    }

    if ($choice -ne '3') {
        Write-Host "`nPress Enter to return to menu..."
        [void][System.Console]::ReadLine()
    }

} while ($choice -ne '3')

