# Log File Info
$LogPath = $PSScriptRoot
$LogFileName = "Delice-WSUS-Updates-by-Type_$(get-date -format `"yyyy-MM-dd`").log"

#--------------------------------------
# Write to log file
#--------------------------------------
$LogFile = Join-Path -Path $LogPath -ChildPath $LogFileName
function Write-LogEntry($msg,$ForegroundColor) {
  If ($null -eq $ForegroundColor) { $ForegroundColor = [System.ConsoleColor]::White }
  If ((Test-Path $LogPath) -eq $False) { new-item -Path $LogPath -ItemType Directory }
  $LogDate = Get-Date -Format "yyyy-MM-dd HH.MM.ss"
  Write-Host "$LogDate : $msg" -ForegroundColor $ForegroundColor
  "$LogDate : $msg" | Out-File $LogFile -Append
}

$WsusServer = ([system.net.dns]::GetHostByName('localhost')).hostname
Write-LogEntry "WSUSServer: $($WsusServer)"
$UseSSL = $False
Write-LogEntry "UseSSL: $($UseSSL)"
$PortNumber = 8530
Write-LogEntry "PortNumber: $($PortNumber)"

[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
$WsusServerAdminProxy = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer,$UseSSL,$PortNumber);

# Updates to decline: Preview, beta, Edge-Dev
$previewUpdates = $WsusServerAdminProxy.GetUpdates() | Where-Object{-not $_.IsDeclined -and ($_.Title -like "*preview*" -or $_.Title -like "*beta*" -or $_.Title -like "*Edge-Dev*" )}
#Write-Output "Found $($previewUpdates.count) Preview updates to decline"
Write-LogEntry "Found $($previewUpdates.count) Preview, beta or Edge-Dev updates to decline"
foreach ($previewUpdate in $previewUpdates){
    #Write-Output $previewUpdate.Title " will be declined"
    Write-LogEntry "$($previewUpdate.Title) will be declined"
    $previewUpdate.Decline()
}

# Architecture updates to decline: ARM64, x86, ia64, itanium
$archUpdates = $WsusServerAdminProxy.GetUpdates() | Where-Object{-not $_.IsDeclined -and ($_.Title -like "*ARM64*" -or $_.Title -like "*x86*" -or $_.Title -like "*ia64*" -or $_.Title -like "*itanium*")}
Write-LogEntry "Found $($archUpdates.count) architecture updates to decline"
foreach ($arm64Update in $archUpdates){
    Write-LogEntry "$($archUpdate.Title) will be declined"
    $archUpdate.Decline()
}

# Windows 10 updates to decline
$Windows10Updates = $WsusServerAdminProxy.GetUpdates() | Where-Object{-not $_.IsDeclined -and ($_.Title -like "*Windows 10 version 1507*" -or $_.Title -like "*Windows 10 version 1511*" -or $_.Title -like "*Windows 10 version 1607*" -or $_.Title -like "*Windows 10 version 1903*" -or $_.Title -like "*Windows 10 version 1909*" -or $_.Title -like "*Windows 10 version 2004*" -or $_.Title -like "*Windows 10 version 20H2*" -or $_.Title -like "*Windows 10 version 21H2*")}
Write-LogEntry "Found $($Windows10Updates.count) Windows 10 updates to decline"
foreach ($Windows10Update in $Windows10Updates){
    Write-LogEntry "$($Windows10Update.Title) will be declined"
    $Windows10Update.Decline()
}




