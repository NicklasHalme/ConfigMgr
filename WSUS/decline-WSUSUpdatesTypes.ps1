
#requires -version 3
#Requires -RunAsAdministrator
<#
.SYNOPSIS
Decline several Update Types in Windows Server Update Services (WSUS)

.DESCRIPTION
Decline several Update Types in Windows Server Update Services (WSUS). For example Beta and Preview Updates, Updates for Itanium, Drivers, Dell Hardware, Surface Hardware, SharePoint Updates in Office Channel, 
Language on Demand Feature updates and superseded updates. The scrips send, if configured a list of the decliened updates.

.COPYRIGHT 
    The MIT License (MIT)

.LICENSEURI 
    https://opensource.org/licenses/MIT

.INPUTS
  None

.OUTPUTS
  Log file stored in path specified by the variable $Logpath. Please see variable section.

.NOTES
  Base on Fabian Niesen's "decline-WSUSUpdatesTypes.ps1" script version 1.4 (Last Published: 17.12.2018) from https://www.powershellgallery.com/packages/decline-WSUSUpdatesTypes/1.4

  Version:        1
  Author:         Nicklas Halme
  Creation Date:  15.03.2016
  Purpose/Change: Initial script development
   

.EXAMPLE 
decline-WSUSUpdatesTypes.ps1 -Preview -Itanium -Superseded -SmtpServer "Mail.domai.tld" -EmailLog

.EXAMPLE
decline-WSUSUpdatesTypes.ps1 -Preview -Itanium -SharePoint 

.PARAMETER Preview
Decline Updates with the phrases -Preview- or -Beta- in the title or the attribute -beta- set

.PARAMETER Itanium
Decline Updates with the phrases -ia64- or -itanium-

.PARAMETER LanguageFeatureOnDemand
Decline Updates with the phrases -LanguageFeatureOnDemand- or -Lang Pack (Language Feature) Feature On Demand- or -LanguageInterfacePack-

.PARAMETER Sharepoint
Decline Updates with the phrases -SharePoint Enterprise Server- or -SharePoint Foundation- or -SharePoint Server- or -FAST Search Server- Some of these are part of the Office Update Channel

.PARAMETER Dell
Decline Updates with the phrases -Dell- for reducing for example the updates in the drivers category, if no Dell Hardware is used

.PARAMETER Surface
Decline Updates with the phrases -Surface- and -Microsoft-

.PARAMETER Drivers
Decline Updates with the Classification -Drivers-

.PARAMETER OfficeWebApp
Decline Updates with the phrases -Excel Web App- or -Office Web App- or -Word Web App- or -PowerPoint Web App-

.PARAMETER Officex86
Decline Updates with the phrases -32-Bit- and one of these -Microsoft Office-, -Microsoft Access-, -Microsoft Excel-, -Microsoft Outlook-, -Microsoft Onenote-, -Microsoft PowerPoint-, -Microsoft Publisher-, -Microsoft Word-

.PARAMETER Officex64
Decline Updates with the phrases -64-Bit- and one of these -Microsoft Office-, -Microsoft Access-, -Microsoft Excel-, -Microsoft Outlook-, -Microsoft Onenote-, -Microsoft PowerPoint-, -Microsoft Publisher-, -Microsoft Word-

.PARAMETER Superseded
Decline Updates with the attribute -IsSuperseded-
                     
#>
[cmdletbinding()]
Param(
	[Parameter(Position=1)]
    [string]$WsusServer = ([system.net.dns]::GetHostByName('localhost')).hostname,
	[Parameter(Position=2)]
    [bool]$UseSSL = $False,
	[Parameter(Position=3)]
    [int]$PortNumber = 8530
)
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
function Write-Log($msg,$ForegroundColor) {
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
$Updates = $null
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
$WsusServerAdminProxy = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer,$UseSSL,$PortNumber);

# Beta and Preview updates
    Write-Log "Declining of Beta and Preview updates selected, starting query."
    $BetaUpdates = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and ($_.Title -match  preview|beta  -or -not $_.IsDeclined -and $_.IsBeta -eq $true)}
    Write-Log "Found $($BetaUpdates.count) Preview or Beta Updates to decline"
    If($BetaUpdates) 
    {
      #IF (! $WhatIF) {$BetaUpdates | %{$_.Decline()}}
	  $BetaUpdates | Add-Member -MemberType NoteProperty -Name PatchType -value BetaUpdate 
      $Updates = $Updates + $BetaUpdates
        
    }
    Else
    {Write-Log "No Preview / Beta Updates found that needed declining. Come back next 'Patch Tuesday' and you may have better luck."}

# Itanium, ARM64, ia64
    Write-Log "Declining of Itanium updates selected, starting query."
    $ItaniumUpdates = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match  ia64|itanium|ARM64 }
    Write-Log "Found $($ItaniumUpdates.count) Itanium Updates to decline"
    If($ItaniumUpdates) 
    {
      #IF (! $WhatIF) {$ItaniumUpdates | %{$_.Decline()}}
      $ItaniumUpdates | Add-Member -MemberType NoteProperty -Name PatchType -value "Itanium"
      $Updates = $Updates + $ItaniumUpdates
    }
    Else
    {Write-Log "No Itanium Updates found that needed declining. Come back next 'Patch Tuesday' and you may have better luck."}  

# Language Feature on Demand 
    Write-Log "Declining of Language Feature on Demand selected, starting query."
    $LanguageFeatureOnDemandU = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match  LanguageFeatureOnDemand|Lang Pack (Language Feature) Feature On Demand|LanguageInterfacePack }
    Write-Log "Found $($LanguageFeatureOnDemandU.count) LanguageFeatureOnDemand to decline"
    If($LanguageFeatureOnDemandU) 
    {
      IF (! $WhatIF) {$LanguageFeatureOnDemandU | %{$_.Decline()}}
      $LanguageFeatureOnDemandU | Add-Member -MemberType NoteProperty -Name PatchType -value "LanguageFeatureOnDemand"
      $Updates = $Updates + $LanguageFeatureOnDemandU
    }
    Else
    {Write-Log "No LanguageFeatureOnDemand Updates found that needed declining. Come back next 'Patch Tuesday' and you may have better luck."}

# Sharepoint Updates
    Write-Log "Declining of Sharepoint Updates selected, starting query..."
    $SharepointU = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match  SharePoint Enterprise Server|SharePoint Foundation|SharePoint Server|FAST Search Server }
    Write-Log "Found $($SharepointU.count) Sharepoint Updates to decline"
    If($SharepointU) 
    {
      $SharepointU | Add-Member -MemberType NoteProperty -Name PatchType -value "SharePoint"
      $Updates = $Updates + $SharepointU
    }
    Else
    {Write-Log "No Sharepoint Updates found that needed declining. Come back next 'Patch Tuesday' and you may have better luck."}

# Drivers
    Write-Log "Declining of Drivers selected, starting query..."
    $DriversUpdates = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Classification -match  Drivers }
    Write-Log "Found $($DriversUpdates.count) Drivers to decline"
    If($DriversUpdates) 
    {
      $DriversUpdates | Add-Member -MemberType NoteProperty -Name PatchType -value "Driver Update" 
      $Updates = $Updates + $DriversUpdates
        
    }
    Else
    {Write-Log "No Driver found that needed declining. Come back next 'Patch Tuesday' and you may have better luck."}

# Superseded Updates
    Write-Log "Declining Superseded Updates selected, starting query..."
    $SupersededU = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.IsSuperseded -eq $true}
    Write-Log "Found $($SupersededU.count) Superseded Updates to decline"
    If($SupersededU) 
    {
        $SupersededU | Add-Member -MemberType NoteProperty -Name PatchType -value "Superseded"
        $Updates = $Updates + $SupersededU
    }
    Else
    {Write-Log "No IsSuperseded Updates found that needed declining. Come back next 'Patch Tuesday' and you may have better luck."}

# Results
$Updates | select $Table | sort -Property "KB Article" | ft -AutoSize -Property "Kind of Patch",Title,"KB Article"

    Write-Log "List needed updates selected, starting query."
    $updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $updatescope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::NotApproved
    $updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled
    $NeededUpdates = $WsusServerAdminProxy.GetUpdates($updatescope)
    Write-Log "Found $($NeededUpdates.count) needed Updates"
    If($NeededUpdates) 
    {
      Write-Log "Needed Updates:"
      $NeededUpdates | Select $Table | FT -AutoSize 
      Write-Log "Needed Updates "+$($NeededUpdates | Select $Table | ConvertTo-HTML -head $Style)
    }
    Else
    {Write-Log "No Needed Updates found to list. Come back next 'Patch Tuesday' and you may have better luck."}
