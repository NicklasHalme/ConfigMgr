#requires -version 3

<#Requires –Modules ActiveDirectory
##Requires -RunAsAdministrator
.SYNOPSIS
  Add or delete computers from databases.

.DESCRIPTION
  With this tool you can add computers to MDT Database. You can add one by one, or use a CSV file to add many computers at once.
  You can also remove computers from AD, ConfigMgr and MDT databases. You can remove one by one, or use a CSV file to add many computers at once.

.COPYRIGHT 
    The MIT License (MIT)

.LICENSEURI 
    https://opensource.org/licenses/MIT

.INPUTS
  Fill in the Text Boxies in the tab you want to use or browse for CSV -file.

.OUTPUTS
  Log file stored in the script root.

.REFERENS FILES
    - ico.ico -> 32x32 px icon file alongside this script. Showinf the logo
    - MDTDB.psm1 -> MDT PowerShell module, provided by Michael Niehaus
    - AD PowerShell module available on the computer running this script

.NOTES
  
  You need to fix an issue with a view in the MDT database used by the MDTDB.psm1 module. Otherwise this script will not work.
  More about it here: https://syscenramblings.wordpress.com/2016/01/15/mdt-database-the-powershell-module-fix/

  NOTE!! Make sure you edit the global variables to reflect your environment.

  Version/date:   07.04.2020
  Author:         Nicklas Halme
  Change history: Initial script development
  
#>

#----------------------------------------------------------------------------------------------------------
#
#                                         Global Editable Variables
#
#---------------------------------------------------------------------------------------------------------- 
$ScriptVersion = "07.04.2020"

# Log File Info
$LogPath = "$PSScriptRoot"
$LogFileName = "$env:Computername-Logging_$(get-date -format `"yyyy-MM-dd`").log"


# MDT Server
$MDTsqlServer = "CM01"
$MDTinstance = ""            # Empty if Default Instance is used (MSSQLSERVER)
$MDTdatabase = "MDT"

# ConfigMgr Site Server
$SCCMSiteServer = "CM01"
$SCCMSiteCode = "NHL"


#----------------------------------------------------------------------------------------------------------
#
#                                          Function Definition
#
#---------------------------------------------------------------------------------------------------------- 
#--------------------------------------
# Write to log file
#--------------------------------------
$LogFile = Join-Path -Path $LogPath -ChildPath $LogFileName
function Write-LogEntry($msg,$ForegroundColor) {
  If ($ForegroundColor -eq $null) { $ForegroundColor = [System.ConsoleColor]::White }
  If ((Test-Path $LogPath) -eq $False) { new-item -Path $LogPath -ItemType Directory }
  $LogDate = Get-Date -Format "yyyy-MM-dd HH.MM.ss"
  Write-Host "$LogDate : $msg" -ForegroundColor $ForegroundColor
  "$LogDate : $msg" | Out-File $LogFile -Append
}

#--------------------------------------
# Create the form
#--------------------------------------
function CreateForm {
    #[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    #[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.drawing
    
    # Form Setup
    $form1 = New-Object System.Windows.Forms.Form
    #$listBox = New-Object System.Windows.Forms.ListBox
    $Button_Run_Add_One = New-Object System.Windows.Forms.Button
    $Button_Browse_Add_CSV = New-Object System.Windows.Forms.Button
    $Button_Run_Add_CSV = New-Object System.Windows.Forms.Button
    $Button_Run_Remove_One = New-Object System.Windows.Forms.Button
    $Button_Reset_Remove_One = New-Object System.Windows.Forms.Button
    $Button_Browse_Remove_CSV = New-Object System.Windows.Forms.Button
    $Button_Run_Remove_CSV = New-Object System.Windows.Forms.Button
    $checkBox_AD = New-Object System.Windows.Forms.CheckBox
    $checkBox_SCCM = New-Object System.Windows.Forms.CheckBox
    $checkBox_MDT = New-Object System.Windows.Forms.CheckBox
    $TabControl = New-object System.Windows.Forms.TabControl
    $AddOne = New-Object System.Windows.Forms.TabPage
    $AddCSV = New-Object System.Windows.Forms.TabPage    
    $RemoveOne = New-Object System.Windows.Forms.TabPage
    $RemoveCSV = New-Object System.Windows.Forms.TabPage

    
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    
    # Form Parameter
    $form1.Text = "Työasemahallinta - versio $ScriptVersion"
    $form1.Name = "CompMgmt"
    $form1.TopMost = $true
    $form1.ControlBox = $true
    $form1.MaximizeBox = $false
    $form1.FormBorderStyle = "Fixed3D"
    $form1.Icon = New-Object system.drawing.icon("$PSScriptRoot\ico.ico")
    $form1.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 510
    $System_Drawing_Size.Height = 280
    $form1.ClientSize = $System_Drawing_Size
    
    #Tab Control 
    $tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 30
    $System_Drawing_Point.Y = 30
    $tabControl.Location = $System_Drawing_Point
    $tabControl.Name = "tabControl"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 230
    $System_Drawing_Size.Width = 450
    $tabControl.Size = $System_Drawing_Size
    $form1.Controls.Add($tabControl)
    
    #---------------------------
    # Tabs
    #---------------------------
    #AddOne Page
    $AddOne.DataBindings.DefaultDataSourceUpdateMode = 0
    $AddOne.UseVisualStyleBackColor = $True
    $AddOne.Name = "AddOne"
    $AddOne.Text = "Lisää kone"
    $tabControl.Controls.Add($AddOne)
    
    #$AddCSV Page
    $AddCSV.DataBindings.DefaultDataSourceUpdateMode = 0
    $AddCSV.UseVisualStyleBackColor = $True
    $AddCSV.Name = "AddCSV"
    $AddCSV.Text = "Lisää koneita CSV:n avulla"
    $tabControl.Controls.Add($AddCSV)
        
    #Remove Single Computer Page
    $RemoveOne.DataBindings.DefaultDataSourceUpdateMode = 0
    $RemoveOne.UseVisualStyleBackColor = $True
    $RemoveOne.Name = "RemoveOne"
    $RemoveOne.Text = "Poista kone"
    $tabControl.Controls.Add($RemoveOne)
    
    #Remove Computers with CSV Page
    $RemoveCSV.DataBindings.DefaultDataSourceUpdateMode = 0
    $RemoveCSV.UseVisualStyleBackColor = $True
    $RemoveCSV.Name = "RemoveCSV"
    $RemoveCSV.Text = "Poista koneita CSV:n avulla"
    $tabControl.Controls.Add($RemoveCSV)
    
    #---------------------------
    # Labels
    #---------------------------

    #Add Label to main form
    $Label_Main = New-Object System.Windows.Forms.Label
    $Label_Main.Location = New-Object System.Drawing.Size(10,10)  
    $Label_Main.Size = New-Object System.Drawing.Size(300,20)  
    $Label_Main.Text = "Tällä työkalulla voit lisätä tai poistaa koneita."
    $form1.Controls.Add($Label_Main)

    #Add Label to add one tab
    $Label_AddOne_CName = New-Object System.Windows.Forms.Label
    $Label_AddOne_CName.Location = New-Object System.Drawing.Size(20,13)  
    $Label_AddOne_CName.Size = New-Object System.Drawing.Size(150,20)  
    $Label_AddOne_CName.Text = "Anna koneen nimi:"
    $AddOne.Controls.Add($Label_AddOne_CName)

    #Add Label to add one tab
    $Label_AddOne_MAC = New-Object System.Windows.Forms.Label
    $Label_AddOne_MAC.Location = New-Object System.Drawing.Size(20,40)  
    $Label_AddOne_MAC.Size = New-Object System.Drawing.Size(150,20)  
    $Label_AddOne_MAC.Text = "Anna koneen MAC:"
    $AddOne.Controls.Add($Label_AddOne_MAC)

    #Add Label to add one tab
    $Label_AddOne_TXT = New-Object System.Windows.Forms.Label
    $Label_AddOne_TXT.Location = New-Object System.Drawing.Size(20,80)  
    $Label_AddOne_TXT.Size = New-Object System.Drawing.Size(250,20)  
    $Label_AddOne_TXT.Text = "*** Kone lisätään vain MDT -kantaan! ***"
    $AddOne.Controls.Add($Label_AddOne_TXT)

    #Add Label to add CSVe tab
    $Label_AddCSV_TXT = New-Object System.Windows.Forms.Label
    $Label_AddCSV_TXT.Location = New-Object System.Drawing.Size(20,80)  
    $Label_AddCSV_TXT.Size = New-Object System.Drawing.Size(250,20)  
    $Label_AddCSV_TXT.Text = "*** Koneet lisätään vain MDT -kantaan! ***"
    $AddCSV.Controls.Add($Label_AddCSV_TXT)

    #Add Label to remove one tab
    $Label_RemoveOne = New-Object System.Windows.Forms.Label
    $Label_RemoveOne.Location = New-Object System.Drawing.Size(20,13)  
    $Label_RemoveOne.Size = New-Object System.Drawing.Size(100,20)  
    $Label_RemoveOne.Text = "Anna koneen nimi:"
    $RemoveOne.Controls.Add($Label_RemoveOne)
    
    #---------------------------
    # TextBoxes
    #---------------------------
    #Add TextBox on Add One Tab
    $TextBox_AddOne_CName = New-Object System.Windows.Forms.TextBox 
    $TextBox_AddOne_CName.Location = New-Object System.Drawing.Size(190,10) 
    $TextBox_AddOne_CName.Size = New-Object System.Drawing.Size(200,60)
    $TextBox_AddOne_CName.CharacterCasing = "Upper"  
    $AddOne.Controls.Add($TextBox_AddOne_CName)

    #Add TextBox on Add One Tab
    $TextBox_AddOne_MAC= New-Object System.Windows.Forms.TextBox 
    $TextBox_AddOne_MAC.Location = New-Object System.Drawing.Size(190,35) 
    $TextBox_AddOne_MAC.Size = New-Object System.Drawing.Size(200,60)  
    $AddOne.Controls.Add($TextBox_AddOne_MAC)

    #TextBox on Add CSV Page
    $TextBox_Browse_AddCSV = New-Object System.Windows.Forms.TextBox 
    $TextBox_Browse_AddCSV.Location = New-Object System.Drawing.Size(20,40) 
    $TextBox_Browse_AddCSV.Size = New-Object System.Drawing.Size(400,60)
    $AddCSV.Controls.Add($TextBox_Browse_AddCSV)

    #TextBox on Remove One Page
    $TextBox_RemoveOne_CName = New-Object System.Windows.Forms.TextBox 
    $TextBox_RemoveOne_CName.Location = New-Object System.Drawing.Size(190,10) 
    $TextBox_RemoveOne_CName.Size = New-Object System.Drawing.Size(200,60)
    $TextBox_RemoveOne_CName.CharacterCasing = "Upper"
    $RemoveOne.Controls.Add($TextBox_RemoveOne_CName)

    #TextBox on Remove CSV Page
    $TextBox_Browse_RemoveCSV = New-Object System.Windows.Forms.TextBox 
    $TextBox_Browse_RemoveCSV.Location = New-Object System.Drawing.Size(20,40) 
    $TextBox_Browse_RemoveCSV.Size = New-Object System.Drawing.Size(400,60)
    $RemoveCSV.Controls.Add($TextBox_Browse_RemoveCSV)

    #---------------------------
    # Check Boxes for Clean Actions
    #---------------------------
    # CheckBox Clean AD
    $checkBox_AD.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 110
    $System_Drawing_Size.Height = 24
    $checkBox_AD.Size = $System_Drawing_Size
    $checkBox_AD.TabIndex = 0
    $checkBox_AD.Text = "Poista AD:sta"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 25
    $System_Drawing_Point.Y = 60
    $checkBox_AD.Location = $System_Drawing_Point
    $checkBox_AD.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox_AD.Name = "checkBox_AD"
    $checkBox_AD.Checked = $True
    $RemoveOne.Controls.Add($checkBox_AD)
        
    #CheckBox Clean SCCM
    $checkBox_SCCM.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 110
    $System_Drawing_Size.Height = 24
    $checkBox_SCCM.Size = $System_Drawing_Size
    $checkBox_SCCM.TabIndex = 1
    $checkBox_SCCM.Text = "Poista SCCM:stä"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 25
    $System_Drawing_Point.Y = 80
    $checkBox_SCCM.Location = $System_Drawing_Point
    $checkBox_SCCM.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox_SCCM.Name = "checkBox_SCCM"
    $checkBox_SCCM.Checked = $True
    $RemoveOne.Controls.Add($checkBox_SCCM)
    
    #CheckBox Clean MDT
    $checkBox_MDT.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 110
    $System_Drawing_Size.Height = 24
    $checkBox_MDT.Size = $System_Drawing_Size
    $checkBox_MDT.TabIndex = 0
    $checkBox_MDT.Text = "Poista MDT:stä"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 25
    $System_Drawing_Point.Y = 100
    $checkBox_MDT.Location = $System_Drawing_Point
    $checkBox_MDT.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox_MDT.Name = "checkBox_MDT"
    $checkBox_MDT.Checked = $false
    $RemoveOne.Controls.Add($checkBox_MDT)

    #---------------------------
    # Buttons
    #---------------------------
    #---------------------------
    # Add buttons actions
    #---------------------------
    
    #Button Add Action on Add One Tab
    $button_Run_AddOne_RunOnClick=
    {   
        AddMDT  
    }
    
    #Button Browsen Action on Add CSV Tab
     $button_Browse_AddCSV_RunOnClick= 
     {   
         GetFileNameAdd
     }
 
     #Button Run Action on Add CSV Tab
     $button_Run_AddCSV_RunOnClick= 
     {   
         ImportAddCSV
     }
    #---------------------------
    # Remove buttons actions
    #---------------------------
    #Button Run Action on Remove One Tab
    $button_RemoveOne_RunOnClick= 
    {   
        if ($checkBox_AD.Checked)     {  CleanAD }
        if ($checkBox_SCCM.Checked)    {  CleanSCCM }
        if ($checkBox_MDT.Checked)    {  CleanMDT }   
    }
    
    #Button Reset Action on Remove One Tab
    $button_Remove_ResetOnClick= 
    {   
        if ($checkBox_AD.Checked) {$checkBox_AD.CheckState = 0}
        if ($checkBox_SCCM.Checked) {$checkBox_SCCM.CheckState = 0}
        if ($checkBox_MDT.Checked) {$checkBox_MDT.CheckState = 0}
    }

    #Button Browsen Action on Remove CSV Tab
    $button_Browse_Remove_CSV_RunOnClick= 
    {   
        GetFileNameRemove
    }

    #Button Run Action on Remove CSV Tab
    $button_Run_RemoveCSV_RunOnClick= 
    {   
        ImportRemoveCSV
    }
    
    $OnLoadForm_StateCorrection=
    {
        $form1.WindowState = $InitialFormWindowState
    }

    #Button_Run_Add_One
    $Button_Run_Add_One.TabIndex = 4
    $Button_Run_Add_One.Name = "button_Run_Add_One"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 25
    $Button_Run_Add_One.Size = $System_Drawing_Size
    $Button_Run_Add_One.UseVisualStyleBackColor = $True
    $Button_Run_Add_One.Text = "Suorita"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 300
    $System_Drawing_Point.Y = 150
    $Button_Run_Add_One.Location = $System_Drawing_Point
    $Button_Run_Add_One.DataBindings.DefaultDataSourceUpdateMode = 0
    $Button_Run_Add_One.add_Click($button_Run_AddOne_RunOnClick)
    $AddOne.Controls.Add($Button_Run_Add_One)
 
    #Button_Browse_Add_CSV
    $Button_Browse_Add_CSV.TabIndex = 4
    $Button_Browse_Add_CSV.Name = "Button_Browse_Add_CSV"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 110
    $System_Drawing_Size.Height = 25
    $Button_Browse_Add_CSV.Size = $System_Drawing_Size
    $Button_Browse_Add_CSV.UseVisualStyleBackColor = $True
    $Button_Browse_Add_CSV.Text = "Hae CSV tiedosto"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 20
    $System_Drawing_Point.Y = 10
    $Button_Browse_Add_CSV.Location = $System_Drawing_Point
    $Button_Browse_Add_CSV.DataBindings.DefaultDataSourceUpdateMode = 0
    $Button_Browse_Add_CSV.add_Click($button_Browse_AddCSV_RunOnClick)
    $AddCSV.Controls.Add($Button_Browse_Add_CSV)

    #Button_Run_Add_CSV
    $Button_Run_Add_CSV.TabIndex = 4
    $Button_Run_Add_CSV.Name = "button_Run_AddCSV"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 25
    $Button_Run_Add_CSV.Size = $System_Drawing_Size
    $Button_Run_Add_CSV.UseVisualStyleBackColor = $True
    $Button_Run_Add_CSV.Text = "Suorita"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 300
    $System_Drawing_Point.Y = 150
    $Button_Run_Add_CSV.Location = $System_Drawing_Point
    $Button_Run_Add_CSV.DataBindings.DefaultDataSourceUpdateMode = 0
    $Button_Run_Add_CSV.add_Click($button_Run_AddCSV_RunOnClick)
    $AddCSV.Controls.Add($Button_Run_Add_CSV)

    #---------------------------
    # Remove buttons
    #---------------------------  
    #Button_Run_Remove_One
    $Button_Run_Remove_One.TabIndex = 4
    $Button_Run_Remove_One.Name = "Button_Run_Remove_One"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 25
    $Button_Run_Remove_One.Size = $System_Drawing_Size
    $Button_Run_Remove_One.UseVisualStyleBackColor = $True
    $Button_Run_Remove_One.Text = "Suorita"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 300
    $System_Drawing_Point.Y = 150
    $Button_Run_Remove_One.Location = $System_Drawing_Point
    $Button_Run_Remove_One.DataBindings.DefaultDataSourceUpdateMode = 0
    $Button_Run_Remove_One.add_Click($button_RemoveOne_RunOnClick)
    $RemoveOne.Controls.Add($Button_Run_Remove_One)
    
    #Button_Reset_Remove_One
    $Button_Reset_Remove_One.TabIndex = 5
    $Button_Reset_Remove_One.Name = "button_reset"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 25
    $Button_Reset_Remove_One.Size = $System_Drawing_Size
    $Button_Reset_Remove_One.UseVisualStyleBackColor = $True
    $Button_Reset_Remove_One.Text = "Tyhjennä valinnat"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 20
    $System_Drawing_Point.Y = 150
    $Button_Reset_Remove_One.Location = $System_Drawing_Point
    $Button_Reset_Remove_One.DataBindings.DefaultDataSourceUpdateMode = 0
    $Button_Reset_Remove_One.add_Click($button_Remove_ResetOnClick)
    $RemoveOne.Controls.Add($Button_Reset_Remove_One)

    #Button_Browse_Remove_CSV
    $Button_Browse_Remove_CSV.TabIndex = 4
    $Button_Browse_Remove_CSV.Name = "Button_Browse_Remove_CSV"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 110
    $System_Drawing_Size.Height = 25
    $Button_Browse_Remove_CSV.Size = $System_Drawing_Size
    $Button_Browse_Remove_CSV.UseVisualStyleBackColor = $True
    $Button_Browse_Remove_CSV.Text = "Hae CSV tiedosto"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 20
    $System_Drawing_Point.Y = 10
    $Button_Browse_Remove_CSV.Location = $System_Drawing_Point
    $Button_Browse_Remove_CSV.DataBindings.DefaultDataSourceUpdateMode = 0
    $Button_Browse_Remove_CSV.add_Click($button_Browse_Remove_CSV_RunOnClick)
    $RemoveCSV.Controls.Add($Button_Browse_Remove_CSV)

    #Button_Run_Remove_CSV
    $Button_Run_Remove_CSV.TabIndex = 4
    $Button_Run_Remove_CSV.Name = "Button_Run_Remove_CSV"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 75
    $System_Drawing_Size.Height = 25
    $Button_Run_Remove_CSV.Size = $System_Drawing_Size
    $Button_Run_Remove_CSV.UseVisualStyleBackColor = $True
    $Button_Run_Remove_CSV.Text = "Suorita"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 300
    $System_Drawing_Point.Y = 150
    $Button_Run_Remove_CSV.Location = $System_Drawing_Point
    $Button_Run_Remove_CSV.DataBindings.DefaultDataSourceUpdateMode = 0
    $Button_Run_Remove_CSV.add_Click($button_Run_RemoveCSV_RunOnClick)
    $RemoveCSV.Controls.Add($Button_Run_Remove_CSV)
    
    #Save the initial state of the form
    $InitialFormWindowState = $form1.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $form1.add_Load($OnLoadForm_StateCorrection)
    #Show the Form
    $form1.ShowDialog() | Out-Null
} #End function CreateForm
#---------------------------
# End function CreateForm
#---------------------------

#----------------------------------------------------------------------------------------------------------------
#
#                               Functions 
#
#----------------------------------------------------------------------------------------------------------------
#---------------------------
# Create function for check MDT DB if name exists
#---------------------------
Function CheckIfComputerInMDT($ComputerName){
    $result = Get-MDTComputer | Where-Object -Property OSDComputerName -EQ -Value $ComputerName
    if($result -ne $null){
        Return $True
    }
    else{
        Return $False
    }
}
#---------------------------
# Create function for check MDT DB if MAC exists
#---------------------------
Function CheckIfMacAddressInMDTExists($MacAddress){
    $result = Get-MDTComputer -macAddress $MacAddress
    if($result -ne $null){
        Return $True
    }
    else{
        Return $False
    }
}
#---------------------------
# Get File Name for Add action
#---------------------------
function GetFileNameAdd {
    # Open File Dialog
    # Class Details:  https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.openfiledialog 
    $openFileDialog = New-Object windows.forms.openfiledialog   
    $openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()   
    $openFileDialog.title = "Valitse CSV tiedosto prosessoitavaksi"   
    $openFileDialog.filter = "CSV files (*.csv)| *.csv"   
    $openFileDialog.ShowHelp = $True   
    
    $result = $openFileDialog.ShowDialog()   # Display the Dialog / Wait for user response 
    #$result 
    if($result -eq "OK") {    
        $TextBox_Browse_AddCSV.Text  = $OpenFileDialog.filename
        }     
}
#---------------------------
# Get File Name for Remove action
#---------------------------
function GetFileNameRemove {
    # Open File Dialog
    # Class Details:  https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.openfiledialog 
    $openFileDialog = New-Object windows.forms.openfiledialog   
    $openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()   
    $openFileDialog.title = "Valitse CSV tiedosto prosessoitavaksi"   
    $openFileDialog.filter = "CSV files (*.csv)| *.csv"   
    $openFileDialog.ShowHelp = $True   
    
    $result = $openFileDialog.ShowDialog()   # Display the Dialog / Wait for user response 
    #$result 
    if($result -eq "OK") {    
        $TextBox_Browse_RemoveCSV.Text  = $OpenFileDialog.filename
        }     
}
#---------------------------
# Process the CSV file
#---------------------------
function ImportAddCSV {
    $AddCSVfile = $TextBox_Browse_AddCSV.Text
    Write-LogEntry "Käytetään CSV tiedostoa: $AddCSVfile"
    $computers = import-csv $AddCSVfile -Delimiter ","
    ForEach ($computer in $computers) {
        $Addcomputer = $($computer.CName)
        $MacAddress = $($computer.MAC)
        Write-LogEntry "Prosessoidaan konetta: $Addcomputer, jolla MAC-osoite: $MacAddress."
        AddMDTwithCSV $Addcomputer $MacAddress
    }   
}

function ImportRemoveCSV {
    $RemoveCSVfile = $TextBox_Browse_RemoveCSV.Text
    $computers = import-csv $RemoveCSVfile -Delimiter ","
    ForEach ($computer in $computers) {
        $computername = $($computer.CName)
        $RemoveFromAD = $($computer.AD)
        $RemoveFromSCCM = $($computer.SCCM)
        $RemoveFromMDT = $($computer.MDT)
        Write-LogEntry "Prosessoidaan konetta $computername arvoilla AD: $RemoveFromAD, SCCM: $RemoveFromSCCM, MDT: $RemoveFromMDT."
        if ($RemoveFromAD = "1") {
            CleanADwithCSV $computername
        }
        if ($RemoveFromSCCM = "1") {
            CleanSCCMwithCSV $computername
        }
        if ($RemoveFromMDT = "1") {
            CleanMDTwithCSV $computername
        }
    }
    $TextBox_Browse_RemoveCSV.Clear() 
}

#*****************************
# Add ONE computer to MDT DB *
#*****************************
# Note: That you should also fix the database, since it is broken by default, just follow these steps:
# https://syscenramblings.wordpress.com/2016/01/15/mdt-database-the-powershell-module-fix/
function AddMDT {
    # Validate that computer name is between 3 and 15 characters 
    try {
        [ValidateLength(3,15)][string]$AddComputer = $AddComputer
        Write-LogEntry "Koneen nimen pituus on sallituissa (3-15) rajoissa."
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "ERROR: Tarkista koneen nimen pituus on välillä 3-15 merkkiä: $ErrorMessage" -ForegroundColor Red
        Exit 1
    }
    # Validate MAC Address
    try {
        [ValidatePattern('^([0-9a-fA-F]{2}[:-]{0,1}){5}[0-9a-fA-F]{2}$')][string]$MacAddress = $MacAddress
        Write-Host "MAC-osoite näyttää olevan oikean muotoinen."
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "ERROR: Tarkista MAC-osoite: $ErrorMessage" -ForegroundColor Red
        Exit 1
    }
    Write-LogEntry "Valmiina lisäämään $AddComputer MDT-kantaan. Katsotaan onko käyttäjä valmis..."    
    $msgBoxInput= [System.Windows.Forms.MessageBox]::Show("Oletko varma, että haluat lisätä " + $AddComputer + " koneen MDT-kantaan?", "Koneen lisäys MDT-kantaan","YesNo","Exclamation")
    switch ($msgBoxInput) {
        'Yes' {
            #Import the Modules and connect to the database
            Write-LogEntry "Tuodaan MDT moduuli..." 
            try {
                Import-Module "$PSScriptRoot\MDTDB.psm1" -ErrorAction Stop
                Write-LogEntry "MDT moduuli tuonti onnistui" -ForegroundColor Green
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-LogEntry "ERROR: Moduulin tuonissa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
                exit 2 # The system cannot find the file specified.
            }
            Write-LogEntry "Muodostetaan yhteys MDT-kantaan..." 
            if ($MDTinstance -eq "") {
                try {
                    Write-LogEntry "Yhteystiedot = SQL Server: $MDTsqlServer, Database: $MDTdatabase"
                    Connect-MDTDatabase -sqlServer $MDTsqlServer -database $MDTdatabase -ErrorAction stop
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-LogEntry "ERROR: Yhteydenotossa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
                    Exit 1 # Incorrect function.
                }
            }
            else {
                try {
                    Write-LogEntry "Yhteystiedot = SQL Server: $MDTsqlServer, Instance: $MDTinstance, Database: $MDTdatabase"
                    Connect-MDTDatabase -sqlServer $MDTsqlServer -database $MDTdatabase -instance $MDTinstance -ErrorAction stop
                    Write-LogEntry "Yhteys modostettu MDT-kantaan" -ForegroundColor Green
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-LogEntry "ERROR: Yhteydenotossa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
                    Exit 4313 # Unable to read from or write to the database.
                }
            }
            Write-LogEntry "Tarkistetaan löytyykö $AddComputer MDT kannasta..."
            $CheckMDTcName = CheckIfComputerInMDT -ComputerName $AddComputer
            if($CheckMDTcName -eq $false){
                Write-LogEntry "$AddComputer ei löytynyt MDT-kannasta."
                Write-LogEntry "Tarkistetaan löytyykö $MacAddress MDT-kannasta..."
                $CheckMDT_MAC = CheckIfComputerInMDT -macAddress $MacAddress
                if ($CheckMDT_MAC -eq $false) {
                    Write-LogEntry "$MacAddress ei löytynyt MDT-kannasta."
                }
                else {
                    Write-LogEntry "$MacAddress löytyy jo MDT-kannasta. Tarkista MAC-osoite." -ForegroundColor Red
                    Exit 13 # The data is invalid.
                }
                Write-LogEntry "Lisätään $AddComputer MDT-kantaan..." 
                try {
                    New-MDTComputer –macAddress $MacAddress -description $AddComputer –settings @{
                        OSInstall='YES';                       
                        OSDComputerName=$AddComputer}
                    Write-LogEntry "$AddComputer lisätty MDT-kantaan." -ForegroundColor Green
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-LogEntry "ERROR: $AddComputer lusäyksessä ilmeni virhe: $ErrorMessage" -ForegroundColor Red
                }
            }
            else {
                Write-LogEntry "$AddComputer löytyy jo MDT-kannasta. Tarkista koneen nimi." -ForegroundColor Red
                Exit 13 # The data is invalid.
            }          
        }
        'No' {
            Write-LogEntry "WARNING: Käyttäjä perui operaation." -ForegroundColor Yellow
            }
        }
        # Clear the TextBox from the enterd values
        $TextBox_AddOne_CName.Clear()
        $TextBox_AddOne_MAC.Clear()
}


#***********************************
# Add computers with CSV to MDT DB *
#***********************************
function AddMDTwithCSV ($AddComputer,$MacAddress) {
    # Validate that computer name is between 3 and 15 characters 
    try {
        [ValidateLength(3,15)][string]$AddComputer = $AddComputer
        Write-LogEntry "Koneen nimen pituus on sallituissa (3-15) rajoissa."
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "ERROR: Tarkista koneen nimen pituus on välillä 3-15 merkkiä: $ErrorMessage" -ForegroundColor Red
        Exit 1
    }
    # Validate MAC Address
    try {
        [ValidatePattern('^([0-9a-fA-F]{2}[:-]{0,1}){5}[0-9a-fA-F]{2}$')][string]$MacAddress = $MacAddress
        Write-Host "MAC-osoite näyttää olevan oikean muotoinen."
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "ERROR: Tarkista MAC-osoite: $ErrorMessage" -ForegroundColor Red
        Exit 1
    }    
    #Import the Modules and connect to the database
    Write-LogEntry "Tuodaan MDT moduuli..." 
    try {
        Import-Module "$PSScriptRoot\MDTDB.psm1" -ErrorAction Stop
        Write-LogEntry "MDT moduuli tuonti onnistui" -ForegroundColor Green
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-LogEntry "ERROR: Moduulin tuonissa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
        exit 2 # The system cannot find the file specified.
    }
    Write-LogEntry "Muodostetaan yhteys MDT-kantaan..." 
    if ($MDTinstance -eq "") {
        try {
            Write-LogEntry "Yhteystiedot = SQL Server: $MDTsqlServer, Database: $MDTdatabase"
            Connect-MDTDatabase -sqlServer $MDTsqlServer -database $MDTdatabase -ErrorAction stop
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-LogEntry "ERROR: Yhteydenotossa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
            Exit 1 # Incorrect function.
        }
    }
    else {
        try {
            Write-LogEntry "Yhteystiedot = SQL Server: $MDTsqlServer, Instance: $MDTinstance, Database: $MDTdatabase"
            Connect-MDTDatabase -sqlServer $MDTsqlServer -database $MDTdatabase -instance $MDTinstance -ErrorAction stop
            Write-LogEntry "Yhteys modostettu MDT-kantaan" -ForegroundColor Green
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-LogEntry "ERROR: Yhteydenotossa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
            Exit 4313 # Unable to read from or write to the database.
        }
    }
    Write-LogEntry "Tarkistetaan löytyykö $AddComputer MDT kannasta..."
    $CheckMDTcName = CheckIfComputerInMDT -ComputerName $AddComputer
    if($CheckMDTcName -eq $false){
        Write-LogEntry "$AddComputer ei löytynyt MDT-kannasta."
        Write-LogEntry "Tarkistetaan löytyykö $MacAddress MDT-kannasta..."
        $CheckMDT_MAC = CheckIfComputerInMDT -macAddress $MacAddress
        if ($CheckMDT_MAC -eq $false) {
            Write-LogEntry "$MacAddress ei löytynyt MDT-kannasta."
        }
        else {
            Write-LogEntry "$MacAddress löytyy jo MDT-kannasta. Tarkista MAC-osoite." -ForegroundColor Red
            Exit 13 # The data is invalid.
        }
        Write-LogEntry "Lisätään $AddComputer MDT-kantaan..." 
        try {
            New-MDTComputer –macAddress $MacAddress -description $AddComputer –settings @{
                OSInstall='YES';                       
                OSDComputerName=$AddComputer}
            Write-LogEntry "$AddComputer lisätty MDT-kantaan." -ForegroundColor Green
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-LogEntry "ERROR: $AddComputer lusäyksessä ilmeni virhe: $ErrorMessage" -ForegroundColor Red
        }
    }
    else {
        Write-LogEntry "$AddComputer löytyy jo MDT-kannasta. Tarkista koneen nimi." -ForegroundColor Red
        Exit 13 # The data is invalid.
    }
    $TextBox_Browse_AddCSV.Clear()
}
#****************************
# For removing ONE computer *
#****************************
Function CleanAD {
    # Get the computer name form TextBox
    $RemoveComputer = $TextBox_RemoveOne_CName.Text    
    Write-LogEntry "Valmiina poistamaan $RemoveComputer AD:sta. Katsotaan onko käyttäjä valmis..."
    $msgBoxInput= [System.Windows.Forms.MessageBox]::Show("Oletko varma, että haluat poistaa " + $RemoveComputer + " koneen AD:sta?", "Koneen poisto AD:sta","YesNo","Exclamation")
    switch ($msgBoxInput) {
        'Yes' {
            try {
                Remove-ADComputer -Identity $RemoveComputer -Confirm:$False
                Write-LogEntry "$RemoveComputer poistettu AD:sta" -ForegroundColor Green
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-LogEntry "ERROR: Konetta ($RemoveComputer) ei löytynyt AD:sta." -ForegroundColor Red
                Write-LogEntry "ERROR: Virheen tiedot: $ErrorMessage" -ForegroundColor Red
            }
        }
        'No' {
            Write-LogEntry "WARNING: Käyttäjä perui operaation." -ForegroundColor Yellow
        }
    }
    # Clear the TextBox from the enterd value
    $TextBox_RemoveOne_CName.Clear()
}

Function CleanSCCM {
    # Get the computer name form TextBox
    $RemoveComputer = $TextBox_RemoveOne_CName.Text
    Write-LogEntry "Valmiina poistamaan $RemoveComputer SCCM:stä. Katsotaan onko käyttäjä valmis..."   
    $msgBoxInput= [System.Windows.Forms.MessageBox]::Show("Oletko varma, että haluat poistaa " + $RemoveComputer + " koneen SCCM:stä?", "Koneen poisto SCCM:stä","YesNo","Exclamation")
    switch ($msgBoxInput) {
        'Yes' {
            try {
            $comp = Get-WmiObject -cn $SCCMSiteServer -namespace root\sms\site_$($SCCMSiteCode) -class sms_r_system -filter "Name='$($RemoveComputer)'"
            # Delete the computer account 
            $comp.delete()
            Write-LogEntry "$RemoveComputer jolla ResourceID $($comp.ResourceID) on poistettu SCCM:stä." -ForegroundColor Green
            }
            Catch {
                $ErrorMessage = $_.Exception.Message
                Write-LogEntry "ERROR: $RemoveComputer ei löytynyt SCCM:stä." -ForegroundColor red
                Write-LogEntry "ERROR: Virheen tiedot: $ErrorMessage" -ForegroundColor Red
            }
        }
        'No' {
            Write-LogEntry "WARNING: Käyttäjä perui operaation." -ForegroundColor Yellow
            }
        }
        # Clear the TextBox from the enterd value
        $TextBox_RemoveOne_CName.Clear()     
}

Function CleanMDT {
    
    # Get the computer name form TextBox
    $RemoveComputer = $TextBox_RemoveOne_CName.Text 
    Write-LogEntry "Valmiina poistamaan $RemoveComputer MDT:stä. Katsotaan onko käyttäjä valmis..."    
    $msgBoxInput= [System.Windows.Forms.MessageBox]::Show("Oletko varma, että haluat poistaa " + $RemoveComputer + " koneen MDT:stä?", "Koneen poisto MDT:stä","YesNo","Exclamation")
    switch ($msgBoxInput) {
        'Yes' {
            #Import the Modules and connect to the database
            Write-LogEntry "Tuodaan MDT moduuli..." 
            try {
                Import-Module "$PSScriptRoot\MDTDB.psm1" -ErrorAction Stop
                Write-LogEntry "MDT moduuli tuonti onnistui" -ForegroundColor Green
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-LogEntry "ERROR: MDTDB moduulin tuonissa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
                exit 2 # The system cannot find the file specified.
            }
            Write-LogEntry "Muodostetaan yhteys MDT-kantaan..." 
            if ($MDTinstance -eq "") {
                try {
                    Write-LogEntry "Yhteystiedot = SQL Server: $MDTsqlServer, Database: $MDTdatabase"
                    Connect-MDTDatabase -sqlServer $MDTsqlServer -database $MDTdatabase -ErrorAction stop
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-LogEntry "ERROR: Yhteydenotossa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
                    Exit 1 # Incorrect function.
                }
            }
            else {
                try {
                    Write-LogEntry "Yhteystiedot = SQL Server: $MDTsqlServer, Instance: $MDTinstance, Database: $MDTdatabase"
                    Connect-MDTDatabase -sqlServer $MDTsqlServer -database $MDTdatabase -instance $MDTinstance -ErrorAction stop
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-LogEntry "ERROR: Yhteydenotossa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
                    Exit 1 # Incorrect function.
                }
            }
            Write-LogEntry "Tarkistetaan löytyykö $RemoveComputer MDT-kannasta..."
            $CheckMDT = CheckIfComputerInMDT -ComputerName $RemoveComputer
            if($CheckMDT -eq $true){
                Write-LogEntry "$RemoveComputer löytyi MDT kannasta. "
                Write-LogEntry "Poistetaan $RemoveComputer kannasta..." 
                try {
                    Get-MDTComputer -description $RemoveComputer | Remove-MDTComputer
                    Write-LogEntry "$RemoveComputer poistettu MDT-kannasta." -ForegroundColor Green
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-LogEntry "ERROR: $RemoveComputer poistossa ilmeni virhe: $ErrorMessage" -ForegroundColor Red
                    Exit 1 # Incorrect function.
                }
            }
            Else{
                Write-Host "$RemoveComputer konetta ei löydy MDT-kannasta." 
            }              
        }
        'No' {
            Write-LogEntry "WARNING: Käyttäjä perui operaation." -ForegroundColor Yellow
            }
        }
        # Clear the TextBox from the enterd value
        $TextBox_RemoveOne_CName.Clear()
}

#**********************************
# For removing computers with CSV *
#**********************************
Function CleanADwithCSV ($RemoveComputer) {
    
    Write-LogEntry "Poistetaan $RemoveComputer AD:sta." 
    try {
        Remove-ADComputer -Identity $RemoveComputer -Confirm:$False
        Write-LogEntry "$RemoveComputer poistettu AD:sta" -ForegroundColor Green
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-LogEntry "ERROR: Konetta ($RemoveComputer) ei löytynyt AD:sta." -ForegroundColor Red
        Write-LogEntry "ERROR: Virheen tiedot: $ErrorMessage" -ForegroundColor Red
    }
}

Function CleanSCCMwithCSV ($RemoveComputer) {

    Write-LogEntry "Poistetaan $RemoveComputer SCCM:stä."   
    try {
        $comp = Get-WmiObject -cn $SCCMSiteServer -namespace root\sms\site_$($SCCMSiteCode) -class sms_r_system -filter "Name='$($RemoveComputer)'"
        # Delete the computer account 
        $comp.delete()
        Write-LogEntry "$RemoveComputer jolla ResourceID $($comp.ResourceID) on poistettu SCCM:stä." -ForegroundColor Green
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        Write-LogEntry "ERROR: $RemoveComputer ei löytynyt SCCM:stä." -ForegroundColor red
        Write-LogEntry "ERROR: Virheen tiedot: $ErrorMessage" -ForegroundColor Red
    } 
}

Function CleanMDTwithCSV ($RemoveComputer) {
    
    Write-LogEntry "Valmiina poistamaan $RemoveComputer MDT:stä."    
    #Import the Modules and connect to the database
    Write-LogEntry "Tuodaan MDTDB moduuli..." 
    try {
        Import-Module "$PSScriptRoot\MDTDB.psm1" -ErrorAction Stop
        Write-LogEntry "MDTDB moduuli tuonti onnistui" -ForegroundColor Green
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-LogEntry "ERROR: MDTBB moduulin tuonissa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
        exit 2 # The system cannot find the file specified.
    }
    Write-LogEntry "Muodostetaan yhteys MDT-kantaan..." 
    if ($MDTinstance -eq "") {
        try {
            Write-LogEntry "Yhteystiedot = SQL Server: $MDTsqlServer, Database: $MDTdatabase"
            Connect-MDTDatabase -sqlServer $MDTsqlServer -database $MDTdatabase -ErrorAction stop
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-LogEntry "ERROR: Yhteydenotossa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
            Exit 1 # Incorrect function.
        }
    }
    else {
        try {
            Write-LogEntry "Yhteystiedot = SQL Server: $MDTsqlServer, Instance: $MDTinstance, Database: $MDTdatabase"
            Connect-MDTDatabase -sqlServer $MDTsqlServer -database $MDTdatabase -instance $MDTinstance -ErrorAction stop
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-LogEntry "ERROR: Yhteydenotossa tapahtui virhe: $ErrorMessage" -ForegroundColor Red
            Exit 1 # Incorrect function.
        }
    }
    Write-LogEntry "Tarkistetaan löytyykö $RemoveComputer MDT kannasta..."
    $CheckMDT = CheckIfComputerInMDT -ComputerName $RemoveComputer
    if($CheckMDT -eq $true){
        Write-LogEntry "$RemoveComputer löytyy MDT kannasta. "
        Write-LogEntry "Poistetaan $RemoveComputer kannasta..." 
        try {
            Get-MDTComputer -description $RemoveComputer | Remove-MDTComputer
            Write-LogEntry "$RemoveComputer poistettu MDT-kannasta." -ForegroundColor Green
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-LogEntry "ERROR: $RemoveComputer poistossa ilmeni virhe: $ErrorMessage" -ForegroundColor Red
            Exit 1 # Incorrect function.
        }
    }
    Else{
        Write-Host "$RemoveComputer konetta ei löydy MDT-kannasta." 
    }  
} 
#----------------------------------------------------------------------------------------------------------------
#
#                               Main Execution 
#
#----------------------------------------------------------------------------------------------------------------
<# Write-LogEntry "******** Script Started ********" -ForegroundColor cyan
Write-LogEntry "Script Version: $ScriptVersion"
Write-LogEntry "Logging to: $LogFile" #>
Write-LogEntry "******** Skripti käynnistetty ********" -ForegroundColor cyan
Write-LogEntry "Käytetty versio: $ScriptVersion"
Write-LogEntry "Logi: $LogFile"
Write-LogEntry "Suoritetaan käyttäjänä: $env:USERNAME"

#Call the Function
CreateForm

Write-LogEntry "******** Skripti lopetettu ********" -ForegroundColor cyan