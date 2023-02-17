#needs computer name input and spits out: computername, last logon date and username
  [cmdletBinding()]
param(
	[Parameter(Mandatory=$True)]
	[String[]]$ComputerName
	
)

Get-ADComputer -Identity "$ComputerName" -Properties lastLogonDate, extensionAttribute10 | Select SAMAccountName, lastLogonDate, extensionAttribute10


	# tells you what OS update your are on. enter-pssession then run below. This acts like a 'winver'.
(Get-WmiObject -class Win32_OperatingSystem).Caption +" "+ (Get-WmiObject -class Win32_OperatingSystem).Version + "." + (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR).UBR + " "+ (Get-WmiObject -Class Win32_ComputerSystem).SystemType + " (" + (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId + ")"


	#this checks how a laptop built in 1809 with heat, bitlocker and snapdragon drivers.

#Create text file in the location below C:\path\textfile.txt

$Computers101 = Get-Content C:\path\textfile.txt

#depending on what command you want to run, delete the '#' symbol

forEach ($computers in $Computers101)
{
  #Check bitlocker
  Write-Output "`r`nCheck bitlocker status and encryption percentage:`r`n`r`n"
  Invoke-Command -ComputerName $Computers -ScriptBlock {bitlockervolume}
  
  #Check heat - if TRUE - it will have all 3 modules
  Write-Output "`r`nCheck HEAT status is TRUE:`r`n`r`n"
  Invoke-Command -ComputerName $Computers -ScriptBlock {(get-content 'C:\Path\heat.xml).Contains("<components pending=`"false`">")}

  #Check simcard - if this comes up with results - it has what we need.
  Write-Output "`r`nCheck result for sim card functionality:`r`n`r`n"
  Invoke-Command -ComputerName $Computers -ScriptBlock {Get-WmiObject win32_pnpentity | Where-Object {$_.name -like "DW5821*"} | select description}

}


Write-Output "`r`nCheck results and amend your computer list accordingly`r`n`r`n`rYou have 10 secs to do this`r`n"

start-sleep -Seconds 10


forEach ($computers in $Computers101) {

  #Start bitlocker
  Invoke-Command -ComputerName $Computers -ScriptBlock {manage-bde -on C:}

}

Write-Output "`r`nCheck results and amend your computer list accordingly`r`n`r`n`rYou have 10 secs to do this`r`n"

start-sleep -Seconds 10


forEach ($computers in $Computers101) {

  #Restart computer to push bitlocker ON
  Invoke-Command -ComputerName $Computers -ScriptBlock {restart-computer}

}

Write-Output "`r`nCheck results and amend your computer list accordingly`r`n`r`n`rYou have 120 seconds to do this`r`n"

start-sleep -Seconds 120

forEach ($computers in $Computers101) {

  #Restart computer to push bitlocker ON
  Invoke-Command -ComputerName $Computers -ScriptBlock {bitlockervolume}

}

	#fetches a website
start-process "https://www.google.com/"

	# battery info
Get-wmiobject -class win32_battery

	# bios info
Get-wmiobject -class win32_bios

	# Tells you if bitlocker has ecrypted and percentage
bitlockervolume

	# suspend bitlocker for two reboots 
Suspend-bitlocker -mountpoint c: -rebootcount 2

	# encrypts bitlocker
manage-bde -on C: 

restart-computer -force

	# shutdown and restart computer in 120 seconds with "…" message.
shutdown /r /t 120 /c "This is a test on computer" 

	# Opening a file in notepad
notepad $a

copy-item -path file.py -Destination file2.py

rename-item -path C:\Scripts\blah -newname ChristmasForm.py

	# get computer name
	$a = get-computerinfo
	$a.csname

    # remote to a computer
enter-pssession -computername blah
Exit-pssession 

#set parameters for a script. Run script followed by the parameters you want to add, separated by a space.
Param (
    [Parameter()]
    [String]$parameter1,

    [Parameter()]
    [String]$parameter2
)
write-output "write this out and add $parameter1"
write-output "and after that, I would like to add $parameter2"



write-output "I'd like this specific text in my clipboard." | Set-clipboard

    # checks current computer for RAM - one line
$physicalram = (Get-WMIObject -class Win32_PhysicalMemory -ComputerName blah | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)})
    # checks multiple computers for RAM and processor info
    # the comments show you a different way to write it.
$Computers101 = Get-Content C:\Path\list.txt

$cap = get-wmiobject -class win32_physicalmemory

$getcompinfo = Get-ComputerInfo


forEach ($computers in $Computers101)
{
  Write-Output "`r`nChecking CPU specs:`r`n"
  $cap = Get-WmiObject -Class win32_PhysicalMemory -ComputerName $computers
  Write-Host "$($getcomputerinfo.name)"
 # Invoke-Command -ComputerName $Computers -ScriptBlock {write-host $getcompinfo.CsProcessors.name}
  Write-Output "`r`nChecking RAM specs:`r`n"
  $getcomputerinfo = Get-WmiObject -Class Win32_Processor -ComputerName $computers
  write-host "$($cap.capacity / 1GB)"

 # Invoke-Command -ComputerName $Computers -ScriptBlock {write-host ($cap.capacity/ 1GB)}
 # if (Test-Connection -ComputerName $computers -Count 1)
 # {
 # $cap = Get-WmiObject -Class win32_PhysicalMemory -ComputerName $computers
 # Write-Host "Physical Memory: $($cap.Capacity / 1GB)"

 # $getcomputerinfo = Get-WmiObject -Class Win32_Processor -ComputerName $computers
 # Write-Host "CPU: $($getcomputerinfo.name)"

 # }
 # else
 # {
 # Write-Warning "Unable to connect to computer"
 # }
}

# reading and reporting data from a excel spreadsheet

$excel = New-Object -Com Excel.Application
$wb = $excel.Workbooks.Open("C:\Users\lrobertson\Desktop\local-dispatch.xlsx")

$wb.sheets.item(1).activate()
$WbTotal=$wb.Worksheets.item(1)
#$Value = $wbTotal.Cells.Item(13,7)
#$Value.Text
#This will return 110. It searches 13 cells down first and then 7 cells to the right

$SearchString = read-host "company name please"

$Range = $WbTotal.Range("A1").EntireColumn
$Search = $Range.find($SearchString)
$Range2 = $WbTotal.Range("A1").EntireColumn
$Search2 = $Range2.find("Company")

#$Search2.EntireRow.Value2
$Search3 = $Search2.EntireRow.Value2
#$Search[1,4]

#$Search.EntireRow.Value2

$Search1 = $Search.EntireRow.Value2

For ($i = 4; $i -lt 17; $i++) {
    write-host "$($Search1[1,$($i)]) minutes for: $($Search3[1,$($i)])"
}

#Form1
#Version 1.4
    [reflection.assembly]::LoadwithPartialName("System.windows.Forms") | Out-Null

    $basicForm = New-Object System.windows.Forms.Form
    $folderForm = New-Object System.Windows.Forms.Form


    $folderForm.Text ="Luke's CI Form"
    $folderForm.width = 300
    $folderForm.Height = 470
    $folderForm.AutoSize = $True

    $label = New-Object System.Windows.Forms.Label
    $label.Name = "Status"
    $label.Location = '20,15'
    $label.Size = '70,15'
    $label.Text = "Status"
    $folderForm.Controls.Add($label)

    $pathTextBox1 = New-Object System.Windows.Forms.TextBox
    $pathTextBox1.Location = '20,30'
    $pathTextBox1.Size = '100,20'
    $folderForm.Controls.Add($pathTextBox1)

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Name = "FriendlyName"
    $label2.Location = '20,55'
    $label2.Size = '200,15'
    $label2.Text = "Friendly Name, Alt CI, Serial Number"
    $folderForm.Controls.Add($label2)

    $pathTextBox2 = New-Object System.Windows.Forms.TextBox
    $pathTextBox2.Location = '20,70'
    $pathTextBox2.Size = '100,20'
    $folderForm.Controls.Add($pathTextBox2)

    $label3 = New-Object System.Windows.Forms.Label
    $label3.Name = "Asset"
    $label3.Location = '20,95'
    $label3.Size = '200,15'
    $label3.Text = "Asset Number / DOH Number"
    $folderForm.Controls.Add($label3)

    $pathTextBox3 = New-Object System.Windows.Forms.TextBox
    $pathTextBox3.Location = '20,110'
    $pathTextBox3.Size = '100,20'
    $folderForm.Controls.Add($pathTextBox3)

    $label4 = New-Object System.Windows.Forms.Label
    $label4.Name = "CI SubType"
    $label4.Location = '20,135'
    $label4.Size = '200,15'
    $label4.Text = "CI SubType"
    $folderForm.Controls.Add($label4)

    $pathTextBox4 = New-Object System.Windows.Forms.TextBox
    $pathTextBox4.Location = '20,150'
    $pathTextBox4.Size = '100,20'
    $folderForm.Controls.Add($pathTextBox4)

    $label5 = New-Object System.Windows.Forms.Label
    $label5.Name = "Model"
    $label5.Location = '20,175'
    $label5.Size = '200,15'
    $label5.Text = "Model"
    $folderForm.Controls.Add($label5)

    $pathTextBox5 = New-Object System.Windows.Forms.TextBox
    $pathTextBox5.Location = '20,190'
    $pathTextBox5.Size = '100,20'
    $folderForm.Controls.Add($pathTextBox5)

    $label6 = New-Object System.Windows.Forms.Label
    $label6.Name = "Primary Contact"
    $label6.Location = '20,215'
    $label6.Size = '200,15'
    $label6.Text = "Primary Contact"
    $folderForm.Controls.Add($label6)

    $pathTextBox6 = New-Object System.Windows.Forms.TextBox
    $pathTextBox6.Location = '20,230'
    $pathTextBox6.Size = '150,20'
    $folderForm.Controls.Add($pathTextBox6)

    $label7 = New-Object System.Windows.Forms.Label
    $label7.Name = "Building Location"
    $label7.Location = '20,255'
    $label7.Size = '200,15'
    $label7.Text = "Building Location"
    $folderForm.Controls.Add($label7)

    $pathTextBox7 = New-Object System.Windows.Forms.TextBox
    $pathTextBox7.Location = '20,270'
    $pathTextBox7.Size = '150,20'
    $folderForm.Controls.Add($pathTextBox7)

    $label8 = New-Object System.Windows.Forms.Label
    $label8.Name = "Room Location"
    $label8.Location = '20,295'
    $label8.Size = '200,15'
    $label8.Text = "Room Location"
    $folderForm.Controls.Add($label8)

    $pathTextBox8 = New-Object System.Windows.Forms.TextBox
    $pathTextBox8.Location = '20,310'
    $pathTextBox8.Size = '50,20'
    $folderForm.Controls.Add($pathTextBox8)

    

    $selectButton = New-Object System.Windows.Forms.Button
    $selectButton.Text = 'Register'
    $selectButton.Location = '20,350'
    $selectButton.Size = '170,50'

    $folderForm.Controls.Add($selectButton)

    #Add Button event 
    $SelectButton.Add_Click(
        {    
        [reflection.assembly]::LoadwithPartialName("System.windows.Forms") | Out-Null

            if ((-not [string]::IsNullOrEmpty($pathTextBox2.text)) -and [string]::IsNullOrEmpty($pathTextBox4.text)) {
 	    
        $basicForm21 = New-Object System.windows.Forms.Form
        $folderForm21 = New-Object System.Windows.Forms.Form

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

 $folderForm21.Text ="!"
        $folderForm21.width = 190
        $folderForm21.Height = 100
        $folderForm21.AutoSize = $True

        $label21 = New-Object System.Windows.Forms.Label
        $label21.Name = "Status3"
        $label21.Location = '20,15'
        $label21.Size = '150,15'
        $label21.Text = "Please Put in CI SubType"
        $folderForm21.Controls.Add($label21)
        $folderForm21.ShowDialog()
}
else {
        $basicForm20 = New-Object System.windows.Forms.Form
        $folderForm20 = New-Object System.Windows.Forms.Form

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing


        $folderForm20.Text ="CI Output"
        $folderForm20.width = 300
        $folderForm20.Height = 325
        $folderForm20.AutoSize = $True

        $label20 = New-Object System.Windows.Forms.Label
        $label20.Name = "Status2"
        $label20.Location = '20,15'
        $label20.Size = '70,15'
        $label20.Text = "Output:"
        $folderForm20.Controls.Add($label20)

	$outputBox20 = New-Object System.Windows.Forms.TextBox
        $outputBox20.Location = '20,32'
        $outputBox20.Size = '240,225'
        $outputBox20.Multiline = $True
        $folderForm20.Controls.Add($outputBox20)

        $outputBox20.AppendText("-----------These items below have been updated / created in the CMDB-----------")
	if ([string]::IsNullOrEmpty($pathTextBox1.text)) {
	$outputBox20.AppendText("`r`n`r`n")
}
else {
	$outputBox20.AppendText("`r`n`r`nStatus: ")
	$outputBox20.AppendText($pathTextBox1.text)
}
	if ([string]::IsNullOrEmpty($pathTextBox2.text)) {
}
else {
	$outputBox20.AppendText("`r`nFriendly Name: ")
	$outputBox20.AppendText($pathTextBox2.text)
	if ($pathTextBox4 -match "Laptop") {
	$outputBox20.AppendText("`r`nAlt CI: UCLP")
	$outputBox20.AppendText($pathTextBox2.text)
}
elseif ($pathTextBox4 -match "Desktop") {
	$outputBox20.AppendText("`r`nAlt CI: UCDP")
	$outputBox20.AppendText($pathTextBox2.text)
}
else {
	$outputBox20.AppendText("`r`nAlt CI: ")
	$outputBox20.AppendText($pathTextBox2.text)
}
	$outputBox20.AppendText("`r`nSerial Number: ")
	$outputBox20.AppendText($pathTextBox2.text)
}
	if ([string]::IsNullOrEmpty($pathTextBox3.text)) {
}
else {
	$outputBox20.AppendText("`r`nAsset Number: ")
	$outputBox20.AppendText($pathTextBox3.text)
}
	if ([string]::IsNullOrEmpty($pathTextBox4.text)) {
}
else {
	$outputBox20.AppendText("`r`nCI SubType: ")
	$outputBox20.AppendText($pathTextBox4.text)
}
	if ([string]::IsNullOrEmpty($pathTextBox5.text)) {
}
else {
	$outputBox20.AppendText("`r`nModel: ")
	$outputBox20.AppendText($pathTextBox5.text)
}
	if ([string]::IsNullOrEmpty($pathTextBox6.text)) {
}
else {
	$outputBox20.AppendText("`r`nPrimary Contact: ")
	$outputBox20.AppendText($pathTextBox6.text)
}
	if ([string]::IsNullOrEmpty($pathTextBox7.text)) {
}
else {
	$outputBox20.AppendText("`r`nBuilding Location: ")
	$outputBox20.AppendText($pathTextBox7.text)
}
	if ([string]::IsNullOrEmpty($pathTextBox8.text)) {
}
else {
	$outputBox20.AppendText("`r`nRoom Location: ")
	$outputBox20.AppendText($pathTextBox8.text)
}

	$StatusBar = New-Object System.Windows.Forms.StatusBar
    	$StatusBar.Text = "Copied To Clipboard"
    	$StatusBar.Height = 22
    	$StatusBar.Width = 200
    	$StatusBar.Location = New-Object System.Drawing.Point( 0, 250 )
    	$folderForm20.Controls.Add($StatusBar)
    
        $copyText = $outputBox20.Text.Trim()

        [System.Windows.Forms.Clipboard]::SetText($copyText)

        if ([System.Windows.Forms.Clipboard]::ContainsText() -AND
            [System.Windows.Forms.Clipboard]::GetText() -eq $copyText)
	{
	Write-Progress -Activity Updating -Status 'Copied To Clipboard' 
    	}

        $folderForm20.ShowDialog()
 } }
    )


$folderForm.ShowDialog()

#Form 2

#Version 1.5 - add buttons - 
    [reflection.assembly]::LoadwithPartialName("System.windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    $itemlist = @("Latitude 7480","Latitude 7490","Latitude 7400","Latitude 7410","Latitude 7420","Latitude 5400","Latitude 7390","Latitude 3410")
    $itemlist2 = @("Sirius Building","Scarborough House")
    $itemlist3 = @("In Stock","In Service","In Repair","On Loan")	
    $basicForm = New-Object System.windows.Forms.Form
    $folderForm = New-Object System.Windows.Forms.Form
    $comboBox1 = New-Object System.Windows.Forms.ComboBox
    $comboBox2 = New-Object System.Windows.Forms.ComboBox
    $comboBox3 = New-Object System.Windows.Forms.ComboBox
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    $folderForm.Text ="Luke's CI Form"
    $folderForm.width = 300
    $folderForm.Height = 430
    $folderForm.AutoSize = $True

    $label = New-Object System.Windows.Forms.Label
    $label.Name = "Status"
    $label.Location = '20,15'
    $label.Size = '70,15'
    $label.Text = "Status"
    $folderForm.Controls.Add($label)

    $comboBox3.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBox3.FormattingEnabled = $True
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 20
    $System_Drawing_Point.Y = 30
    $comboBox3.Location = $System_Drawing_Point
    $comboBox3.Name = "comboBox3"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 100
    $comboBox3.Size = $System_Drawing_Size
    $comboBox3.TabIndex = 2
    $comboBox3.Text = "Status"
    $folderForm.Controls.Add($comboBox3)
    

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Name = "FriendlyName"
    $label2.Location = '20,55'
    $label2.Size = '200,15'
    $label2.Text = "Friendly Name, Alt CI, Serial Number"
    $folderForm.Controls.Add($label2)

    $pathTextBox2 = New-Object System.Windows.Forms.TextBox
    $pathTextBox2.Location = '20,70'
    $pathTextBox2.Size = '100,20'
    $folderForm.Controls.Add($pathTextBox2)

    $label3 = New-Object System.Windows.Forms.Label
    $label3.Name = "Asset"
    $label3.Location = '20,95'
    $label3.Size = '200,15'
    $label3.Text = "Asset Number / DOH Number"
    $folderForm.Controls.Add($label3)

    $pathTextBox3 = New-Object System.Windows.Forms.TextBox
    $pathTextBox3.Location = '20,110'
    $pathTextBox3.Size = '100,20'
    $folderForm.Controls.Add($pathTextBox3)

    $label5 = New-Object System.Windows.Forms.Label
    $label5.Name = "Model"
    $label5.Location = '20,135'
    $label5.Size = '200,15'
    $label5.Text = "Model"
    $folderForm.Controls.Add($label5)

    $comboBox1.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBox1.FormattingEnabled = $True
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 20
    $System_Drawing_Point.Y = 150
    $comboBox1.Location = $System_Drawing_Point
    $comboBox1.Name = "comboBox1"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 150
    $comboBox1.Size = $System_Drawing_Size
    $comboBox1.TabIndex = 2
    $comboBox1.Text = "Model"
    $folderForm.Controls.Add($comboBox1)

    $label6 = New-Object System.Windows.Forms.Label
    $label6.Name = "Primary Contact"
    $label6.Location = '20,175'
    $label6.Size = '200,15'
    $label6.Text = "Primary Contact"
    $folderForm.Controls.Add($label6)

    $pathTextBox6 = New-Object System.Windows.Forms.TextBox
    $pathTextBox6.Location = '20,190'
    $pathTextBox6.Size = '150,20'
    $folderForm.Controls.Add($pathTextBox6)

    $label7 = New-Object System.Windows.Forms.Label
    $label7.Name = "Building Location"
    $label7.Location = '20,215'
    $label7.Size = '200,15'
    $label7.Text = "Building Location"
    $folderForm.Controls.Add($label7)
    
    $comboBox2.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboBox2.FormattingEnabled = $True
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 20
    $System_Drawing_Point.Y = 230
    $comboBox2.Location = $System_Drawing_Point
    $comboBox2.Name = "comboBox2"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 20
    $System_Drawing_Size.Width = 150
    $comboBox2.Size = $System_Drawing_Size
    $comboBox2.TabIndex = 2
    $comboBox2.Text = "Building"
    $folderForm.Controls.Add($comboBox2)



    #$pathTextBox7 = New-Object System.Windows.Forms.TextBox
    #$pathTextBox7.Location = '20,255'
    #$pathTextBox7.Size = '150,20'
    #$folderForm.Controls.Add($pathTextBox7)

    $label8 = New-Object System.Windows.Forms.Label
    $label8.Name = "Room Location"
    $label8.Location = '20,255'
    $label8.Size = '200,15'
    $label8.Text = "Room Location"
    $folderForm.Controls.Add($label8)

    $pathTextBox8 = New-Object System.Windows.Forms.TextBox
    $pathTextBox8.Location = '20,270'
    $pathTextBox8.Size = '50,20'
    $folderForm.Controls.Add($pathTextBox8)
 
    $selectButton = New-Object System.Windows.Forms.Button
    $selectButton.Text = 'Register'
    $selectButton.Location = '20,320'
    $selectButton.Size = '170,50'

    $folderForm.Controls.Add($SelectButton)

    foreach ($i in $itemlist)
{
    $comboBox1.items.Add($i)
}

    foreach ($i in $itemlist2)
{
     $comboBox2.items.Add($i)
}

    foreach ($i in $itemlist3)
{
     $comboBox3.items.Add($i)
}

#$SelectButton.Add_Click(
#{
#    $combobox1.items.Clear()
#   get-content U:\desktop\list.txt | % {

#   $comboBox1.items.add($_)

#   } 
#}
#)



    #Add Button event 
    $SelectButton.Add_Click(
        {    
        [reflection.assembly]::LoadwithPartialName("System.windows.Forms") | Out-Null

        #$Combobox1.items.Clear()
        #get-content U:\desktop\list.txt | % {

        #$comboBox1.items.add($_)
        #}

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $basicForm20 = New-Object System.windows.Forms.Form
        $folderForm20 = New-Object System.Windows.Forms.Form

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $folderForm20.Text ="CI Output"
        $folderForm20.width = 300
        $folderForm20.Height = 325
        $folderForm20.AutoSize = $True

        $label20 = New-Object System.Windows.Forms.Label
        $label20.Name = "Status2"
        $label20.Location = '20,15'
        $label20.Size = '70,15'
        $label20.Text = "Output:"
        $folderForm20.Controls.Add($label20)

	$outputBox20 = New-Object System.Windows.Forms.TextBox
        $outputBox20.Location = '20,32'
        $outputBox20.Size = '240,225'
        $outputBox20.Multiline = $True
        $folderForm20.Controls.Add($outputBox20)

        $outputBox20.AppendText("-----------These items below have been updated / created in the CMDB-----------")
	
	$outputBox20.AppendText("`r`n`r`n")

	$outputBox20.AppendText("Status: ")
	$outputBox20.AppendText($comboBox3.text)

	$outputBox20.AppendText("`r`nAlt CI: UCLP")
	$outputBox20.AppendText($pathTextBox2.text)

	$outputBox20.AppendText("`r`nAsset Number: ")
	$outputBox20.AppendText($pathTextBox3.text)

	$outputBox20.AppendText("`r`nCI SubType: Laptop")
	
	$outputBox20.AppendText("`r`nModel: ")
	$outputBox20.AppendText($comboBox1.text)
	
	$outputBox20.AppendText("`r`nPrimary Contact: ")
	$outputBox20.AppendText($pathTextBox6.text)
	
	$outputBox20.AppendText("`r`nBuilding Location: ")
	$outputBox20.AppendText($comboBox2.text)


	$outputBox20.AppendText("`r`nRoom Location: ")
	$outputBox20.AppendText($pathTextBox8.text)

        $outputBox20.AppendText("`r`nSerial Number: ")
	$outputBox20.AppendText($pathTextBox2.text)


	$StatusBar = New-Object System.Windows.Forms.StatusBar
    	$StatusBar.Text = "Copied To Clipboard"
    	$StatusBar.Height = 22
    	$StatusBar.Width = 200
    	$StatusBar.Location = New-Object System.Drawing.Point( 0, 250 )
    	$folderForm20.Controls.Add($StatusBar)
    
        $copyText = $outputBox20.Text.Trim()

        [System.Windows.Forms.Clipboard]::SetText($copyText)

        if ([System.Windows.Forms.Clipboard]::ContainsText() -AND
            [System.Windows.Forms.Clipboard]::GetText() -eq $copyText)
	{
	Write-Progress -Activity Updating -Status 'Copied To Clipboard' 
    	}
	$comboBox3.SelectedIndex =-1
	$pathTextBox2.Clear()
	$pathTextBox3.Clear()
	$pathTextBox6.Clear()
	$pathTextBox8.Clear()
	$comboBox1.SelectedIndex =-1
	$comboBox2.SelectedIndex =-1
        $folderForm20.ShowDialog()
  }
    )


$folderForm.ShowDialog()

# calendar access in exchange
$a = read-host mailbox email ?
$b = read-host userID needing access? 
$c = read-host editor or reviewer?
$UserCredential = Get-Credential
read-host are you logged in?
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://blah -Authentication Kerberos -Credential $UserCredential
Start-Sleep -s 2
Import-PSSession $Session -DisableNameChecking
Start-Sleep -s 2
$e = (("Add-MailboxFolderPermission $a") + (":\Calendar -User $b -AccessRights $c"))
write-output $e | set-clipboard
write-host "`r`nYour script has been saved to your clipboard. Paste and run`r`n"

#os checker
$a = read-host computer name?
start-sleep -s 2
invoke-command -ComputerName "$a" -ScriptBlock {(Get-WmiObject -class Win32_OperatingSystem).Caption +" "+ (Get-WmiObject -class Win32_OperatingSystem).Version + "." + (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR).UBR + " "+ (Get-WmiObject -Class Win32_ComputerSystem).SystemType + " (" + (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId + ")"}

#group policy update
$a = read-host computer name?
start-sleep -s 2
invoke-command -ComputerName "$a" -ScriptBlock {gpupdate /force}

#stress test
$NumberOfLogicalProcessors = Get-WmiObject win32_processor | Select-Object -ExpandProperty NumberOfLogicalProcessors

ForEach ($core in 1..$numberOfLogicalProcessors){

start-job -ScriptBlock{

    $result = 1;
    foreach ($loopnumber in 1..2147483647){
        $result=1;

        foreach ($loopnumber1 in 1..2147483647){
        $result=1;
            
            foreach($number in 1..2147483647){
                $result = $result * $number
	    }
	}

	    $result
	}
    }
}

Read-Host "Press any key to exit..."
Stop-Job *
}
