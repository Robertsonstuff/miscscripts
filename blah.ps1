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
