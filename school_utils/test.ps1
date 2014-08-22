
clear

function Get-OS { 
  Param([string]$computername=$(Throw "You must specify a computername.")) 
  Write-Host $computername
  Write-Debug "In Get-OS Function" 
  $wmi=Get-WmiObject Win32_OperatingSystem -computername $computername -Credential $cred -ea stop 
   
  write $wmi.Caption 
  
  }



$myPath =  split-path $SCRIPT:MyInvocation.MyCommand.Path -parent # (the path to this powershell script file)
cd $myPath

$schools = Import-Csv -Path $myPath/schools.txt -Delimiter `t
#$cred = Get-Credential



#foreach ($school in $schools) {
#$OS = Get-OS ($school.IP)
#Write-Host $school.School `t  $OS
#}

$file = "d:\servicetags.txt"
"Code,Name,ServiceTag" | Out-File $file
foreach ($school in $schools) {
$tag = (Get-WmiObject Win32_BIOS -computerName $school.IP -Credential $cred).SerialNumber
$school.Code + ',' + $school.School + ',' + $tag  | Write-Output
$school.Code + ',' + $school.School + ',' + $tag | Out-File $file -Append
}

