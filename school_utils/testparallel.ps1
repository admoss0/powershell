."H:\powershell\school_utils\Invoke-Parallel.ps1"
clear
$file = "d:\servicetags.txt" 
"Name,ServiceTag" | Out-File $file	

$myPath =  split-path $SCRIPT:MyInvocation.MyCommand.Path -parent # (the path to this powershell script file)
cd $myPath



$schools = Import-Csv -Path "d:\allserversindds.txt" 
$cred = Get-Credential

$servers = $schools | Select-Object Name

$info = New-Object -TypeName psobject -Property @{ 
            cred = $cred
			outfile = $file		
			}

Invoke-Parallel -InputObject $servers -parameter $info -Throttle 30 -runspaceTimeout 30 -ScriptBlock {
$thing = $_.Name + '.dds.eq.edu.au';
$tag = (Get-WmiObject Win32_BIOS -computerName $thing -Credential $parameter.cred).SerialNumber;
 if ($tag.Length -eq 7) { # Dell Service tags are 7 characters long
 $_.Name +  ',' + $tag | Out-File $parameter.outfile -Append
}
}



#(Get-WmiObject Win32_BIOS -computerName $_ + ".dds.eq.edu.au" -Credential $parameter.cred).SerialNumber;