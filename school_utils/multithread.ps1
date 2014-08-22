
clear
Get-Job | Remove-Job -Force
$cred = Get-Credential

$myPath =  split-path $SCRIPT:MyInvocation.MyCommand.Path -parent # (the path to this powershell script file)
cd $myPath

$schools = Import-Csv -Path $myPath/schools.txt -Delimiter `t


    $MaxThreads = 20
    $SleepTimer = 50
$i = 0
$Computers = $schools 
 
ForEach ($Computer in $Computers){
    # Check to see if there are too many open threads
    # If there are too many threads then wait here until some close
    While ($(Get-Job -state running).count -ge $MaxThreads){
        Write-Progress -Activity "Creating Server List" -Status "Waiting for threads to close" -CurrentOperation "$i threads created - $($(Get-Job -state running).count) threads open" -PercentComplete ($i / $Computers.count * 100)
        Start-Sleep -Milliseconds $SleepTimer
    }
 
    #"Starting job - $Computer"
    $i++
    #Start-Job  {param($Computer, $file, $cred); $tag = (Get-WmiObject Win32_BIOS -computerName $Computer -Credential $cred).SerialNumber ; $tag } -ArgumentList $Computer, $file, $cred
    Start-Job  {param($Computer, $file, $cred); $tag = $Computer.IP ; $tag } -ArgumentList $Computer, $file, $cred | Out-Null
    
	Write-Progress  -Activity "Creating Server List" -Status "Starting Threads" -CurrentOperation "$i threads created - $($(Get-Job -state running).count) threads open" -PercentComplete ($i / $Computers.count * 100)
	}
Get-Job | Wait-Job
Get-Job | Receive-Job 