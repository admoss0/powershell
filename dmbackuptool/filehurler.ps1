$myPath =  split-path $SCRIPT:MyInvocation.MyCommand.Path -parent # (the path to this powershell script file)
$myOutput = "last_result.html" # (the file to put the results in)
$myXml = "last_result.xml" # (serialisation of $schools object from last data acquisition)

$Excel_file = $myPath + "\schools.xls"
$theFile = $myPath + "\DDS.exe"

$username = ""
$password = ""
$cred = $null
$myDrive = "z:"


# **********************************************************************
function getFreeDrive {  
    68..90 | 
    ForEach-Object { "$([char]$_):" } |   
    Where-Object { 'h:', 'k:', 'z:' -notcontains $_  } |   
    Where-Object {     (new-object System.IO.DriveInfo $_).DriveType -eq 'noRootdirectory'   }
}

# ************************************************************************************************************
function getCredentials {
    $cred = Get-Credential

    if ($cred -eq $null) {exit}
    
    $cred
    
} # end getCredentials
# ************************************************************************************************************

function Excel_load {
    Write-Progress -Activity "Starting Excel" -Status "loading $Excel_file" -PercentComplete (0)
    $xl = New-Object -comobject "Excel.Application"
    $xl.Visible = $false

    $wb = $xl.Workbooks.Open($Excel_file)
    $ws = $wb.Sheets.Item(1)

    $rowcounter = 2
    $rowdata = @()
    $data = $ws.Cells.Item($rowcounter, 3).Value()

    do {
        Write-Progress -Activity $rowcounter -Status "loading  $data" -PercentComplete (0)
        $rowdata += $data
        $rowcounter++
        $data = $ws.Cells.Item($rowcounter, 3).Value()
    }
    while ($data -ne $null)
    
    $rowdata
    $null = $xl.Quit()
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
} # end of Excel_load

# ************************************************************************************************************

function sendData {
    Write-Progress -Activity "Startup" -Status "reading targets from local file ..." -PercentComplete (0)
        $targets = Excel_load
        $count = $targets.Length
        $myDrive = (getFreeDrive)[0]
        $targets | putData

} # end of sendData

# ************************************************************************************************************

function putData { 
    Process {
    Write-Progress -Activity "Testing connectivity" -Status "pinging $_" -PercentComplete ($counter/$count *100)
    if (Test-Connection $_ -Quiet -TimeToLive 10 -Count 1) {
    try{
        $counter++
        $target = "\\" + $_  + "\c$"
        $username = $cred.GetNetworkCredential().Domain + "\" + $cred.GetNetworkCredential().UserName
        $password = $cred.GetNetworkCredential().Password


$osver = Get-WmiObject Win32_OperatingSystem -ComputerName $_ -Credential $cred
$myOs = $osver.Caption

if ($myOs -like "*2008*" ) {
	$targetpath = "\Users\Public\Desktop\"
}
else {
	$targetpath = "\Documents and Settings\All Users\Desktop\"
}

        Write-Progress -Activity "Processing $_" -Status "mapping drive ..." -PercentComplete ($counter/$count *100);
        try {
        $net = new-object -ComObject WScript.Network;
        $net.MapNetworkDrive($myDrive, $target, $false, $username, $password);
        }
        catch [exception] { Write-Progress -Activity $_ -Status "Oh dear, I couldn't map the drive ..." -PercentComplete ($counter/$count *100);
        Start-Sleep -Seconds 2;
        return}
        Set-Location  FILESYSTEM::$myDrive
		
        Write-Progress -Activity "Processing $_" -Status "sending file ..." -PercentComplete ($counter/$count *100);
		
####
Copy-Item $theFile $targetpath
####		
		Write-Progress -Activity "Processing $_" -Status "file sent ..." -PercentComplete ($counter/$count *100);

            
        Set-Location filesystem::c:
        # Start-Sleep -Milliseconds 500
        Write-Progress -Activity "Resetting" -Status "removing mapped drive ..." -PercentComplete ($counter/$count *100)
        $null = net use $myDrive /DELETE /y
  
        }
     catch [system.exception] {echo $error} 
    }
    }
} # end of putData

# ************************************************************************************************************

function getCredentials {
    $cred = Get-Credential

    if ($cred -eq $null) {exit}
    
    $cred
    
} # end getCredentials

# ************************************************************************************************************


$cred = getCredentials
sendData
