
$myPath =  split-path $SCRIPT:MyInvocation.MyCommand.Path -parent # (the path to this powershell script file)
$myOutput = "last_result.html" # (the file to put the results in)
$myXml = "last_result.xml" # (serialisation of $schools object from last data acquisition)


$username = ""
$password = ""
$cred = $null

$myDrive = "z:"

# ************************************************************************************************************

$reportfooter = @"
<HR />
<SMALL> David Moss's Backup Report Tool is written in Powershell V2.0</SMALL>
"@

# ************************************************************************************************************


function getFreeDrive {  
    68..90 | 
    ForEach-Object { "$([char]$_):" } |   
    Where-Object { 'h:', 'k:', 'z:' -notcontains $_  } |   
    Where-Object {     (new-object System.IO.DriveInfo $_).DriveType -eq 'noRootdirectory'   }
}

# ************************************************************************************************************

$Excel_file = $myPath + "\schools.xls"

#Zero some variables
$school = ""
$schools = ""
$schools = @()
$counter = 0
$count = 0

# ************************************************************************************************************

#Define a custom object
Add-Type -TypeDefinition @"
public class MyResultSet{
public int id;
public string status, description;
}
"@

#Define a custom object
Add-Type -TypeDefinition @"
public class MySchool{
public string school_name;
public string last_backup;
public object[] backup_results = new object[0];
}
"@

# ************************************************************************************************************


# gets school data for addresses piped in from an array
# returns an array of MySchool objects
function getSchoolData { 
    Process {
    Write-Progress -Activity "Testing connectivity" -Status "pinging $_" -PercentComplete ($counter/$count *100)
    if (Test-Connection $_ -Quiet -TimeToLive 10 -Count 1) {
    try{
        $counter++
        $target = "\\" + $_  + "\log$"
        $username = $cred.GetNetworkCredential().Domain + "\" + $cred.GetNetworkCredential().UserName
        $password = $cred.GetNetworkCredential().Password

        Write-Progress -Activity "Processing $_" -Status "mapping drive ..." -PercentComplete ($counter/$count *100);
        try {
        
        # $null = net use $myDrive $target\log$ $password /USER:$username }
        
        $net = new-object -ComObject WScript.Network;
        $net.MapNetworkDrive($myDrive, $target, $false, $username, $password);
        # $net.MapNetworkDrive($myDrive, $target, $false, "dds\st-admos0", $password)
        }

        
        
        catch [exception] { Write-Progress -Activity $_ -Status "Oh dear, I couldn't map the drive ..." -PercentComplete ($counter/$count *100);
        Start-Sleep -Seconds 2;
        return}
        Write-Progress -Activity "Processing $_" -Status "checking for file ..." -PercentComplete ($counter/$count *100);
        Set-Location  FILESYSTEM::$myDrive
        
        
        if (Test-Path -Path $myDrive\check.txt ) {
            Write-Progress -Activity "Processing $_" -Status "downloading file ..." -PercentComplete ($counter/$count *100);
            $temp = get-content check.txt;
                Write-Progress -Activity "Processing $_" -Status "file downloaded ..." -PercentComplete ($counter/$count *100);
            }
            else {Write-Progress -Activity "Error on $_" -Status "Oh dear - no file to get ..." -PercentComplete ($counter/$count *100);
            Set-Location filesystem::c: ;
             $null = net use $myDrive /DELETE /y;
             Start-Sleep -Milliseconds 2000;
             return}
            
        Set-Location filesystem::c:
        Start-Sleep -Milliseconds 500
        Write-Progress -Activity "Processing $_" -Status "removing mapped drive ..." -PercentComplete ($counter/$count *100)
        $null = net use $myDrive /DELETE /y
     
        
        Write-Progress -Activity "Processing $_" -Status "Parsing file ..." -PercentComplete ($counter/$count *100)
        $school = New-Object MySchool
        $school.school_name = ($temp[0] -split " \*\*\* Backup Summary \*\*\*")[0]
        $school.last_backup = $temp | Select-String "^[0-9]{2}/"
        $temp | Select-String -pattern "^\s*[0-9]{1,2}\." |
        ForEach-Object {$result = New-Object MyResultSet ; 
        $result.id, $result.status, $result.description = $_ -split "!" ; 
        $school.backup_results += $result }
        Write-Progress -Activity "Processing $_" -Status "Adding results ..." -PercentComplete ($counter/$count *100)
        
        return $school
        
        }
     catch [system.exception] {echo $error} 
    }
    }
} # end of getSchoolData

# ************************************************************************************************************

function getCredentials {
    $cred = Get-Credential

    if ($cred -eq $null) {exit}
    
    $cred
    
} # end getCredentials

# ************************************************************************************************************

function maxIE
{
param($ie)
$asm = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $ie.Width = $screen.width
    $ie.Height =$screen.height
    $ie.Top =  0
    $ie.Left = 0
}


# ************************************************************************************************************

function launchIE {
    $ie = new-object -comobject "InternetExplorer.Application"   
    $ie.navigate("file://" + $myPath + "/" + $myOutput) 
    maxIE $ie
    $ie.visible = $true 
} # end of launchIE

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

function makeWebPage {
    $today = Get-Date -Format D
    $outthing = "<HTML><HEAD><TITLE>Backup Report for " + $today + "</TITLE></HEAD><BODY>" + $reportheader + "<TABLE border = 1>"
    $schools | ForEach-Object -Process { 
    $outthing += "<TR><TD>" + $_.school_name  + "</TD><TD>" + $_.last_backup + "</TD>" ; 
    ForEach ($result in $_.backup_results) {
        if ($result.status -eq "Success") {$outthing += "<TD bgcolor = 'lightgreen' title='" + $result.description + "'>" + $result.status + "</TD>"}
        elseif ($result.status -eq "Warning") {$outthing += "<TD bgcolor = 'yellow' title='" + $result.description + "'>" + $result.status + "</TD>"}
        else {$outthing += "<TD bgcolor = 'red' title='" + $result.description + "'>" + $result.status + "</TD>"}
    }  
    $outthing += "</TR>`n"
    } 

    $outthing += "</TABLE> " + $reportfooter + "</BODY></HTML>`n"
    $outthing | Out-File -FilePath $myPath/$myOutput
} # end of makeWebPage

# ************************************************************************************************************

function acquireData {
    Write-Progress -Activity "Startup" -Status "reading targets from local file ..." -PercentComplete (0)

    if ($remoteload) {
    
    
        $targets = Excel_load
        $count = $targets.Length
        $myDrive = (getFreeDrive)[0]
        $schools = $targets | getSchoolData
        $schools | Export-Clixml -Path $myPath\$myXml
    }
    else {
        $schools = Import-Clixml -Path $myPath\$myXml
    }
    $schools
} # end of acquireData

# ************************************************************************************************************

# Main program

if ($args[0] -eq "local") { 
$remoteload = $false; 
$storedDate = Get-Item -Path $myPath\$myXml | select LastWriteTime | Get-Date -Format D

$reportheader = @"
<H1>David Moss's Backup Reporting Tool</H1>
<H3>Report generated from data archived on $storedDate</H3>
<HR />
<SMALL> hover mouse over the status boxes and the reason for the status will appear</SMALL>
<HR />
"@

} 

else {
$remoteload = $true ;
 $cred = getCredentials
 $reportDate = Get-Date -Format D

$reportheader = @"
<H1>David Moss's Backup Reporting Tool</H1>
<H3>Report generated on $reportDate</H3>
<HR />
<SMALL> hover mouse over the status boxes and the reason for the status will appear</SMALL>
<HR />
"@

}

$schools = acquireData
makeWebPage
launchIE

