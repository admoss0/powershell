# Scan for dodgy disks in Dell servers
# by David Moss
# 21/8/2014
# ****************************************

# dot reference the parallel magic
# source : http://gallery.technet.microsoft.com/scriptcenter/Run-Parallel-Parallel-377fd430
."H:\powershell\school_utils\Invoke-Parallel.ps1"

# Credit to Stray Muse Energy Company for the magic of "psexec $server omreport storage pdisk controller=0"
# http://social.technet.microsoft.com/Forums/windowsserver/en-US/0b6d9a02-08cd-4782-a983-263b46381ef6/how-to-list-the-total-no-of-physical-hdd-size-of-each-hdd?forum=winserverpowershell

clear

# some housekeeping and setting up
# the output file
$file = "d:\test.txt"
"School,Host,Disk,Status,State,PredFail" | Out-File -FilePath $file
# path to the input file
$myPath = "H:\powershell\school_utils"

# get user credentials suitable for the target systems, ie dds\st-admos0 and a password
$cred = Get-Credential


# Initialize an array
$colDiskErrors = @()
 
# import the target systems from the input file
$targets = Import-Csv -Path "$myPath/stuff.txt" -Delimiter `t

# Example from stuff.txt, itself generated from a powershell script:
# Name	LastLogonDate	OperatingSystem
# EQDDS0499005	23/06/2009 16:09	Windows Server 2003
# EQDDS2119005	27/04/2012 10:28	Windows Server 2003
# EQDDS2038013	29/07/2012 14:44	Windows Server 2008 R2 Enterprise

# set the parameters for the parallel magic
$info = New-Object -TypeName psobject -Property @{ 
            cred = $cred
			outfile = $file		
			}

# spin the algorithm to be parallelised into a here-string
$engine = @"
`$diskstatus = Invoke-Command -ComputerName (`$_.Name + ".dds.eq.edu.au") -ScriptBlock {omreport storage pdisk controller=0} -Credential `$parameter.cred
for (`$i=0; `$i -lt `$diskstatus.Count; `$i++) {
	if (`$diskstatus[`$i] -like "ID*") {
             	`$objDiskError = New-Object System.Object 
                `$objDiskError | Add-Member -type NoteProperty -name Hostname -value `$_.Name
				`$objDiskError | Add-Member -type NoteProperty -name ID  -value `$diskstatus[(`$i)].Substring(`$diskstatus[(`$i)].length-3)
                `$objDiskError | Add-Member -type NoteProperty -name Status  -value `$diskstatus[(`$i+1)].split(":")[1].trim()
                `$objDiskError | Add-Member -type NoteProperty -name State  -value `$diskstatus[(`$i+3)].split(":")[1].trim()
                `$objDiskError | Add-Member -type NoteProperty -name PredFail  -value `$diskstatus[(`$i+11)].split(":")[1].trim()
				`$line = `$_.Name + ',' +`$objDiskError.Hostname  + ',' + `$objDiskError.ID + ',' + `$objDiskError.Status + ',' + `$objDiskError.State + ',' + `$objDiskError.PredFail
				`$line | Out-File -FilePath `$parameter.outfile -Append
	}		
}

"@

# convert the here-string into a script block
$go = [Scriptblock]::Create($engine)


# Perform parallel magic !
Invoke-Parallel -InputObject $targets -parameter $info -Throttle 90 -runspaceTimeout 30 -ScriptBlock $go

# Example output:
# School,Host,Disk,Status,State,PredFail
# EQDDS1179001,EQDDS1179001,0:0,Ok,Online,No
# EQDDS0464001,EQDDS0464001,0:0,Non-Critical,Online,Yes
# EQDDS0003025,EQDDS0003025,0:4,Ok,Online,No
# EQDDS0003025,EQDDS0003025,0:5,Ok,Online,No
# EQDDS1177001,EQDDS1177001,0:1,Critical,Failed,No
# EQDDS0003025,EQDDS0003025,0:6,Ok,Online,No