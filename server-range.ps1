<#
.Description
uses the current IP address from WMI and uses it to
generate 10 ip addresses from the server range of the same IP range.
It then pings each of these addresses to determine if they are on-line
and returns an array of objects representing the servers. 
#>
function get-ServerRange() {

$serverRange = @();
$colItems = Get-WmiObject Win32_NetworkAdapterConfiguration -Namespace "root\CIMV2" | where{$_.IPEnabled -eq “True”};

$thisip = $colItems.IPAddress[0].split(".");
# $thisip = "10.153.60.22".split(".");

for ($i=1; $i -lt 11; $i++) {

    $serverip= $thisip[0]+"."+$thisip[1]+"."+$thisip[2]+"."+$i;
    $serveripdec =  [int]$thisip[0] * [math]::pow(256,3) + [int]$thisip[1] * [math]::pow(256,2) + [int]$thisip[2] * 256 + [int]$i
    
    $myob = New-Object -TypeName PSOBJECT
    $myob | Add-Member -MemberType ScriptMethod -Name 'checkOK' -Value {Test-Connection -ComputerName $this.ipAddress -Quiet -Count 1} -passthru
    $myob | Add-Member -MemberType NoteProperty -Name 'ipAddress' -Value $serverip
    $myob | Add-Member -MemberType NoteProperty -Name 'ipAddressAsDecimal' -Value $serveripdec
    $myob | Add-Member -MemberType NoteProperty -Name 'isOK' -Value ""
    $myob.isOK = $myob.checkOK()
    $serverRange += $myob
    
} # end for

$serverRange = $serverRange.GetEnumerator() | Sort-Object ipAddressAsDecimal
 
return $serverRange
} # end function
