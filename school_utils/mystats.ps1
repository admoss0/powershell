import-module activedirectory
$target = "eqdds0163001"

 $os = Get-ADComputer -Filter {name -like $target} -Properties OperatingSystem | Select-Object OperatingSystem
 
 $Cdrive = Get-WmiObject Win32_LogicalDisk -ComputerName localhost -Filter "DeviceID='C:'" | Select-Object Size,FreeSpace
 $Ddrive = Get-WmiObject Win32_LogicalDisk -ComputerName localhost -Filter "DeviceID='D:'" | Select-Object Size,FreeSpace
 $Zdrive = Get-WmiObject Win32_LogicalDisk -ComputerName localhost -Filter "DeviceID='Z:'" | Select-Object Size,FreeSpace
 
 $C = "Free space on C drive is " + [int]($Cdrive.FreeSpace/1GB) +" Gigabytes, or " + [int](100 *$Cdrive.FreeSpace/$Cdrive.Size) + "%"
 $D = "Free space on D drive is " + [int]($Ddrive.FreeSpace/1GB) +" Gigabytes, or " + [int](100 *$Ddrive.FreeSpace/$Ddrive.Size) + "%"
 $Z = "Free space on Z drive is " + [int]($Zdrive.FreeSpace/1GB) +" Gigabytes, or " + [int](100 *$Zdrive.FreeSpace/$Zdrive.Size) + "%"
 
Write-Output ("Vital statistics for " + $target)
 $os.OperatingSystem
 $C
 $D
 $Z