$target = Read-Host -Prompt "Printer Info`r`nEnter server to scan"
Get-Printer -ComputerName $target | Export-Csv c:\temp\$target.csv

out-host -InputObject "results are in c:\temp\$target.csv"

