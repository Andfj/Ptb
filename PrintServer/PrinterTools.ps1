Import-Module C:\Scripts\InstallPrinter.psm1
Import-Module C:\Scripts\AddLHSPrinterPermissionSDDL.psm1
Import-Csv C:\Scripts\printers-vlan129.csv -Delimiter ";" -Encoding UTF8 | Install-Printer

#Get-Content "C:\Users\ABOROZ~1\AppData\Local\Temp\printer-log.txt" -Wait