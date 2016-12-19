
$printer = @{
"\\imb-prn\old printer 01" = "\\imb-prn01\new printer #01";
"\\imb-prn\old printer 02" = "\\imb-prn01\new printer #02";
}

try {

if (((Get-WmiObject win32_operatingsystem).caption).Contains("Windows 7")) {

if ($OldDefaultPrinter = (Get-WmiObject -Class Win32_Printer -ComputerName . | where { (($_.systemname -eq "\\imb-prn") -or ($_.systemname -eq "\\imb-prn.imb.local")) -and ($_.Default -eq $true)} | Select-Object name).name)
{
    $filename = "$env:COMPUTERNAME.txt"  
    $OldDefaultPrinter = $OldDefaultPrinter.ToLower()

    if ($OldDefaultPrinter.StartsWith("\\imb-prn.imb.local\")) { $OldDefaultPrinter = $OldDefaultPrinter.Replace("\\imb-prn.imb.local\","\\imb-prn\") }

    if ($NewDefaultPrinter = $printer.Get_Item($OldDefaultPrinter)) 
    {  
        (New-Object -ComObject WScript.Network).SetDefaultPrinter($NewDefaultPrinter)
        Write-Output "Note`r`n$((Get-Date -Format 'yyyy-MM-dd').ToString())`r`n$NewDefaultPrinter (default)`r`n$env:USERNAME"  | Out-File "\\imb-fs\printer$\$filename"
    }
    else 
    {   Write-Output "$((Get-Date -Format 'yyyy-MM-dd').ToString())`r`nNew Default Printer (instead of $OldDefaultPrinter) is not set as default`r`n$env:USERNAME"  | Out-File "\\imb-fs\printer$\-$filename"
    }
}

else 
{ 
    $filename = "_$env:COMPUTERNAME.txt"
    Write-Output "Note`r`n$((Get-Date -Format 'yyyy-MM-dd').ToString())`r`nno default printer from imb-prn`r`n$env:USERNAME" | Out-File -FilePath "\\imb-fs\printer$\$filename"
}

}

}
catch {

$filename = "_ERROR_$env:COMPUTERNAME-$((Get-Date -Format "yyyy-MM-dd").ToString()).txt"
Write-Output $Error[0] | Out-File -FilePath "\\imb-fs\printer$\$filename"
$Error.Clear()
}

##### section to remove printers from imb-prn
try {

if (((Get-WmiObject win32_operatingsystem).caption).Contains("Windows 7")) {

if ($OldPrinterToRemove = Get-WmiObject -Class Win32_Printer -ComputerName . | where { (($_.systemname -eq "\\imb-prn") -or ($_.systemname -eq "\\imb-prn.imb.local"))} | Select-Object name)
{
    $filename = "_R_$env:COMPUTERNAME-$((Get-Date -Format 'yyyy-MM-dd').ToString()).txt"  

    foreach ($p in $OldPrinterToRemove)
    {
        (New-Object -ComObject WScript.Network).RemovePrinterConnection($p.name)
        Write-Output "$((Get-Date -Format 'yyyy-MM-dd').ToString())`r`n$($p.name) removed from imb-prn`r`n$env:USERNAME" | Out-File -FilePath "\\imb-fs\printer$\$filename" -Append
    }

}

else 
{ 
    $filename = "_R_NO_$env:COMPUTERNAME.txt"
    Write-Output "$((Get-Date -Format 'yyyy-MM-dd').ToString())`r`nno printers from imb-prn`r`n$env:USERNAME" | Out-File -FilePath "\\imb-fs\printer$\$filename"
}

}

}
catch {

$filename = "_R_ERROR_$env:COMPUTERNAME.txt"
Write-Output $Error[0] | Out-File -FilePath "\\imb-fs\printer$\$filename" -a
$Error.Clear()
}