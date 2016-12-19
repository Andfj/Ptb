<#
.Synopsis
   Install printer
.DESCRIPTION
   Install printers from text file. Log will be created in %TEMP%\printer-log.txt
.EXAMPLE
   Import-Csv printer-list.csv -delimeter ";" | Install-Printer
#>
function Install-Printer
{
    [CmdletBinding()]
    Param
    (
        # Printer IP Address
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        #[ValidatePattern("\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b")]
        [string]$PrnIPAddress,

        # Printer display name
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        #[ValidatePattern("^\w[a-zA-Z0-9_\s\-]+[a-zA-Z0-9_]$")]
        [String]$PrnName,
        
        # Printer driver name
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        #[ValidatePattern("^\w[a-zA-Z0-9_\s\-]+[a-zA-Z0-9_]$")]
        [String]$PrnDriver,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [String]$PrnLocation = "Default location",

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [String]$PrnComment = "Comment"

    )

    Begin
    {
        $s = New-PSSession -ComputerName DC1  -Credential (Get-Credential imb\aborozenets)
        Import-Module ActiveDirectory -PSSession $s
        Write-Output "Start of installing printer(s)..." | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append
    }
    Process
    {
        try
        {
            Write-Output "$((get-date -Format "yyyy/MM/dd HH:mm:ss").ToString()) INFO: Trying to create PortName $PrnIPAddress..." | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append
            Add-PrinterPort -Name $PrnIPAddress -PrinterHostAddress $PrnIPAddress  -PortNumber 9100 -SNMP 1 -SNMPCommunity public -ErrorAction Stop -ErrorVariable PrnError
            Write-Output "$((get-date -Format "yyyy/MM/dd HH:mm:ss").ToString()) INFO: PortName $PrnIPAddress created" | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append
     
            if ($SharedPrinter = Get-Printer | where { ($_.Shared -eq $true) -and ($_.ShareName -like 'PRN-*') } |`
                Sort-Object ShareName  | Select-Object -Property ShareName -Last 1) 
            {   
                $SharedPrinter = $SharedPrinter.ShareName
                $SharedPrinter = $SharedPrinter.Substring(4,3)
                $SharedPrinter = [int]$SharedPrinter +1
                $SharedPrinter = $SharedPrinter.ToString("000")
                $SharedPrinter = "PRN-$($SharedPrinter)"
            }
            else
            {
                $SharedPrinter = "PRN-001"
            }

            if ($PrinterName = Get-Printer | where { ($_.Shared -eq $true) -and ($_.ShareName -like 'PRN-*') -and ($_.Name -like "$PrnName*") } |`
                Sort-Object ShareName  | Select-Object -Property Name -Last 1) 
            {
                $PrinterName = $PrinterName.Name
                $PrinterName = $PrinterName.Replace("$PrnName #","")
                $PrinterName = [int]$PrinterName + 1
                $PrinterName = $PrinterName.ToString("00")
                $PrinterName = $PrnName + " #"+ $PrinterName
            }
            else
            {
                $PrinterName = $PrnName + " #01"
            }

            Write-Output "$((get-date -Format "yyyy/MM/dd HH:mm:ss").ToString()) INFO: Trying to create Printer $PrinterName..." | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append
            Add-Printer -Name $PrinterName -DriverName $PrnDriver -PortName $PrnIPAddress  -Comment $PrnComment -Location $PrnLocation -Shared -ShareName $SharedPrinter -Published -ErrorAction Stop -ErrorVariable PrnError
            Write-Output "$((get-date -Format "yyyy/MM/dd HH:mm:ss").ToString()) INFO: Printer $PrinterName [$PrnIPAddress] created" | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append

            #creating security group for Printer and apply permission
            Write-Output "$((get-date -Format "yyyy/MM/dd HH:mm:ss").ToString()) INFO: Trying to create Active Directory Security Group GS-$SharedPrinter..." | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append
            New-ADGroup -Name "GS-$SharedPrinter" -GroupCategory security -Description "Print permission for $PrinterName ($SharedPrinter)" -GroupScope Global -Path "OU=IMBGroups,DC=imb,DC=local" -ErrorAction Stop  -ErrorVariable PrnError
            Write-Output "$((get-date -Format "yyyy/MM/dd HH:mm:ss").ToString()) INFO: Security Group GS-$SharedPrinter created" | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append

            Start-Sleep -Seconds 20
            Write-Output "$((get-date -Format "yyyy/MM/dd HH:mm:ss").ToString()) INFO: Trying to add GS-$SharedPrinter group for $PrinterName..." | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append
            $PermissionSDDL = Get-Printer -full -Name $PrinterName | select PermissionSDDL -ExpandProperty PermissionSDDL
            $newSDDL = Add-LHSPrinterPermissionSDDL -Account "GS-$SharedPrinter" -existingSDDL $PermissionSDDL
            Get-Printer -Name $PrinterName | Set-Printer -PermissionSDDL $newSDDL -verbose -ErrorAction stop -ErrorVariable PrnError
            Write-Output "$((get-date -Format "yyyy/MM/dd HH:mm:ss").ToString()) INFO: Update security permission for $PrinterName" | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append
        }
        catch
        {
            Write-Output "$((get-date -Format "yyyy/MM/dd HH:mm:ss").ToString()) ERROR: $($PrnError[0].ErrorRecord)" | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append
            $PrnError.Clear()
        }
    }#end Process

    End
    {
        Get-Module ActiveDirectory | Remove-Module            
        Get-PSSession | Remove-PSSession     
        Write-Output "End of installing printer(s)." | Out-File -FilePath "$env:TEMP\printer-log.txt" -Append
    } 
} #end function  Install-Printer
