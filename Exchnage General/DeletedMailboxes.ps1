$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://CASServer/PowerShell/ -Authentication Kerberos
Import-PSSession $Session

#Remove old mailboxes after 14 day after disabling in Active Directory

Import-Module activedirectory

$d = Get-Date
$d = $d.AddDays(-14)
$encoding = [System.Text.Encoding]::UTF8
$num = 1

$output="<html>
<body>
<h3 align=""center"" style=""font-size:12pt;font-family:Calibri"">Deleted mailboxes after 14 days of inactivity</h3>
<table border=""1"" cellpadding=""2"" style=""font-size:12pt;font-family:Calibri"">
<tr>
<td>#</td>
<td><b>Display Name</b></td>
<td><b>Title</b></td>
<td><b>DismissalDate</b></td>
</tr>"


$mailboxes_to_delete = Get-ADUser -SearchBase "OU=ORGSTR,DC=imb,DC=local" -Properties * -Filter { (Enabled -eq $false) -and (mail -like "*") -and (samaccountname -ne "aborozenets") } -ResultSetSize $null | where  { $_.imbUserDismissalDate -ne $null } 


$Culture = [Globalization.cultureinfo]::GetCultureInfo("en-US")



try {

foreach ($mailbox_to_delete in $mailboxes_to_delete) {

    $DismissalDate = [datetime]::ParseExact($mailbox_to_delete.imbUserDismissalDate, "dd.MM.yyyy", $Culture)
    
    if ($DismissalDate -lt $d) {
        Disable-Mailbox -Identity $mailbox_to_delete.DistinguishedName -confirm:$false -ErrorAction stop
        $output+="<tr><td>$($num)</td><td>$($mailbox_to_delete.DisplayName)</td><td>$($mailbox_to_delete.title)</td><td>$($mailbox_to_delete.imbUserDismissalDate)</td>"
        $num+=1
    }

}

$output +="</table></body></html>"
Send-MailMessage -From "Deleted Mailboxes <np-exch-scripts@platinumbank.com.ua>" -To "DeletedMailboxesReport@platinumbank.com.ua" -Subject "Deleted mailboxes" -Body $output -BodyAsHtml -SmtpServer mail.imb.local -Encoding $encoding
}

catch {
Send-MailMessage -From "Deleted Mailboxes <np-exch-scripts@platinumbank.com.ua>" -To "DeletedMailboxesReport@platinumbank.com.ua" -Subject "Deleted mailboxes Error" -Body $error[0] -SmtpServer mail.imb.local -Encoding $encoding
}
