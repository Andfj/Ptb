$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://CASServer/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -AllowClobber

$Output="<html>
<body>
<font size=""1"" face=""Arial,sans-serif"">
<h5 align=""center"">Generated $((Get-Date).ToString())</h5>
</font>
<table border=""1"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"">
<tr>
<td>#</td>
<td><b>e-mail</b></td>
<td><b>Subject</b></td>
<td><b>sent time</b></td>
<td><b>response time</b></td>
<td><b>delta time</b></td>
<td><b>Real Sender</b></td>
</tr>"

$num = 1
$items_green = 0
$items_yellow = 0
$items_orange = 0
$items_red = 0
$items_all = 0

$CurrentDate = Get-Date -format "MM\/dd\/yyyy"

$Tomorrow = (Get-Date).AddDays(-1)
$Tomorrow = Get-Date $Tomorrow -format "MM\/dd\/yyyy"

$StartDate = "$Tomorrow 06:00:00"
$EndDate   = "$CurrentDate 06:00:00"

#$StartDate = "12/18/2015 06:00:00"
#$EndDate   = "12/19/2015 06:00:00"


$ReceivingLogs = Get-TransportServer imb-mx-ca* | Get-MessageTrackingLog -resultsize unlimited -Start $StartDate -EventId "DELIVER" -End $EndDate -Recipients SalaryProject@platinumbank.com.ua | `
	  where { (($_.Sender -NotLike "MicrosoftExchange*@platinumbank.com.ua") -and `
			   ($_.Sender -NotContains "Mediatel@platinumbank.com.ua")) `
            } | `
            Sort-Object timestamp

$bb = $ReceivingLogs.Clone()


foreach ($a in $ReceivingLogs) {
	$exist = 0
	$SubjectRe = $a.MessageSubject
	$SubjectReIndex = 0
	$firstsearch = 1

	$TimeSent = $a.timestamp

	if ($TimeSent.Minute -lt 10) { $TimeSentMinute  =  "0$($TimeSent.Minute)" } 
	else { $TimeSentMinute  =  $TimeSent.Minute }
	
	if ($TimeSent.Hour -lt 10) { $TimeSentHour  =  "0$($TimeSent.Hour)" } 
	else { $TimeSentHour  =  $TimeSent.Hour }

	if ($TimeSent.Month -lt 10) { $TimeSentMonth  =  "0$($TimeSent.Month)" } 
	else { $TimeSentMonth  =  $TimeSent.Month }

	if ($TimeSent.Day -lt 10) { $TimeSentDay  =  "0$($TimeSent.Day)" } 
	else { $TimeSentDay  =  $TimeSent.Day }
	
	$TimeSentStr = "${TimeSentHour}:$TimeSentMinute  $($TimeSent.Year)/$TimeSentMonth/$TimeSentDay"

	
	if ($SubjectRe.length -gt 3) { 	
		if ($SubjectRe.Substring(0,4) -eq "Re: ")  { $SubjectReIndex = 1 }
		if ($SubjectRe.Substring(0,4) -eq "Hà: ")  { $SubjectReIndex = 1 }
		if ($SubjectRe.Substring(0,4) -eq "FW: ")  { $SubjectReIndex = 1 }
	}
	
	if ($SubjectRe.length -gt 4) { 	
		if ($SubjectRe.Substring(0,5) -eq "Fwd: ") { $SubjectReIndex = 1 }
	}

	foreach ($b in $bb) {
		if ($b.timestamp -gt $a.timestamp) {
	
			if (($b.MessageSubject -eq "Re: " + $a.MessageSubject) -or (($SubjectReIndex -eq 1) -and (($b.MessageSubject -eq $a.MessageSubject) -or (($a.MessageSubject).Substring(5) -eq ($b.MessageSubject).Substring(4)) -or (($a.MessageSubject).Substring(4) -eq ($b.MessageSubject).Substring(4)) ))) {
				
				$TimeResponce = $b.timestamp

				if ($TimeResponce.Minute -lt 10) { $TimeResponceMinute  =  "0$($TimeResponce.Minute)" } 
				else { $TimeResponceMinute  =  $TimeResponce.Minute }
	
				if ($TimeResponce.Hour -lt 10) { $TimeResponceHour  =  "0$($TimeResponce.Hour)" } 
				else { $TimeResponceHour  =  $TimeResponce.Hour }

				if ($TimeResponce.Month -lt 10) { $TimeResponceMonth  =  "0$($TimeResponce.Month)" } 
				else { $TimeResponceMonth  =  $TimeResponce.Month }

				if ($TimeResponce.Day -lt 10) { $TimeResponceDay  =  "0$($TimeResponce.Day)" } 
				else { $TimeResponceDay  =  $TimeResponce.Day }

				$TimeResponceStr = "${TimeResponceHour}:$TimeResponceMinute $($TimeResponce.Year)/$TimeResponceMonth/$TimeResponceDay"
			
				$TimeToAnswer = $b.timestamp - $a.timestamp
				
				if ($TimeToAnswer.Minutes -lt 10) { $TimeToAnswerMinutes  =  "0$($TimeToAnswer.Minutes)" } 
				else { $TimeToAnswerMinutes  =  $TimeToAnswer.Minutes }
	
				if ($TimeToAnswer.Hours -lt 10) { $TimeToAnswerHours  =  "0$($TimeToAnswer.Hours)" } 
				else { $TimeToAnswerHours  =  $TimeToAnswer.Hours }

				$TimeDelta = "$($TimeToAnswer.Days) days ${TimeToAnswerHours}:$TimeToAnswerMinutes"
				
				if (($TimeToAnswer  -ge 0)  -and ($firstsearch -eq 1))  { 
					$cc = Get-TransportServer imb-mx-ca* | Get-MessageTrackingLog -resultsize unlimited -Start $StartDate -EventId "RECEIVE" -End $EndDate -MessageId $b.MessageId 
					$Output +=  "<tr><td>$($num)</td><td>$($a.Sender)</td><td>$($a.MessageSubject)</td><td>$TimeSentStr</td><td>$TimeResponceStr</td><td>$TimeDelta</td><td>$($cc.Sender) </td></tr>"
					$num += 1
					$firstsearch +=1
					$exist = 1
				}
			}
		}
	}
	
	if ($exist -ne  1) { 
		$Output +=  "<tr><td>$($num)</td><td>$($a.Sender)</td><td>$($a.MessageSubject)</td><td>$TimeSentStr</td><td>&nbsp;</td><td>&nbsp;</td></tr>" 
		$num += 1
	}
	
}

$Output += "</table>"
$Output += "</body></html>"
$Output | Out-File c:\scripts\SalaryProject-daily.htm
Send-MailMessage -Attachments "C:\scripts\SalaryProject-daily.htm" -To "SalaryProjectMailReport@platinumbank.com.ua"  -From "SalaryProjectMailReport@platinumbank.com.ua" -Subject "SalaryProject Report Daily" -BodyAsHtml $Output -SmtpServer mail.imb.local