#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://imb-mx-ca02.imb.local/PowerShell/ -Authentication Kerberos
#Import-PSSession $Session

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

#statistics for all, except back office
$items_green = 0
$items_yellow = 0
$items_red = 0
$items_all = 0

#statistics for back office
$items_light_blue = 0 #light blue
$items_purple = 0 #purple
$items_dark_blue = 0 #dark_blue
$items_all_bo = 0


$CurrentDate = Get-Date -format "MM\/dd\/yyyy"

$Tomorrow = (Get-Date).AddDays(-1)
$Tomorrow = Get-Date $Tomorrow -format "MM\/dd\/yyyy"

$StartDate = "$Tomorrow 00:00:00"
$EndDate   = "$CurrentDate 00:00:00"

#$StartDate = "05/12/2015 00:00:00"
#$EndDate   = "05/13/2015 00:00:00"


$aa = Get-TransportServer casserver* | Get-MessageTrackingLog -resultsize unlimited -Start $StartDate -EventId "DELIVER" -End $EndDate -Recipients ifobssupport@platinumbank.com.ua | `
	  where { (($_.Sender -NotLike "MicrosoftExchange*@platinumbank.com.ua") -and `
			   -not ($_.MessageSubject -Like "Подключили клиента*") -and `
			   -not (($_.MessageSubject -Like "IFOBS*В работе") -and ($_.Sender -eq "sharepoint@platinumbank.com.ua")) -and `
			   (($_.MessageSubject).trim() -ne "сертификаты")) } | `
	  Sort-Object timestamp
$bb = $aa.Clone()


foreach ($a in $aa) {
	$exist = 0
	$SubjectRe = $a.MessageSubject
	$SubjectReIndex = 0
	$firstsearch = 1
    $SubjectStart = ""

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
		if ($SubjectRe.Substring(0,4) -eq "Hа: ")  { $SubjectReIndex = 1 }
		if ($SubjectRe.Substring(0,4) -eq "FW: ")  { $SubjectReIndex = 1 }
	}
	
	if ($SubjectRe.length -gt 4) { 	
		if ($SubjectRe.Substring(0,5) -eq "Fwd: ") { $SubjectReIndex = 1 }
	}


    if ($a.Sender -eq "sharepoint@platinumbank.com.ua") {
        if ($a.MessageSubject -like "IFOBS*Новое"  ) { 
            $SubjectReIndex = 2 
            $SubjectStart = ($a.MessageSubject).Substring(0,($a.MessageSubject).length - 6)
        }
    }

###### ставим $SubjectReIndex=3 чтобы письма с этой темой не попали в отчет
    if ($a.Sender -eq "sharepoint@platinumbank.com.ua") {
        if (($a.MessageSubject -like "IFOBS*Выполнено") -or ($a.MessageSubject -like "I FOBS*Отмена. неверно оформлено")) { 
        $SubjectReIndex = 3 
        }
    }

###### ставим $SubjectReIndex=3 чтобы письма с этой темой не попали в отчет
    if ($a.Sender -eq "sharepoint@platinumbank.com.ua") {
        if (($a.MessageSubject -like "Received application:*Заявление обработано") -or `
            ($a.MessageSubject -like "Received application:*Отказ в приёме заявления")) { 
        $SubjectReIndex = 3 
        }
    }


    if (($a.Sender -eq "oracle@b2.imb.local") -and ($a.MessageSubject -like "Received application*;")) {
        $SubjectReIndex = 4 
        $SubjectStart = ($a.MessageSubject).Substring(0,($a.MessageSubject).length - 1)
    }

    if (($a.Sender -eq "sharepoint@platinumbank.com.ua") -and ($a.MessageSubject -like "Received application:*Заявление в работе")) {
        $SubjectReIndex = 5 
        $SubjectStart = ($a.MessageSubject).Substring(0,($a.MessageSubject).length - 19)
    }



	foreach ($b in $bb) {
		if ($b.timestamp -gt $a.timestamp) {

            $SubjectEnd = ""
            if ($b.Sender -eq "sharepoint@platinumbank.com.ua") {
                if (($b.MessageSubject -like "IFOBS*Выполнено") -and ($SubjectReIndex -eq 2)) { 
                    $SubjectEnd = ($b.MessageSubject).Substring(0,($b.MessageSubject).length - 10)
                }
                if (($b.MessageSubject -like "IFOBS*Отмена. неверно оформлено") -and ($SubjectReIndex -eq 2)) { 
                    $SubjectEnd = ($b.MessageSubject).Substring(0,($b.MessageSubject).length - 26)
                }

            }

            if ($SubjectReIndex -eq 3) { $exist = 1 }

############### 
            if (($SubjectReIndex -eq 4) -and ($b.MessageSubject -like "Received application*Заявление в работе"))  { 
                $SubjectEnd = ($b.MessageSubject).Substring(0,($b.MessageSubject).length - 19)
            }

############### 
            if (($SubjectReIndex -eq 5) -and ($b.MessageSubject -like "Received application*Заявление обработано"))  { 
                $SubjectEnd = ($b.MessageSubject).Substring(0,($b.MessageSubject).length - 21)
            }
            if (($SubjectReIndex -eq 5) -and ($b.MessageSubject -like "Received application:*Отказ в приёме заявления"))  { 
                $SubjectEnd = ($b.MessageSubject).Substring(0,($b.MessageSubject).length - 25)
            }


# подсчет времени, если тема начинается на IFOBS*Новое или Received application...
            if (($SubjectStart -eq $SubjectEnd) -and (($SubjectReIndex -eq 2) -or ($SubjectReIndex -eq 4))) { 
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
				
				if ($TimeToAnswer -lt "00:30:00") { $row_col = "66FF00" } #green
				if (($TimeToAnswer -ge "00:30:00") -and ($TimeToAnswer -lt "01:00:00")) { $row_col = "FFFF66" } #yellow
				if ($TimeToAnswer -gt "01:00:00") { $row_col = "FF3333" } #red
				
				if (($TimeToAnswer  -ge 0)  -and ($firstsearch -eq 1))  { 
					if ($row_col -eq "66FF00") { $items_green +=1 }
					if ($row_col -eq "FFFF66") { $items_yellow +=1 }
					if ($row_col -eq "ff3333") { $items_red +=1 }
				
					$cc = Get-TransportServer casserver* | Get-MessageTrackingLog -resultsize unlimited  -Start $StartDate -EventId "RECEIVE" -End $EndDate -MessageId $b.MessageId 
					$Output +=  "<tr bgcolor=$($row_col)><td>$($num)</td><td>$($a.Sender)</td><td>$($a.MessageSubject)</td><td>$TimeSentStr</td><td>$TimeResponceStr</td><td>$TimeDelta</td><td>$($cc.Sender)</td></tr>"
					$num += 1
					$firstsearch +=1
					$exist = 1
                }
            }


# подсчет времени, если тема начинается на Received application...Заявление в работе
            if (($SubjectStart -eq $SubjectEnd) -and ($SubjectReIndex -eq 5)) { 
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
				
				if ($TimeToAnswer -lt "00:15:00") { $row_col = "7cfff1" } #light blue
				if (($TimeToAnswer -ge "00:15:00") -and ($TimeToAnswer -lt "00:30:00")) { $row_col = "b99cff" } #purple
				if ($TimeToAnswer -ge "00:30:00") { $row_col = "468f88" } # dark blue
				
				if (($TimeToAnswer  -ge 0)  -and ($firstsearch -eq 1))  { 
					if ($row_col -eq "7cfff1") { $items_light_blue +=1 }
					if ($row_col -eq "b99cff") { $items_purple +=1 }
					if ($row_col -eq "468f88") { $items_dark_blue +=1 }
				
					$cc = Get-TransportServer casserver* | Get-MessageTrackingLog -resultsize unlimited -Start $StartDate -EventId "RECEIVE" -End $EndDate -MessageId $b.MessageId 
					$Output +=  "<tr bgcolor=$($row_col)><td>$($num)</td><td>$($a.Sender)</td><td>$($a.MessageSubject)</td><td>$TimeSentStr</td><td>$TimeResponceStr</td><td>$TimeDelta</td><td>$($cc.Sender)</td></tr>"
					$num += 1
					$firstsearch +=1
					$exist = 1
                }
            }

#подсчет времени по общим правилам
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
				
				if ($TimeToAnswer -lt "00:30:00") { $row_col = "66FF00" } #green
				if (($TimeToAnswer -ge "00:30:00") -and ($TimeToAnswer -lt "01:00:00")) { $row_col = "FFFF66" } #yellow
				if ($TimeToAnswer -gt "01:00:00") { $row_col = "FF3333" } #red
				
				if (($TimeToAnswer  -ge 0)  -and ($firstsearch -eq 1))  { 
					if ($row_col -eq "66FF00") { $items_green +=1 }
					if ($row_col -eq "FFFF66") { $items_yellow +=1 }
					if ($row_col -eq "ff3333") { $items_red +=1 }
				
					$cc = Get-TransportServer casserver* | Get-MessageTrackingLog -resultsize unlimited -Start $StartDate -EventId "RECEIVE" -End $EndDate -MessageId $b.MessageId 
					$Output +=  "<tr bgcolor=$($row_col)><td>$($num)</td><td>$($a.Sender)</td><td>$($a.MessageSubject)</td><td>$TimeSentStr</td><td>$TimeResponceStr</td><td>$TimeDelta</td><td>$($cc.Sender)</td></tr>"
					$num += 1
					$firstsearch +=1
					$exist = 1
				}
			}
		}
	}
#вывести данные, если найдено больше 2-х ответов или если ответа вообще не найдено	
	if ($exist -ne  1) { 
		$Output +=  "<tr><td>$($num)</td><td>$($a.Sender)</td><td>$($a.MessageSubject)</td><td>$TimeSentStr</td><td>&nbsp;</td><td>&nbsp;</td></tr>" 
		$num += 1
	}
	
}

#подсчет статистики для всех кроме back office  (проценты)
$items_all 		= $items_green + $items_yellow + $items_red
$items_green_p 	= [decimal]::round(($items_green/$items_all)*100)
$items_yellow_p 	= [decimal]::round(($items_yellow/$items_all)*100)
$items_red_p		= [decimal]::round(($items_red/$items_all)*100)

$items_green_b  = $items_green_p*9
$items_yellow_b = $items_yellow_p*9
$items_red_b	= $items_red_p*9


#подсчет статистики для back office (проценты)
$items_all_bo 		= $items_light_blue + $items_purple + $items_dark_blue
$items_light_blue_p	= [decimal]::round(($items_light_blue/$items_all_bo)*100)
$items_purple_p 	= [decimal]::round(($items_purple/$items_all_bo)*100)
$items_dark_blue_p	= [decimal]::round(($items_dark_blue/$items_all_bo)*100)

$items_light_blue_b  = $items_light_blue_p*9
$items_purple_b = $items_purple_p*9
$items_dark_blue_b    = $items_dark_blue_p*9



$Output += "</table><br><b>Total (except Back Office):</b><br>"
$Output += "<div style=""width:$($items_green_b)px;height:20px;background:#66FF00;text-align:center;white-space:nowrap;font-size:10pt;font-family:Arial,sans-serif""><b>$($items_green_p)% ($($items_green) messages)</b></div>"
$Output += "<div style=""width:$($items_yellow_b)px;height:20px;background:#FFFF66;text-align:center;white-space:nowrap;font-size:10pt;font-family:Arial,sans-serif""><b>$($items_yellow_p)% ($($items_yellow) messages)</b></div>"
$Output += "<div style=""width:$($items_red_b)px;height:20px;background:#ff3333;text-align:center;white-space:nowrap;font-size:10pt;font-family:Arial,sans-serif""><b>$($items_red_p)% ($($items_red) messages)</b></div>"

$Output += "<br><b>Total (only for Back Office):</b><br>"

$Output += "<div style=""width:$($items_light_blue_b)px;height:20px;background:#7cfff1;text-align:center;white-space:nowrap;font-size:10pt;font-family:Arial,sans-serif""><b>$($items_light_blue_p)% ($($items_light_blue) messages)</b></div>"
$Output += "<div style=""width:$($items_purple_b)px;height:20px;background:#b99cff;text-align:center;white-space:nowrap;font-size:10pt;font-family:Arial,sans-serif""><b>$($items_purple_p)% ($($items_purple) messages)</b></div>"
$Output += "<div style=""width:$($items_dark_blue_b)px;height:20px;background:#468f88;text-align:center;white-space:nowrap;font-size:10pt;font-family:Arial,sans-serif""><b>$($items_dark_blue_p)% ($($items_dark_blue) messages)</b></div>"

$Output += "</body></html>"
$Output | Out-File c:\scripts\ifobssupport.htm
Send-MailMessage -Attachments "C:\scripts\ifobssupport.htm" -To "iFOBSSupport@platinumbank.com.ua"  -From "iFOBSSupport@platinumbank.com.ua" -Subject "iFOBSSupport Report Daily" -BodyAsHtml $Output -SmtpServer smtp-server