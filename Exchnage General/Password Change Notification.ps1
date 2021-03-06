##################################################################################################################
# Please Configure the following variables....
$smtpServer="mail.imb.local"
$from = "Password Change Notification <no-reply@platinumbank.com.ua>"
$expireindays = 15
###################################################################################################################

#Get Users From AD who are enabled
Import-Module ActiveDirectory
$users = get-aduser -filter * -properties * | where {$_.Enabled -eq "True"} | where { $_.PasswordNeverExpires -eq $false } | where { $_.passwordexpired -eq $false } | where { $_.emailaddress -ne $null }

foreach ($user in $users)
{
  $Name = (Get-ADUser $user | foreach { $_.Name})
  $emailaddress = $user.emailaddress
  $passwordSetDate = (get-aduser $user -properties * | foreach { $_.PasswordLastSet })
  $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
  $expireson = $passwordsetdate + $maxPasswordAge
  $today = (get-date)
  $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
  $expireson = Get-Date $expireson -Format F
  $encoding = [System.Text.Encoding]::UTF8
  $subject="Срок действия вашего пароля истекает через $daystoExpire дней"
  $body ="
$name,<br><br>
<p> Пароль вашей учетной записи истекает через $daystoexpire дней.<br>
 Вам необходимо изменить пароль. Пожалуйста, выберите удобный для Вас способ:<br>
<ol>  
<li>Сменить пароль при помощи электронной почты.<br>
Перейдите по ссылке <a href='https://webmail.platinumbank.com.ua'>https://webmail.platinumbank.com.ua</a>, в правом верхнем углу необходимо выбрать <b>&laquo;Параметры&raquo; - &laquo;Сменить пароль&raquo;</b> и в текущем окне изменить свой пароль. <br>
Подробную инструкцию можно найти на MyPlatinum по ссылке: <a href='https://platinumdocs.platinumbank.com.ua/HD/SitePages/Change_password_owa.aspx'>https://platinumdocs.platinumbank.com.ua/HD/SitePages/Change_password_owa.aspx</a>
<br><br>
</li>
<li>
Сменить пароль на компьютере, подключенном к сети банка.<br>
На своей рабочей станции нажмите <b>Ctrl+Alt+Del</b> и выберите <b>&laquo;Сменить пароль…&raquo;</b>. <br>
Подробную инструкцию можно найти на MyPlatinum по ссылке: <a href='https://platinumdocs.platinumbank.com.ua/HD/SitePages/Change_password_computer.aspx'>https://platinumdocs.platinumbank.com.ua/HD/SitePages/Change_password_computer.aspx</a>
</li>
</ol>

Обращаем ваше внимание, что менять пароли можно не чаще 1 раза в сутки. Ваш новый пароль должен соответствовать требованиям парольной политики, используемой в банке, а именно:
<ul>
<li>не повторять последние 24 пароля;</li>
<li>минимальная длина пароля 8 символов;</li>
<li>содержать 3-и из 4-х категорий символов – большие английские буквы, маленькие английские буквы, цифры от 0 до 9, специальные символы (например  !, $, #, %).</li>
</ul>

<p>Если пароль не будет изменен  до $expireson, учетная запись будет заблокирована.</p><br>

Если у Вас возникнут вопросы либо сложности со сменой пароля, вы всегда можете обратиться в отдел технической поддержки по адресу <a href='mailto:ITHelpdesk@platinumbank.com.ua?subject=Проблема смены пароля учетной записи'>ITHelp@platinumbank.com.ua</a>.
<br>
<br>
Спасибо!
<br>
Ваш ИТ
<br><br><br>
<p>ЭТО АВТОМАТИЧЕСКИ СОЗДАННОЕ СООБЩЕНИЕ, ПРОСЬБА НЕ ОТВЕЧАТЬ НА ДАННОЕ ПИСЬМО</p>
  "
  
  if ($daystoexpire -lt $expireindays)
  {
    Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High -Encoding $encoding
     
  }  
   
}
