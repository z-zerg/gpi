
# Путь к конечной папке
$ArchiveDstPath = "\\srv\cons$\SEND_MAIL"

# Путь к первой архивируемой папке
$ArchiveSrcPath1 = "\\srv\cons$\receive"

# Путь к первой архивируемой папке
$ArchiveSrcPath2 = "\\srv\cons$\adm\STS"

# Имя задания для подстановки в имя конечного файла
$ArchiveTaskName = "USR_геопроектизыскания"

# Указывем путь к архиватору
$ArchivatorPath = "C:\Program Files (x86)\7-Zip\7z.exe"

# Проверяем существование конечной директори
$path = Test-Path $ArchiveDstPath
if ($path -eq $False){
     Write-host "Директория НЕ существует"
     New-Item -Path $ArchiveDstPath -ItemType "directory"   
    }
   else {Write-host "Директория существует"}
   
   Copy-Item -Path $ArchiveSrcPath1 -Recurse $ArchiveDstPath -force
   Copy-Item -Path $ArchiveSrcPath2 -Recurse $ArchiveDstPath -force



#Создаем массив параметров для 7-Zip
	$Arg1="a" 
	$Arg2="-tzip"    
	$Arg3="-ssw"
	$Arg4="-mx5"
	$Arg5=$ArchiveDstPath+"\$(Get-Date -format "yyyy-MM-dd")_"+$ArchiveTaskName+".zip"
    $Arg6=$ArchiveSrcPath1
    $Arg7=$ArchiveSrcPath2    

# Архивируем
& $ArchivatorPath ($Arg1,$Arg2,$Arg3,$Arg4,$Arg5,$Arg6)
& $ArchivatorPath ($Arg1,$Arg2,$Arg3,$Arg4,$Arg5,$Arg7)



# Задаём параметры для отправки на E-mail

# Задаём кодировку письма
$encoding = [System.Text.Encoding]::UTF8

# Адрес сервера
$emailSmtpServer = "exch1.geops.local"

# Обратный адрес
$emailFrom = "Геопроектизыскания <support@geops.ru>"

# Адрес получателя
# $emailTo = "trushkov@geops.ru", "serge-trushkov@yandex.ru"
 $emailTo = "Вихарева Дарья Салаватовна <dvikhare@4dk.ru>"

# Адрес, куда посылается копия письма
$CC = "support@geops.ru"

# Тема письма
$emailSubject = "USR файлы от Геопроектизыскания"

# Тело письма. Возможны HTML теги
$emailBody = @"
<p>Здравствуйте,<br>направляем Вам архив с файлами, которые Вы запрашивали.</p></br>
-------------------<br>
С уважением,</br>служба поддержки ООО "Геопроектизыскания"
"@

# Указываем файл вложения
# $attachment = "d:\mail1.txt","d:\mail2.txt"
  $attachment = $Arg5
  
# Отправляем письмо 
Send-MailMessage -To $emailTo -CC $CC -From $emailFrom -Subject $emailSubject -Body $emailBody -BodyAsHTML -Attachments $attachment -SmtpServer $emailSmtpServer -Encoding $encoding

# Удаляем конечную папку
Remove-Item $ArchiveDstPath -recurse