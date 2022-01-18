$currentTime = Get-Date
$searchBase = "ou=Users, ou=moscow, dc=geops, dc=local"
$server = "srv.geops.local"
$path = "F:\Obmenka1"
$netPath = "\\geops.local\document\Обменная"
$currentComp = $env:COMPUTERNAME

# Задаём параметры для отправки оповещения на E-mail
    $encoding = [System.Text.Encoding]::UTF8
    $emailSmtpServer = "exch1.geops.local"
    $emailFrom = "Геопроектизыскания <support@geops.ru>"
    $emailTo = "support@geops.ru"
    $emailSubject = "Очистка обменки"
    $emailSubjectError = "Ошибка выполнения скрипта"
    $emailBody = "<h2>Выполнено</h2>"

    $emailBodyError = "
    <h2>Ошибка выполнения скрипта</h2>
    Причина: <b>Каталог недоступен</b><br>
    $path<br><br>"

    $emailFoter = "
    Путь к каталогу резервных копий:<br>
    $netPath <br><br>
    <b>Исполняющий компьютер: $currentComp</b><br>
    <b>Дата и время выполнения: $currentTime</b>"

# Отправка уведомления
    function sendMail {
        param ($eBody)
        $eBody += $emailFoter
        Send-MailMessage -To $emailTo -From $emailFrom -Subject $emailSubject -Body $eBody -BodyAsHTML -SmtpServer $emailSmtpServer -Encoding $encoding

    }



# проверяем доступность пути для выгрузки и отправляем в очередь
    $TestBackupPath = Test-Path $path
    if (!($TestBackupPath)) {
        Write-Host "Каталог бэкапов недоступен. Дальнейшее выполнение невозможно
        отправляем уведомление"
        $emailBody = $emailBodyError
        $emailSubject = $emailSubjectError
        sendMail $emailBody
        exit
    }


Import-Module ActiveDirectory


# --- Находим всех пользователей в AD в пределах области поиска
$users = Get-ADUser -Filter * -SearchBase $searchBase

# Сначала удаляем всё из папок
remove-item $path\ -recurse -force

# Создаём новые папки
Foreach ($folder in $users) {
    New-Item -Path "$path\$($folder.Name)" -type "directory"
    }

    
# Отправляем оповещение о выполнении
$emailBody = "
<h2>Обменка очищена.Папки пользователей созданы</h2>.<br>"

sendMail $emailBody