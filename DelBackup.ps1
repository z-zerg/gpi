# Количество дней, по истечении файлы удаляются
$EndDays = 35

$path = "D:\Backup\SQL\"
$netPath = "\\geops.local\admin\Backup\SQL"
$currentTime = Get-Date
$currentComp = $env:COMPUTERNAME

# Задаём параметры для отправки оповещения на E-mail
$encoding = [System.Text.Encoding]::UTF8
$emailSmtpServer = "exch1.geops.local"
$emailFrom = "Геопроектизыскания <support@geops.ru>"
$emailTo = "support@geops.ru"
$emailSubject = "Очистка избытка бэкапов"
$emailSubjectError = "Ошибка выполнения скрипта"
$emailBody = "<h2>Выполнено</h2>"

$emailFoter = "
Путь к каталогу резервных копий:<br>
$netPath <br><br>
<b>Исполняющий компьютер: $currentComp</b><br>
<b>Дата и время выполнения: $currentTime</b>"

$emailBodyError = "
<h2>Ошибка выполнения скрипта</h2>
Причина: <b>Каталог  бэкапов недоступен</b><br>
$path<br><br>"

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


# Сначала получаем все объекты и отбираем из них все, что не являются контейнерами, т.е. папками, а также старше указанного количества дней 
# Затем удаляем полученные файлы
Get-ChildItem -Path $path -Recurse | Where-Object {$_.PSisContainer -eq $false -and $_.LastWriteTime -lt ($(Get-Date).AddDays(-$EndDays))}| ForEach-Object {Remove-Item $_.FullName}

# Отправляем оповещение о выполнении
$emailBody = "
<h2>Избыток резервных копий удалён</h2>.<br><br>
<b>Итого:</b><br>
Число дней хранения резервных копий - <b>$EndDays</b><br>"
sendMail $emailBody