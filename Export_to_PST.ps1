# Скрипт принимает два параметра - это минимальный и максимальный размер ящика
# По этим параметрам выполняется выборка ящиков и последующий экспорт
# Передача параметров выполняется следующим образом
# ./Export_to_PST.ps1 0Gb 5Gb



# -----------   Входные данные ------------------------

# Указываем сервер, с которого получаем список пользователей
    $ADserver = "srv.geops.local"    

# Указываем область поиска пользователей в домене
    $searchBase = "OU=Users,OU=Moscow,DC=geops,DC=local"  

# Указываем OU в домене для выборки пользователей
    $ExportUsersOU = "geops.local/Moscow/Users"      

# Указываем путь для выгрузки PST файлов
    $path = "\\geops.local\admin\Backup\Mailbox"   

# Минимальный размер ящика (для фильтрации)
    if ($args -eq $null){
        $minMailboxSize = 0
    } else {
        $minMailboxSize = $args[0]  
    }
    

# Максимальный размер ящика (для фильтрации)
    if ($args -eq $null){
        $maxMailboxSize = 100Mb
    } else {
        $maxMailboxSize = $args[1]
    }
    


    $minScope = $minMailboxSize / 1024 / 1024
    $maxScope = $maxMailboxSize / 1024 / 1024

# Имя подключаемой оснастки
    $exchSnapName = "Microsoft.Exchange.Management.PowerShell.E2010"

# текущие дата и время и компьютер
    $currentTime = Get-Date
    $currentComp = $env:COMPUTERNAME

# Задаём параметры для отправки оповещения на E-mail
    $encoding = [System.Text.Encoding]::UTF8
    $emailSmtpServer = "exch1.geops.local"
    $emailFrom = "Геопроектизыскания <support@geops.ru>"
    $emailTo = "support@geops.ru"
    $emailSubject = "Backup почтовых ящиков"
    $emailSubjectError = "Ошибка выполнения скрипта"
    $emailBody = "<h3>Backup ящиков объёмом от $minScope Мб до $maxScope Мб</h3>"
    $errorBody = "<h2>Ошибка выполнения скрипта</h2>Причина: <b>Каталог для выгрузки недоступен</b><br><br>Проверьте доступность пути:<br>$path<br><br>"
    $emailFoter = "
    <b>Исполняющий компьютер: $currentComp</b><br>
    <b>Время и дата отправки: $currentTime</b>
    "

    $SendMailboxList = ""
    $FailedSendMailboxList = ""
    $ProgressSendMailboxList = ""
    $CompletedSendMailboxList = ""
    $QueuedSendMailboxList = ""

    $CurrentProcess = @()
    $FailedProcess = @()
    $CompletedProcess = @()
    $QueuedProcess = @()

    $CompletedCount = 0
    $FailedCount = 0
    $CurrentCount = 0
    $QueuedCount = 0

# По умолчанию оснастка Exchange не установлена
    $exShell = $false

# ------------------ Функции ------------------

# Сравниваем объекты на предмет совпадения
    function CompareObject{
        param ($p1, $p2, $p11, $p22)

        $ob = @()
        ForEach ($i in $p1) {
            ForEach ($j in $p2) {
                if ($i.$p11 -eq $j.$p22 ) {            
                    $ob += $i                
                }
            }
        }
       return $ob
    }

# Выборка по фильтру статуса
    function getStatusMailbox{
        param ($arr)        
        $res = ""
        $temp = @()     
          $arr | ForEach-Object {
                if ($global:exShell -eq $false){
                    $temp = $_.Mailbox -split "/"
                    $res += $_.Name + " - " + $temp[-1] + " - " + $_.Status + "<br>"
					write-host "YES"
                }
                else {
                    $res += $_.Name, $_.Mailbox.Name, $_.Status -join " - "
					$res += "<br>"
					write-host "NO"
                }
            }
        return $res    
    }

    
# Выводим заданный тект
    function OK {
        Write-Host "Выполнено!
        "
    }

# Получаем статистику ящиков с учётом фильтрации
    function getUsers { 
    param ($umbx)                 
        $var = $umbx | Get-MailboxStatistics | ScopeSize
        return $var
    }


# Ищем строку с названием "Microsoft.Exchange.Management.PowerShell.E2010" в установленных оснастках
    function GetExchShell{
        param ($p1, $p2)
        $ob = $false
            for ($i=0; $i -lt $p1.Count; $i++){
                if ($p1[$i] -eq $p2 ) {            
                    $ob = $true               
                }
            }
        return $ob   
    }


# Отправка уведомления
function sendMail {
    param ($eBody)
    $eBody += $emailFoter
    Send-MailMessage -To $emailTo -From $emailFrom -Subject $emailSubject -Body $eBody -BodyAsHTML -SmtpServer $emailSmtpServer -Encoding $encoding

}





# ------------------   Тело скрипта   -----------------------


# проверяем доступность пути для выгрузки и отправляем в очередь
    $TestBackupPath = Test-Path $path
    if (!($TestBackupPath)) {
        Write-Host "Каталог для выгрузки недоступен. Дальнейшее выполнеие невозможно
        отправляем уведомление"
        $emailSubject = $emailSubjectError
        sendMail $errorBody
        exit
    }


# подключаем инструменты Exchange      
    $pssnapName = (Get-PSSnapin).Name
    write-host "Проверяем доступность установленной оснастки ExchangeMangmentShell"
    $obName = GetExchShell $pssnapName $exchSnapName
    OK

    if ($obName -eq $false){
        $pssnapName = (Get-PSSnapin -Registered).Name
        $obName = GetExchShell $pssnapName $exchSnapName

        if ($obName){
        # Добавляем оснастку Exchange (если установлена консоль ExchangeMangmentShell)
            Write-Host "Добавляем оснастку Exchange"
	        $exShell = $true

            Add-PSSnapin $exchSnapName 
            . $env:ExchangeInstallPath\bin\RemoteExchange.ps1 
            $autoConnect = Connect-ExchangeServer -auto
            $obName = GetExchShell $pssnapName $exchSnapName
               if($obName){
                  Write-Host "Оснастка добавлена"
               }
               else{        
                   Write-Host "Подключение не удалось"
                   exit
               }
         } 
         else {
            # устанавливаем сессию с Exchange (если консоль ExchangeMangmentShell не установлена)
            Write-Host "Оснастка не установлена. Устанавливаем сессию с Exchange"
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch1/powershell/ -Authentication Kerberos
            Import-PSSession $session
            Write-Host "Сессия установлена
            "
         }
    }
    else {
        Write-Host "Оснастка уже подключена"
    }

# Фильтр для выборки ящиков со значением объёма в указанном промежутке
    if ($exShell) {
        filter ScopeSize {if (( $_.totalitemsize -ge $minMailboxSize ) -and ( $_.totalitemsize -le $maxMailboxSize )){$_}}  
    }
    else { 
        filter ScopeSize {if (($_.totalitemsize -match "[(]((\d+[,]?)*)") -and ([int64]($Matches[1]) -lt $maxMailboxSize) -and ([int64]($Matches[1]) -gt $minMailboxSize)) {$_} }
    }


# Проверяем доступность серверов почтовых ящиков

    # Получаем список серверов почтовых ящиков
    #$AllMailboxServer = Get-MailboxServer | sort name -Descending
    $AllMailboxServer = Get-MailboxServer | Where-Object {Test-Connection $_ -Count 2} | sort name

    Write-Host "Получаем список доступных серверов почтовых ящиков"
    if (!($AllMailboxServer)){
        Write-Host "Серверы почтовых ящиков недоступны"
        exit
    }
    else{
        Write-Host "Доступны следующие серверы почтовых ящиков: " 
        $AllMailboxServer | ForEach-Object {$_.Name}
    }
    OK

# Получаем статусы предыдущей операции
    Write-Host "Получаем статусы предыдущей операции"
    $PreList = Get-MailboxExportRequest

    $CompletedProcess = $PreList | Where-Object  {$_.status -eq "Completed"}
    $CompletedCount = $CompletedProcess.Count

    $FailedProcess = $PreList | Where-Object  {$_.status -eq "Failed"}
    $FailedCount = $FailedProcess.Count

    $CurrentProcess = $PreList | Where-Object  {$_.status -eq "InProgress"}
    $CurrentCount = $CurrentProcess.Count

    $QueuedProcess = $PreList | Where-Object  {$_.status -eq "Queued"}
    $QueuedCount = $QueuedProcess.Count

# Формируем списки со статусами из предыдущей операции
    $FailedSendMailboxList = getStatusMailbox $FailedProcess
    $ProgressSendMailboxList = getStatusMailbox $CurrentProcess
    $QueuedSendMailboxList = getStatusMailbox $QueuedProcess
    $CompletedSendMailboxList = getStatusMailbox $CompletedProcess
    

# Удаляем все запросы из очереди
    write-host "Удаляем все предыдущие запросы из очереди"
    Get-MailboxExportRequest | Remove-MailboxExportRequest -Confirm:$false
    OK

# Выбираем ящики у пользователей в области поиска
    Write-Host "Запрашиваем почтовые ящики со всех доступных серверов с учётом области поиска"
    $UserMailboxes = $AllMailboxServer | ForEach {Get-Mailbox -OrganizationalUnit $ExportUsersOU -RecipientTypeDetails UserMailbox -Server $_.Name}
    Write-Host "Получено почтовых ящиков:" $UserMailboxes.Count
""


# Получаем список пользователей с учётом фильтрации
    Write-Host "Получаем список пользователей с учётом фильтрации"
    $selectUsers = getUsers $UserMailboxes
    Write-Host "Число пользователей:" $selectUsers.Count
    ""

# Делаем выборку ящиков только тех пользователей, которые соответствую фильтру по объёму ящика
    Write-Host "Делаем выборку ящиков только тех пользователей, которые соответствуют фильтру по объёму ящика"
    $SelectMailbox = CompareObject $UserMailboxes $selectUsers Name DisplayName
    Write-Host "Выбрано почтовых ящиков с объёмом от $minScope Мб до $maxScope Мб:" $SelectMailbox.Count
    ""

if (!($SelectMailbox)){
    Write-Host "Почтовые ящики, с объёмом от $minScope Мб до $maxScope Мб не найдены"
    exit
}
else {
    $sizeBox = $SelectMailbox | Get-MailboxStatistics
    if($exShell){
        Write-Host "Объёмы полученных ящиков (по объёму ящика)"         
        $sizeBox | sort TotalItemSize | ft "DisplayName", "TotalItemSize" -AutoSize    
    }
    else{
        Write-Host "Объёмы полученных ящиков (по алфавиту)"
        $sizeBox | sort DisplayName | ft "DisplayName", "TotalItemSize" -AutoSize    
    }
}

# Отправляем полученные ящики на экспорт
    Write-Host "Добавляем в очередь заданий на экспорт"
        $SelectMailbox | ForEach-Object {        
        New-MailboxExportRequest -Name $_.Alias -Mailbox $_.Alias -FilePath "$path\$($_.DisplayName)_($($_.Alias)).pst"
        $SendMailboxList += " " + $_.DisplayName + "<br>"
        Start-Sleep -Milliseconds 200
        Write-Host $_.Displayname
    }
    OK
    
Write-Host "Итого: 
    Всего получено пользовательских почтовых ящиков:" $UserMailboxes.Count "
    Число пользователей с учётом фильтрации по области поиска и объёма ящика:" $selectUsers.Count "
    Выбрано и отправлено на экспорт почтовых ящиков с объёмом от" $minScope "Мб до" $maxScope "Мб:" $SelectMailbox.Count "
    "
$currentTime.DateTime
$currentComp

# отпраляем оповещение
Write-Host "Отправляем оповещение на адрес" $emailTo

$countAllMailbox = $UserMailboxes.Count
$countSelusers = $selectUsers.Count
$countSelmailbox = $SelectMailbox.Count


$emailBody += "
<h2>Завершено</h2>

<b>Сводка по результатам работы скрипта:</b><br>

Всего получено пользовательских почтовых ящиков - <b>$countAllMailbox</b><br>
Число пользователей с учётом фильтрации по области поиска и объёма ящика - <b>$countSelusers</b><br>
Выбрано и отправлено на экспорт почтовых ящиков с объёмом <b>от $minScope Мб до $maxScope Мб - $countSelmailbox</b><br>
Каталог для выгрузки: $path<br><br>

---------      <b>Отправлено ящиков: ($countSelmailbox штук)</b>      ---------<br>
    $SendMailboxList<br><br>

<h2>Статистика предыдущего запуска</h2>
---------      <b>Успешно завершено ($CompletedCount штук):</b>      ---------<br>
    $CompletedSendMailboxList<br><br>

---------      <b>Операции в процессе выполнения: ($CurrentCount штук)</b>      ---------<br>
    $ProgressSendMailboxList<br><br>

---------      <b>Завершено с ошибкой копирования ($FailedCount штук):</b>      ---------<br>
    $FailedSendMailboxList<br><br>

---------      <b>Задания в очереди копирования ($QueuedCount штук):</b>      ---------<br>
    $QueuedSendMailboxList<br><br>
"
  
# Отправляем письмо 
sendMail $emailBody
OK

# удаляем сессию Exchange (если она есть)
    if ($session){
        Write-Host "Удаляем сессию с Exchange"
        Remove-PSSession $session
    }

Write-Host $exShell

