Import-Module ActiveDirectory

$currentTime = Get-Date
$searchBase = "ou=Users, ou=moscow, dc=geops, dc=local"
$server = "srv.geops.local"
$currentComp = $env:COMPUTERNAME

# Флаг для добавленных пользователей
if (Test-Path variable:added) {Remove-Variable -Name added}
New-Variable -Name added -Value $false -Scope Global

# Флаг для удалённых пользователей
if (Test-Path variable:removed) {Remove-Variable -Name removed}
New-Variable -Name removed -Value $false -Scope Global


# Задаём параметры для отправки оповещения на E-mail
    $encoding = [System.Text.Encoding]::UTF8
    $emailSmtpServer = "exch1.geops.local"
    $emailFrom = "infohost@geops.ru"
    $emailTo = "support@geops.ru"
    $emailSubject = "Сортировка пользователей по группам доступа к устройствам"
    $emailBody = "<h2>Выполнено</h2>"
    $emailFoter = "<br><hr>
    Исполняющий компьютер: $currentComp<br>
    Дата и время выполнения: $currentTime<br>"

    if (Test-Path variable:addedUsersBody) {Remove-Variable -Name addedUsersBody}
    New-Variable -Name addedUsersBody -Value "<u><b>Пользователи добавленные в следующие группы:</b></u><br>" -Scope Global
    
    if (Test-Path variable:removedUsersBody) {Remove-Variable -Name removedUsersBody}
    New-Variable -Name removedUsersBody -Value "<br><u><b>Пользователи удаленные из следущих групп:</b></u><br>" -Scope Global


# Отправка уведомления
    function sendMail {
        param ($eBody)
        $eBody += $emailFoter
        Send-MailMessage -To $emailTo -From $emailFrom -Subject $emailSubject -Body $eBody -BodyAsHTML -SmtpServer $emailSmtpServer -Encoding $encoding
    }

# Добавляем строку
    function delRowBody {
        param ($name, $group)
            $global:removedUsersBody += $name
            $global:removedUsersBody += " удалён из группы "
            $global:removedUsersBody += $group
            $global:removedUsersBody += "<br>" 
    }

# Удаляем строку
    function addRowBody {
        param ($name, $group)
            $global:addedUsersBody += $name
            $global:addedUsersBody += " добавлен в группу "
            $global:addedUsersBody += $group
            $global:addedUsersBody += "<br>" 
    }

# Добавляем пользователя в группу
    function AddUser {
    param ($login, $group)
        Add-ADGroupMember -Identity "$group" -Members $login -Confirm:$false -Server $server        
    }

# Удаляем пользователей из группы
    function DelUser {
    param ($login, $group)
        Remove-ADGroupMember -Identity "$group" -Members $login  -Confirm:$false -Server $server        
    }


# Массив групп с чередованием ReadOnly-FullAccess
$GroupArray = ("GFI_ESEC_StorageDevices_ReadOnly", "GFI_ESEC_StorageDevices_FullAccess", 
                "GFI_ESEC_CdDvd_ReadOnly", "GFI_ESEC_CdDvd_FullAccess",
                "GFI_ESEC_Floppy_ReadOnly", "GFI_ESEC_Floppy_FullAccess" )

# Логика. Проверяем наличие пользователя в группах $GroupArray
    function UsersCompare {
    param ($UserName, $Array)               
        $ArrCount = $Array.Count
        for ($i=0; $i -lt $ArrCount; $i += 2){
            $ob = Get-ADUser -Identity $UserName  -Server $server -Properties memberOf
            [string]$currentName =  $UserName.Name
            [string]$currentGroup = $Array[$i] 
            If (!($ob.memberof -match $Array[$i+1] )){
                If (!($ob.memberof -match $Array[$i] )){                    
                    Write-Host "Добавляем пользователя" $UserName.Name "в группу" $Array[$i]
                    AddUser $UserName $Array[$i]
                    $global:added += 1
                    addRowBody $currentName $currentGroup
                }   
            }
            else{
                If ($ob.memberof -match $Array[$i] ){
                    Write-Host "Удаляем пользователя" $UserName.Name "из группы" $Array[$i]
                    DelUser $UserName $Array[$i]
                    $global:removed += 1
                    delRowBody $currentName $currentGroup
                }
            }
        }
    }

Get-ADUser -Filter * -SearchBase $searchBase | ForEach-Object {UsersCompare $_ $GroupArray}

if ($added -eq $false){
    Clear-Variable -Name addedUsersBody
}

if ($removed -eq $false){
    Clear-Variable -Name removedUsersBody
}

if(($added -eq $false) -and ($removed -eq $false)){
    $emailBody += "Изменений в сортировке пользователей нет"
}


# Отправляем оповещение о выполнении
$emailBody = $emailBody + $addedUsersBody + $removedusersBody

sendMail $emailBody