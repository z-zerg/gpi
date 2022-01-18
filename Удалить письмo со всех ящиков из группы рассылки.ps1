@'
    Работает только на машине, где установлен Exchange
    либо запускать из консоли Exchange
'@


# *********        ИЗМЕНИТЬ ПЕРЕД ЗАПУСКОМ     ************************

$ADMIN = 'geops\trushkov' # учётка с полными правами в Exchange
$DistGroup = 'info' # группа рассылки


# параметры для запроса
$PARAM = @{
    
    TO = '' # адрес, на который прислали    
    FROM = 'info@s2.red47.ru' # адрес от кого получили     
    SUBJECT = 'Сотрудничество' # тема удаляемого письма    
    RECIEVED = '' # когда получено в формате 07.11.2019
}


# ********************    КОНЕЦ БЛОКА ДЛЯ ИЗМЕНЕНИЙ   *********************


# ********************    БЛОК ФУНКЦИЙ   *********************

# Перевод ключей на русский (Для русской версии Exchange)
$WORDS = @{
    TO = 'кому:'
    FROM = 'откого:'
    RECIEVED = 'получено:'
    SUBJECT = 'тема:'
}

# Удаляем полные права админа $ADMIN с почтового ящика $user
function RemovePermission{
param($user)
    Remove-MailboxPermission -Identity $user -User $ADMIN -InheritanceType All -AccessRights FullAccess -Confirm:$false
    Remove-MailboxPermission -Identity $user -User $ADMIN -Deny -AccessRights FullAccess -Confirm:$false
}

# Даём полные права на ящик $user для администратора $ADMIN
function AddPermission{
param($user, $access = 'FullAccess')

    Add-MailboxPermission -Identity $user.Alias -User $ADMIN -AccessRights $access
}

# находим ящики, входящие в группу $DistGroup
function getDistGroupUsers {
param()
    $users = Get-ADGroupMember -Identity $DistGroup -Server srv.geops.local | select name, SamAccountName | sort name
    #$users
    $MBox = $users | ForEach-Object {
            Get-Mailbox -Identity $_.name -ResultSize unlimited    
        }
    return $MBox
}

# формируем запрос для поиска из параметров $PARAM
function getSearchQuery {    

    $SearchQuery = @{}

    ForEach ($key in $PARAM.Keys) 
    {
        if ((($PARAM[$key]).Length -ne 0)){
            
            $SearchQuery[$WORDS[$key]] = $PARAM[$key]
            
        }    
    }

    $Array = @()

    foreach ($Query in $SearchQuery){

        foreach($key in $Query.Keys){

            $Array += $key + '"' + $SearchQuery[$key] + '"'   
              
        }
    }
    return ($Array -join ', ')
}


# ********************    КОНЕЦ БЛОКА ФУНКЦИЙ   *********************

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
Import-Module ActiveDirectory

$QUERY = getSearchQuery

$mailBoxes = getDistGroupUsers

$mailBoxes | ForEach-Object {AddPermission $_}

$mailBoxes | ForEach {Search-Mailbox -Identity $_ -SearchQuery $QUERY -DeleteContent -Force}

$mailBoxes | ForEach-Object {RemovePermission $_}

