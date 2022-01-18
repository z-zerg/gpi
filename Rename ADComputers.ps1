
# Получаем учетные данные
$cred = Get-Credential


# Шаблон нового имени
$tplName = "COMP-"



# Не переименовывать следующие компьютеры
$notNeed = (
"DC",
"ZZERG"

)


# Тестируем компьютер на доступность
function TestComp($comp)
{
    $test = get-WmiObject -Class Win32_OperatingSystem -ComputerName $comp -ErrorAction SilentlyContinue -Credential $cred -WarningAction SilentlyContinue

    if($test -ne $null)
    {
        return $true
    } 
    return $false
}


# Получаем имена всех компьютеров из AD
[System.Collections.ArrayList]$allComps = @()
$allComps = (Get-ADComputer -Credential $cred -Filter * -Properties * | select Name).Name


# формируем новые имена согласно шаблону имён
[System.Collections.ArrayList]$newName = @()
(1.. $allComps.Count) | ForEach-Object {$newName += $tplName + $_}


# убираем ненужные согласно списка имен, которые не нужно переименовывать
[System.Collections.ArrayList]$needComps = @()
$needComps = (Compare-Object -ReferenceObject $allComps -DifferenceObject $notNeed -IncludeEqual | where{$_.SideIndicator -eq "<="} | select InputObject).InputObject


# находим компьютеры, которые уже переимонованы
[System.Collections.ArrayList] $RenamedComps = @()
foreach ($comp in $needComps)
{
    if($comp -match $tplName){
        $RenamedComps += $comp
    }

}

# подготовим список компов для переименования с учётом тех, которые уже преименованы
[System.Collections.ArrayList] $workComps = @()
$workComps += $needComps | where {$_ -notlike "$tplName*"}

# подготовим список новых имён для переименования
[System.Collections.ArrayList] $exName = @()
$exName += (Compare-Object -ReferenceObject $RenamedComps -DifferenceObject $newName -IncludeEqual | where{$_.SideIndicator -ne "=="} | select InputObject).InputObject



# отправляем подготовленные списки на переименование
$errorComp = @()
$renamed = @()
if($workComps.Count -gt 0){
    foreach($comp in $workComps)
    {
        if(TestComp $comp)
        {        
            Rename-Computer -ComputerName $comp -NewName $exName[0] -DomainCredential $cred -WarningAction SilentlyContinue -Restart
            $renamed += $exName[0]
            $exName.RemoveAt(0)
        }
        else
        {
            $errorComp += $comp
        }
    }
}



# Дальше идет информационный блок, который необязателен. Сделан просто для красоты
cls
if ($renamed.Count -eq 0){
    Write-Host "Нечего переименовывать" -ForegroundColor DarkYellow
} else {

    if($errorComp.Count -gt 0)
    {
        Write-Host "Не удалось переименовать следующие компьютеры:`n" -ForegroundColor Magenta
        $errorComp
    }
    else
    {
        Write-Host "Ошибок нет"
        Write-Host "Переименовано копьютеров: " $renamed.count
    }

}



<#
    При возникновении ошибки подключения к WMI,
    читаем статью здесь
    https://www.10-strike.ru/networkinventoryexplorer/help/wmi.shtml
    https://docs.microsoft.com/ru-ru/windows/win32/wmisdk/troubleshooting-a-remote-wmi-connection

    Про сервер RPC недоступен читаем здесь
    https://nastroyvse.ru/opersys/win/sposoby-ustraneniya-oshibki-server-rpc-nedostupen.html
    http://pyatilistnik.org/the-rpc-server-is-unavailable/
    https://viarum.ru/server-rpc-nedostupen/

    Но сначала добавить в политиках предопределенное правило брандмауэра
    Инструментарий управления Windows (WMI)


#>