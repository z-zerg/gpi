# период архивации по умолчанию
$DefaultPeriod = $null

# Корневая папка с базами
$sourceFolder = "D:\1c\"

# Папка для хранения резервных копий
$backupFolder = "D:\1c_backup\"

# принимается параметр Period и возможные значения (Day, Week, Month)
$Periods = @{
    "Day" = 'every_day'
    "Dayly" = 'every_day'
    "Week" = 'every_week'
    "Weekly" = 'every_week'
    "Month" = 'every_month'
    "Monthly" = 'every_month'
}


# Если аргумент не передан , принимается значение по умолчанию
if($args -ne $null)
{
    $period = $Periods[$args[0]]
}
else
{
    $period = $DefaultPeriod
}



# ==========================================================

# Получаем текущую дату в нужном формате для наименования папки
[string]$currentDate = Get-Date -Format yyyy.MM.dd

# Путь к консольному архиватору 7zip
$7zip = "C:\Program Files\7-Zip\7z.exe"
if (!(Test-Path -Path $7zip)){
    Write-Host "Архиватор не найден" -ForegroundColor Magenta
    Break
}


# Список баз для архивации из папки источника
$folderBase = Get-ChildItem -Path $sourceFolder -Directory

# Перебираем каталоги и архивируем
foreach($base in $folderBase){

    # Имя архива 
    if($period -ne $null)
    {
        $archivName = $backupFolder + $base.Name + '\' + $period + '\' +  $currentDate + '_' + $base.Name + '.zip'
    }
    else
    {
        $archivName = $backupFolder + $base.Name + '\' + $currentDate + '_' + $base.Name + '.zip'
    }
    
    # устанавливаем путь к текущей базе
    #$source = $base.FullName + "\*"
    $source = $base.FullName

    # Создаем архив каталога с базой (если такого нет)   
    if(!(Test-Path -Path $archivName)) 
    {
        &$7zip a -tzip -ssw -mx5 $archivName $source
    }    
}

