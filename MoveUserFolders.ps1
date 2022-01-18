# Автоматический перенос пользовательских папок с системного диска C: на диск D:. Миграция существующих данных
# Конечно было бы неплохо, чтобы на Вашем компьютере были установлены два дисковых накопителя. 
# Однако даже если у Вас один винчестер, который логически поделен на два раздела C: (система) и D: (данные), 
# я все равно рекомендую держать всю пользовательскую информацию на отдельном от системного разделе (хотя бы для того, 
# чтобы при необходимости отформатировать раздел с Windows, не нужно было задумываться на счет сохраненных документов, 
# изображений, видео, аудио и других важных файликов, сохраненных на Рабочем столе).
 

#   ПЕРЕД ПРИМЕНЕНИЕМ ОБЯЗАТЕЛЬНО ПРОВЕРИТЬ ПУТИ


# Указываем имена папок, которые хотим перенести
# Нужные папки для переноса раскомментировать
$MovedFolder = @(
    'Desktop'  ,    
    'Downloads' ,
    'Music'    ,
    'Favorites',
    'Pictures' ,
    'Contacts' ,
    'Video'    ,
    'Links'    ,
    'Documents'
)


$DiskProfile = 'D:'          # Букву диска, где будет хранится папка с основными папками пользователя
$AllUsersDir = 'USER_S'    # Здесь создаётся папка папка пользователя с перенесёнными папками




##############################################
#####                                    #####
#####  Дальнейший код НЕ ИЗМЕНЯТЬ !!!    #####
#####                                    #####
##############################################



Clear-Host

# проверим есть вообще диск, на который собираемся переносить
if (Test-Path $DiskProfile){

    Write-Host "При дальнейшем выполнеии программы данные из следующих папок будут перенесены в новое место `n"
    foreach ($mFolder in $MovedFolder){
        Write-Host $mFolder
    }

    Read-Host "
    Для продолжения нажмите ENTER"

    $UserName = $env:USERNAME    # Имя текущего пользователя

    # Путь в реестре с настройками папок
        $ShellFolders     = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
        $UserShellFolders = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

    # Ключи в реестре с текущими путями к папкам
        $ShellProperty = Get-ItemProperty -Path $ShellFolders



    # Значением указаны значения соответствующих ключей в преестре
    $Alias =@{
        "Video"      = "My Video";
        "Documents"  = "Personal";
        "Pictures"   = "My Pictures";
        "Desktop"    = "Desktop";
        "Favorites"  = "Favorites";
        "Music"      = "My Music";
        "Contacts"   = "{56784854-C6CB-462B-8169-88E350ACB882}";
        "Downloads"  = "{374DE290-123F-4565-9164-39C4925E467B}";
        "Links"      = "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}";
    }

    # Свойства по умолчанию для создаваемых объектов
    $prop = @{
        "Name"     = "";
        "SrcPath"  = "";
        "DestPath" = "";
        "RegKey"   = "";
    }

    # Создадим пустой массив объектов
    $Folder = @{}

    foreach ($name in $MovedFolder){

       $Folder[$name] = New-Object -TypeName psobject -Property $prop
       $Folder[$name].Name         = $name    
       $Folder[$name].RegKey       = $Alias.$name
       $Folder[$name].SrcPath      = $ShellProperty.($Folder[$name].RegKey)
       $Folder[$name].DestPath     = $DiskProfile + "\" + $AllUsersDir + "\" + $UserName + "\" + $name     
   

       # Создадим конечную папку, если её нет
       if (!(Test-Path -Path $Folder[$name].DestPath)){
            #Write-Host "Создаём новую папку " $Folder[$name].DestPath
            New-Item -Path $Folder[$name].DestPath -ItemType "directory" -Force | Out-Null
       } else {
            Write-Host "Похоже что конечная папка уже существует и в ней могут содержаться нужные данные. `nДальнейшее выполнение невозможно!"
            Read-Host "Для выхода нажмите любую клавишу"
            Exit
       }

       # Переносим данные в новое место
       Write-Host "Переносим папку " $Folder[$name].Name " в новое место"
   
       $source = '"' + $Folder[$name].SrcPath + '"'
       Write-Host "Источник: " $source

       $dest   ='"' +  $Folder[$name].DestPath + '"'
       Write-Host "Назначение: " $dest

       # Составим строку с аргументами для команды копирования
       $arg =$null
       $arg = $source + " " + $dest  + " /C /E /I /G /F /H /R /K /Y /Z"

       # Запускаем процесс копирования
            Start-Process -FilePath "$env:SystemRoot\System32\xcopy.exe" -Wait -WindowStyle Hidden -ArgumentList $arg
      
       # Назначаем права на скопированную папку
           Write-Host "Установка прав доступа к папкам"   
           # для учётной записи Система       
               $argSystem =  $dest + " /grant:r *S-1-5-18:(OI)(CI)(F) /inheritance:r"
               Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList $argSystem

           # для учётной записи группы Администраторы
               # получаем SID группы Администратры
               $objUser = New-Object System.Security.Principal.NTAccount("Администраторы")
               $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier]) 
               $adminsSID = $strSID.Value

               $argAdmin =  $dest + " /grant:r *`"$adminsSID`":(OI)(CI)(F) /inheritance:r"
               Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList $argAdmin

           # для текущего пользователя
               $argCurrentUser = $dest + " /grant:r `"$UserName`":(OI)(CI)(F) /inheritance:r"  
               Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList $argCurrentUser

        # Переписываем ключи реестра
            Write-Host "Запись новых значений в реестр"
            Set-ItemProperty -Path $ShellFolders -Name $Folder[$name].RegKey -Value $Folder[$name].DestPath | Out-Null
            Set-ItemProperty -Path $UserShellFolders -Name $Folder[$name].RegKey -Value $Folder[$name].DestPath | Out-Null


        # Удаляем старые данные
           Write-Host "Удаляем старые данные"
           Write-Host "Удаляем папку " $Folder[$name].Name
           Write-Host $Folder[$name].SrcPath
         
           $cmd = $env:SystemRoot  + "\System32\cmd.exe"
           $command = "rmdir /S /Q " + $Folder[$name].SrcPath
           $arg = @("/C", $command)
           Start-Process -FilePath $cmd -Wait -NoNewWindow -ArgumentList $arg

        Write-Host "Выполнено  `n"
    }

} else {
    clear-Host
    Write-Host "Похоже что диска, куда будут переноситься данные, не существует. `n`nДальнейшее выполнение программы невозможно `n"
    Read-Host "Для выхода из программы нажните любую клавишу"
    Exit
}
