# принимаются входящие параметры Action и возможные значения (Add, Delete, Check)
param(
[string]$Action
)


# --------   Конфиг данные   ----------  


# папка где будем создавать ярлыки
$DefaultTargetFolder = $env:CommonDesktopDir

# Действие по умолчанию (создаём ярлык)
$DefaultAction = 'Check'

# Если в аргументе не передан Action, принимае по умолчанию
if (!$Action){
    $Action = $DefaultAction
} 

$Soft = @(
    'Trimble',
    'Civil3D_2017',
    'Civil3D_2019',
    'AutoCad_2019',
    'MicroStation_8i',
    'GlobalMapper_19',    
    'GlobalMapper_20', 
    'Outlook_2010'
    #'Calculator'
    'MapInfo',
    'MapInfo_10_5',
    'MicrosoftProject_2010'
)

$Calculator = @{
    'LinkName' = 'Калькулятор'
    'Command' = "$env:windir\system32\calc.exe"
    'WorkFolder' = $null
    'Arg' = $null
    'Icon' = $null
    'Desc' = 'Программа для математических операций'
}

$Civil3D_2017 = @{
    'LinkName' = 'Civil 3D 2017'
    'WorkFolder' = 'C:\Program Files\stdProgramms\Autodesk\AutoCAD 2017\UserDataCache\'
    'Arg' = '/ld "C:\Program Files\stdProgramms\Autodesk\AutoCAD 2017\\AecBase.dbx" /p "<<C3D_Metric>>"  /product "C3D" /language "ru-RU"'
    'Command' = 'C:\Program Files\stdProgramms\Autodesk\AutoCAD 2017\acad.exe'
    'Icon' = "%ProgramFiles%\stdProgramms\Autodesk\AutoCAD 2017\C3D\C3D.ico"
    'Desc' = $null
}


$Civil3D_2019 = @{
    'LinkName' = 'Civil 3D 2019'
    'WorkFolder' = 'C:\Program Files\Autodesk\AutoCAD 2019\UserDataCache\'
    'Arg' = '/ld "C:\Program Files\Autodesk\AutoCAD 2019\\AecBase.dbx" /p "<<C3D_Metric>>"  /product "C3D" /language "ru-RU"'
    'Command' = 'C:\Program Files\Autodesk\AutoCAD 2019\acad.exe'
    'Icon' = "%ProgramFiles%\Autodesk\AutoCAD 2019\C3D\C3D.ico"
    'Desc' = $null
}

$AutoCad_2019 = @{
    'LinkName' = 'AutoCad 2019'
    'WorkFolder' = 'C:\Program Files\Autodesk\AutoCAD 2019\UserDataCache\'
    'Arg' = '/product ACAD /language "ru-RU'
    'Command' = 'C:\Program Files\Autodesk\AutoCAD 2019\acad.exe'
    'Icon' = "C:\Program Files\Autodesk\AutoCAD 2019\acad.exe"
    'Desc' = $null
}

$Trimble = @{
    'LinkName' = 'Trimble Business Center'
    'Command' = 'C:\Program Files\stdProgramms\Trimble\Trimble Business Center\TrimbleBusinessCenter.exe'
    'WorkFolder' = $null
    'Arg' = $null
    'Icon' = $null
    'Desc' = $null
}

$GlobalMapper_19 = @{
    'LinkName' = 'Global Mapper 19'
    'Command' = 'C:\Program Files (x86)\stdProgramms\GlobalMapper19_64bit\global_mapper.exe'
    'WorkFolder' = $null
    'Arg' = $null
    'Icon' = $null
    'Desc' = $null
}

$GlobalMapper_20 = @{
    'LinkName' = 'Global Mapper 20'
    'Command' = 'C:\Program Files (x86)\stdProgramms\GlobalMapper20_64bit\global_mapper.exe'
    'WorkFolder' = $null
    'Arg' = $null
    'Icon' = $null
    'Desc' = $null
}

$Outlook_2010 = @{
    'LinkName' = 'Outlook 2010'
    'Command' = 'C:\Program Files (x86)\stdProgramms\Microsoft Office\Office14\OUTLOOK.EXE'
    'WorkFolder' = 'C:\Program Files (x86)\stdProgramms\Microsoft Office\Office14'
    'Arg' = $null
    'Icon' = $null
    'Desc' = 'Программа для обмена электронными сообщения e-mail'
}

$MapInfo = @{
    'LinkName' = 'Mapinfo Pro 15'
    'Command' = 'C:\Program Files (x86)\stdProgramms\MapInfo\Professional\MapInfow.exe'
    'WorkFolder' = $null
    'Arg' = $null
    'Icon' = $null
    'Desc' = $null
}

$MapInfo_10_5 = @{
    'LinkName' = 'MapInfo Pro 10.5'
    'Command' = 'C:\Program Files (x86)\stdProgramms\MapInfo\Professional 10.5\MapInfow.exe'
    'WorkFolder' = $null
    'Arg' = $null
    'Icon' = $null
    'Desc' = $null
}

$MicroStation_8i = @{
    'LinkName' = 'Microstation 8i'
    'Command' = 'C:\Program Files (x86)\stdProgramms\Bentley\MicroStation V8i (SELECTseries)\MicroStation\ustation.exe'
    'WorkFolder' = $null
    'Arg' = $null
    'Icon' = $null
    'Desc' = $null
}


$MicrosoftProject_2010 = @{
    'LinkName' = 'Microsoft Project 2010'
    'Command' = 'C:\Program Files (x86)\stdProgramms\Microsoft Office\Office14\WINPROJ.EXE'
    'WorkFolder' = $null
    'Arg' = $null
    'Icon' = $null
    'Desc' = $null
}



# --------   END Конфиг данные   ----------  

# Добавляет ярлык на рабочий стол (по умолчания создаётся в папке для всех пользователей)
function AddToDesktop {
param(
        [hashtable]$programm=$null, 
        [string]$Target
)    

    if ($Target -like 'User'){
        $Target = $env:DesktopDir            
    } else {
        $Target = $DefaultTargetFolder
    }

    
    # проверим есть ли исполняемый файл
    if (Test-Path $programm.Command){
        
                # Заполним свойства объекта (поля ярлыка) полученными значениями
                if ($prog.LinkName -ne $null){
                    $Name = $programm.LinkName
                }
        
                if($prog.Desc -ne $null){
                    $Desc = $programm.Desc
                }

                if($prog.Arg -ne $null){
                    $Arg = $programm.Arg
                }
        
                if ($prog.Icon -ne $null){
                    $Icon = $programm.Icon
                }
        
                if($prog.Command -ne $null){
                    $Command = $programm.Command
                }
        
                if($prog.WorkFolder -ne $null){
                    $WorkFolder = $programm.WorkFolder
                }

                # имя файла для ярлыка
                $LnkFile = "$Target\$Name.lnk"                

                # Создадим объект ярлыка
                    $shell = New-Object -ComObject WScript.Shell    
                    $lnk = $shell.CreateShortcut($LnkFile)

                # Заполним свойства объекта (поля ярлыка) полученными значениями
                    if ($Desc -ne $null){
                        $lnk.Description = "$Desc"
                    }
                    if ($Arg -ne $null){
                        $lnk.Arguments = "$Arg" 
                    }

                    if ($Icon -ne $null){
                        $Lnk.IconLocation = $Icon
                    }
            
                    $lnk.WorkingDirectory = $WorkFolder

                    $lnk.TargetPath = "$Command"     
        
        
                # Сохраняем ярлык
                    $lnk.Save()
            
            } else {
        return 'Отсутствует исполняемый файл для программы - ' + $prog.LinkName        
    }
}

# Удаляет ярлык с рабочего стола текущего пользователя
function RemoveFromDesktop {
param(
        [hashtable]$programm=$null, 
        [string]$Target
)    

    if ($Target -like 'User'){
        $Target = $env:DesktopDir            
    } else {
        $Target = $DefaultTargetFolder
    }

    $ShortcutFile = $Target + '\' + $programm.LinkName + '.lnk' 
    if (Test-Path $ShortcutFile){
         Remove-Item $ShortcutFile -Force -ErrorAction SilentlyContinue
    }
}


foreach ($Name in $Soft) {
    if ((Get-Variable | Where {$_.Name -eq "$Name"})){
       $prog = (Get-Variable -Name $Name).Value
       $ShortcutFile = "$env:DesktopDir\" + $prog.LinkName + '.lnk' 

       switch ($Action)
       {
           Add{     
               AddToDesktop $prog 
                
           }

           Delete{
               RemoveFromDesktop $prog
           }

           Check{
                # Если файл существует
                if (Test-Path $prog.Command){ 
                    # если ярлык отсутствует
                    if (!(Test-Path $ShortcutFile)){
                        AddToDesktop $prog
                    }
                # иначе если файл недоступен
                } else {
                    RemoveFromDesktop $prog
                    ## если ярлык существует
                    #if (Test-Path "$ShortcutFile"){
                    #    RemoveFromDesktop $prog
                    #}
                }
           }
       }
    }  
}


