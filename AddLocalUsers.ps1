
$root = "C:\PS\000" # корневая директория для всех файлов скрипта

# Корневой каталог (если не существует) для сохранения текущих значений
    $currentDir = $root + "\current\"
    if ((Test-Path $currentDir) -eq $false)
    {
        New-Item -Path $root -Name current -ItemType "directory"
    }

    $currentGroupsFile = $currentDir + "currentGroups.txt"
    $currentUsersFile =  $currentDir + "currentUsers.txt"

# Каталог для списка членов групп
    $currentGroupMembersDir = $currentDir + "groups\"
    if ((Test-Path $currentGroupMembersDir) -eq $false)
    {
        New-Item -Path $currentDir -Name groups -ItemType "directory"
    }

# Каталог для списка вхождений в группы каждого пользователя
    $currentUsersGroupDir = $currentDir + "users\"
    if ((Test-Path $currentUsersGroupDir) -eq $false)
    {
        New-Item -Path $currentDir -Name users -ItemType "directory"
    }


# Получаем имена новых групп из файла
    $groupsFile = $root + "\new\groups.txt"
    $newGroups = Get-Content -Path $groupsFile

# Получаем имена новых пользователей из файла
    $usersFile = $root + "\new\users.txt"
    $newUsers = Get-Content -Path $usersFile

# Получим и сохраним текущий список пользователей
    $currentUsers = Get-LocalUser
    foreach ($curUser in $currentUsers){
        $curUser.Name | Out-File $currentUsersFile -Encoding default -Append
    }


# Получим и сохраним текущий список групп
    $currentGroups = Get-LocalGroup
    foreach ($curGroup in $currentGroups){
        $curGroup.Name | Out-File -FilePath $currentGroupsFile -Encoding default -Append
    }


#   ----------   БЛОК ФУНКЦИЙ   --------- #


# Получим имена новых пользователей из файла
function getNewUserNames{
param($users)
    $userNames = @()
    foreach ($user in $users){
        $usr = $user -split " "  |  Where { $_.length -gt 0 }
        $userNames += $usr[0]
    }    
    return $userNames
}


# Получим и сохраним в файл текущих членов каждой группы
    function getCurrentGroups{
    ForEach ($group in $currentGroups){
            $members = Get-LocalGroupMember -Group $group
            foreach ($member in $members){
                ($member.Name -split "\\")[1] | Out-File -FilePath ($currentGroupMembersDir + "\" + $group + ".txt") -Encoding default -Append
            }
        }
    }



# Получим и сохраним вхождения каждого пользователя в группы
    function getCurrentUsers{
        foreach ($user in Get-LocalUser){
            foreach ($LocalGroup in Get-LocalGroup){
                if (Get-LocalGroupMember $LocalGroup -Member $user –ErrorAction SilentlyContinue){
                    $LocalGroup.Name | Out-File -FilePath ($currentUsersGroupDir + $user.Name + ".txt") -Encoding default -Append
                }
            }
        }
    }



# Создаём группы из массива
    function AddNewGroups{
       param($groups)
       foreach ($group in $newGroups){
            New-LocalGroup $group
        }
    }
    

# Создаём пользователей из массива
    function AddNewUsers {
    param ($users)
        foreach ($user in $users){   
            $usr = $user -split " "  |  Where { $_.length -gt 0 }
            $userName = $usr[0]
            $userPassword = $usr[1] | ConvertTo-SecureString -AsPlainText -Force
            New-LocalUser -Name $UserName -Password $UserPassword -PasswordNeverExpires -UserMayNotChangePassword
        }
    }


# Удаляем новые созданные группы
    function RemoveNewGroups{
    param ($groups)
    foreach ($group in $groups){
            Remove-LocalGroup -Name $group
        }
    }



# Удаляем добавленных пользователей
function RemoveNewUsers{
param($users)   
    $usrNames = getNewUserNames $users
    foreach ($Name in $usrNames){     
        Remove-LocalUser -Name $Name
    }
}




# Добавляем пользователей в группы согласно записей в файлах
    foreach ($group in $newGroups){    
        $groupFile = $root + "\new\groups\" + $group + ".txt"
        if (Test-Path $groupFile){
            $userList = Get-Content $groupFile
            foreach ($userName in $userList){
                Add-LocalGroupMember -Group $group -Member $userName
            }
        }  
    }


getCurrentGroups
getCurrentUsers

AddNewGroups $newGroups
AddNewUsers $newUsers

RemoveNewGroups $newGroups
RemoveNewUsers $newUsers


