# Отключаем учётные записи в заданной OU
# ----------------------------------------
$Log = "D:\SheduleScript\DisableUsersLog.txt"

$filter = 'OU=DisabledUsers, DC=geops, DC=local'

$Users = Get-ADUser -Filter * -SearchBase $filter | sort

$date = Get-Date

$data = ($date.Year).ToString() + "." + ($date.Month).ToString() + "." + ($date.Day).ToString() + " - "

$Users | ForEach-Object {
            if ($_.Enabled -like 'True'){
                $str = $data + $_.Name + " - Отключена"
                $str | Out-File -Encoding utf8 -FilePath $Log -Append
                Set-ADUser -Identity $_.SamAccountName -Enabled $false
            }
        }

