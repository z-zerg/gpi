# Указываем путь для выгрузки PST файлов
    $path = "\\geops.local\admin\BackupMailbox\"

# Указываем статус ящика (Completed, Failed, )
#$statusMailbox = "Completed"
$statusMailbox = "Failed"

# получаем ящики со статусом
$StatusMailbox = Get-MailboxExportRequest -Status $statusMailbox
$StatusMailboxCount = $StatusMailbox.Count
Write-Host "Число ящиков со статусом Failed: " $StatusMailboxCount

# получаем алиасы
$Alias = @()
$Alias = $statusMailbox | ForEach-Object {$_.Name}
$UserMailbox = $Alias | ForEach-Object {Get-Mailbox -Identity $_}


# получаем полное имя ящика
$FullName = @()
$FullName = $UserMailbox | ForEach-Object{$_.Name}

# Отправляем на экспорт полученные ящики

$UserMailbox | ForEach-Object {
    $fullPath = ""
    $TaskName = "new-" + $_.Alias
    $MailboxName = $_.Name
    $FullPath = $path + $_.Name + "_(" + $_.Alias + ").pst"
    Write-Host "Конечный файл:" $fullPath
    New-MailboxExportRequest -Mailbox $MailboxName -Name $TaskName -FilePath $FullPath
}
