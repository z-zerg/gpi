# Автоматический перенос пользовательских папок с системного диска C: на диск D:. Миграция существующих данных
# Конечно было бы неплохо, чтобы на Вашем компьютере были установлены два дисковых накопителя. 
# Однако даже если у Вас один винчестер, который логически поделен на два раздела C: (система) и D: (данные), 
# я все равно рекомендую держать всю пользовательскую информацию на отдельном от системного разделе (хотя бы для того, 
# чтобы при необходимости отформатировать раздел с Windows, не нужно было задумываться на счет сохраненных документов, 
# изображений, видео, аудио и других важных файликов, сохраненных на Рабочем столе).
 



#   ПЕРЕД ПРИМЕНЕНИЕМ ОБЯЗАТЕЛЬНО ПРОВЕРИТЬ ПУТИ




if (Test-Path -LiteralPath 'D:\')
{
    $curUser = "$env:USERDOMAIN\$env:USERNAME"

    if (!(Test-Path -LiteralPath 'D:\_DESKTOP'))
        {
            Write-Host -ForegroundColor DarkGray "START: Migrate Desktop"

            New-Item -Path "D:\_DESKTOP" -ItemType Directory -Force | Out-Null
            Start-Process -FilePath "$env:SystemRoot\System32\xcopy.exe" -Wait -WindowStyle Minimized -ArgumentList """$env:USERPROFILE\Desktop\*"" ""D:\_DESKTOP"" /C /H /K /O /X /R /E /I /G /Q /Y"
            $getDir = Get-Item -Path "D:\_DESKTOP"
            $getDir.Attributes = "Hidden, System"

            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_DESKTOP"" /grant:r *S-1-5-18:(OI)(CI)(F) /inheritance:r"
            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_DESKTOP"" /grant ""$curUser"":(OI)(CI)(F) /inheritance:r"

            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'Desktop' -Value 'D:\_DESKTOP' | Out-Null
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Desktop' -Value 'D:\_DESKTOP' | Out-Null
        }


    if (!(Test-Path -LiteralPath 'D:\_DOWNLOADS'))
        {
            Write-Host -ForegroundColor DarkGray "START: Migrate Downloads"

            New-Item -Path "D:\_DOWNLOADS" -ItemType Directory -Force | Out-Null
            Start-Process -FilePath "$env:SystemRoot\System32\xcopy.exe" -Wait -WindowStyle Minimized -ArgumentList """$env:USERPROFILE\Downloads\*"" ""D:\_DOWNLOADS"" /C /H /K /O /X /R /E /I /G /Q /Y"
            $getDir = Get-Item -Path "D:\_DOWNLOADS"
            $getDir.Attributes = "ReadOnly"

            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_DOWNLOADS"" /grant:r *S-1-5-18:(OI)(CI)(F) /inheritance:r"
            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_DOWNLOADS"" /grant ""$curUser"":(OI)(CI)(F) /inheritance:r"

            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name '{374DE290-123F-4565-9164-39C4925E467B}' -Value 'D:\_DOWNLOADS' | Out-Null
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name '{374DE290-123F-4565-9164-39C4925E467B}' -Value 'D:\_DOWNLOADS' | Out-Null
        }


    if (!(Test-Path -LiteralPath 'D:\_DOCUMENTS'))
        {
            Write-Host -ForegroundColor DarkGray "START: Migrate Documents"

            New-Item -Path "D:\_DOCUMENTS" -ItemType Directory -Force | Out-Null
            Start-Process -FilePath "$env:SystemRoot\System32\xcopy.exe" -Wait -WindowStyle Minimized -ArgumentList """$env:USERPROFILE\Documents\*"" ""D:\_DOCUMENTS"" /C /H /K /O /X /R /E /I /G /Q /Y"
            $getDir = Get-Item -Path "D:\_DOCUMENTS"
            $getDir.Attributes = "ReadOnly"

            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_DOCUMENTS"" /grant:r *S-1-5-18:(OI)(CI)(F) /inheritance:r"
            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_DOCUMENTS"" /grant ""$curUser"":(OI)(CI)(F) /inheritance:r"

            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'Personal' -Value 'D:\_DOCUMENTS' | Out-Null
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Personal' -Value 'D:\_DOCUMENTS' | Out-Null
        }


    if (!(Test-Path -LiteralPath 'D:\_AUDIO'))
        {
            Write-Host -ForegroundColor DarkGray "START: Migrate Audios"

            New-Item -Path "D:\_AUDIO" -ItemType Directory -Force | Out-Null
            Start-Process -FilePath "$env:SystemRoot\System32\xcopy.exe" -Wait -WindowStyle Minimized -ArgumentList """$env:USERPROFILE\Music\*"" ""D:\_AUDIO"" /C /H /K /O /X /R /E /I /G /Q /Y"
            $getDir = Get-Item -Path "D:\_AUDIO"
            $getDir.Attributes = "ReadOnly"

            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_AUDIO"" /grant:r *S-1-5-18:(OI)(CI)(F) /inheritance:r"
            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_AUDIO"" /grant ""$curUser"":(OI)(CI)(F) /inheritance:r"

            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'My Music' -Value 'D:\_AUDIO' | Out-Null
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'My Music' -Value 'D:\_AUDIO' | Out-Null
        }


    if (!(Test-Path -LiteralPath 'D:\_PICTURES'))
        {
            Write-Host -ForegroundColor DarkGray "START: Migrate Pictures"

            New-Item -Path "D:\_PICTURES" -ItemType Directory -Force | Out-Null
            Start-Process -FilePath "$env:SystemRoot\System32\xcopy.exe" -Wait -WindowStyle Minimized -ArgumentList """$env:USERPROFILE\Pictures\*"" ""D:\_PICTURES"" /C /H /K /O /X /R /E /I /G /Q /Y"
            $getDir = Get-Item -Path "D:\_PICTURES"
            $getDir.Attributes = "ReadOnly"

            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_PICTURES"" /grant:r *S-1-5-18:(OI)(CI)(F) /inheritance:r"
            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_PICTURES"" /grant ""$curUser"":(OI)(CI)(F) /inheritance:r"

            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'My Pictures' -Value 'D:\_PICTURES' | Out-Null
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'My Pictures' -Value 'D:\_PICTURES' | Out-Null
        }


    if (!(Test-Path -LiteralPath 'D:\_VIDEO'))
        {
            Write-Host -ForegroundColor DarkGray "START: Migrate Videos"

            New-Item -Path "D:\_VIDEO" -ItemType Directory -Force | Out-Null
            Start-Process -FilePath "$env:SystemRoot\System32\xcopy.exe" -Wait -WindowStyle Minimized -ArgumentList """$env:USERPROFILE\Videos\*"" ""D:\_VIDEO"" /C /H /K /O /X /R /E /I /G /Q /Y"
            $getDir = Get-Item -Path "D:\_VIDEO"
            $getDir.Attributes = "ReadOnly"

            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_VIDEO"" /grant:r *S-1-5-18:(OI)(CI)(F) /inheritance:r"
            Start-Process -FilePath "$env:SystemRoot\System32\icacls.exe" -Wait -WindowStyle Minimized -ArgumentList """D:\_VIDEO"" /grant ""$curUser"":(OI)(CI)(F) /inheritance:r"

            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'My Video' -Value 'D:\_VIDEO' | Out-Null
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'My Video' -Value 'D:\_VIDEO' | Out-Null
        }
}
