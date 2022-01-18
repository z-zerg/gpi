# Установка директорию по умолчанию
 Set-Location C:\
 
# Новый алиас для Get-Help
 Set-Alias HelpMе Get-Help
 
# Добавление всех зарегистрированных оснасток и модулей
 Get-Pssnapin -Registered | Add-Pssnapin -Passthru -ErrorAction SilentlyContinue
 Get-Module -ListAvailable| Import-Module -PassThru -ErrorAction SilentlyContinue
 

# Подключение табов для Exchange 
 $MailServer = "exch1.geops.local"

$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add(
    "Подключение к Exchange (оснастка установлена)", {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
        . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
		Connect-ExchangeServer –Server $MailServer
       #   Connect-ExchangeServer -auto
            },
    "Control+Alt+0"
)

# Добавляем кнопку в "Дополнительные компоненты" на подключение к Exchange Server в текущую сессию
$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add(
  "Подключение к Exchange (оснастка НЕ установлена)",
    {
        $s = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri http://$MailServer/PowerShell/ `
        -Authentication Kerberos

        Import-PSSession $s
    },
  "Control+Alt+Z"
)

# Автоматическое добавление подключения к Exchange Server в текущую сессию
# $session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$MailServer/powershell -Credential (Get-Credential)
# $session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$MailServer/powershell -Authentication Kerberos
# Import-PSSession $session

# Очиcтка экрана
 Clear-Host
 
  
# Приветствие себя любимого
 Write-Host "Приветствую тебя, мой дорогой друг !!!"