# ��������� ���������� �� ���������
 Set-Location C:\
 
# ����� ����� ��� Get-Help
 Set-Alias HelpM� Get-Help
 
# ���������� ���� ������������������ �������� � �������
 Get-Pssnapin -Registered | Add-Pssnapin -Passthru -ErrorAction SilentlyContinue
 Get-Module -ListAvailable| Import-Module -PassThru -ErrorAction SilentlyContinue
 

# ����������� ����� ��� Exchange 
 $MailServer = "exch1.geops.local"

$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add(
    "����������� � Exchange (�������� �����������)", {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
        . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
		Connect-ExchangeServer �Server $MailServer
       #   Connect-ExchangeServer -auto
            },
    "Control+Alt+0"
)

# ��������� ������ � "�������������� ����������" �� ����������� � Exchange Server � ������� ������
$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add(
  "����������� � Exchange (�������� �� �����������)",
    {
        $s = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri http://$MailServer/PowerShell/ `
        -Authentication Kerberos

        Import-PSSession $s
    },
  "Control+Alt+Z"
)

# �������������� ���������� ����������� � Exchange Server � ������� ������
# $session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$MailServer/powershell -Credential (Get-Credential)
# $session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$MailServer/powershell -Authentication Kerberos
# Import-PSSession $session

# ���c��� ������
 Clear-Host
 
  
# ����������� ���� ��������
 Write-Host "����������� ����, ��� ������� ���� !!!"