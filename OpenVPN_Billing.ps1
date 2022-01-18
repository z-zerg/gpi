Add-Type -AssemblyName System.Web

$ServiceName = "OpenVPNService"
$ActionFile = "https://yadi.sk/d/ZbeyFIEmqMu7Ww"

function getContentFromYandex ()
{
    # $ud = [System.Web.HttpUtility]::UrlEncode($ActionFile)
    $url =  "https://cloud-api.yandex.net:443/v1/disk/public/resources/download?public_key=$ActionFile"

    $content = [string](Invoke-WebRequest (Invoke-RestMethod $url).href).ToString()
    if($content -ne ""){
        return $content
    }
    return $null    
}



# получим объект службы (если такая служба есть)
function getService($name=$ServiceName)
{
    return Get-Service -Name $ServiceName -ErrorAction SilentlyContinue    
}


$Service = getService
if($Service -eq $null){
        Write-Host "Cлужба $ServiceName не установлена"  
        exit
} 

# получим первую строку из файла action.txt как действие
[string]$ManageStatus = getContentFromYandex

# получим текущий статус службы
[string]$ServiceStatus = $Service.Status

# переводим тип запуска службы в отключено
function changeStartTypeService ($service, $type)
{
    Set-Service $service.Name -StartupType $type
}



switch ($ManageStatus){
    "ON" 
        {
            if($ServiceStatus -like 'Running')
            {
                changeStartTypeService $Service Automatic
            } 
            elseif ($ServiceStatus -like 'Stopped') 
            {
                # меняем тип запуска на автоматический и запускаем службу  
                Write-Host "меняем тип запуска на автоматический и запускаем службу"
                changeStartTypeService $Service Automatic
                Start-Service $Service.Name
            }
        }
    "OFF" 
        {
            if($ServiceStatus -like 'Running')
            {
                # останавливаем службу и меняем тип запуска на "Отключено"
                Write-Host "останавливаем службу и меняем тип запуска на Отключено"
                Stop-Service $Service.Name 
                changeStartTypeService $Service Manual
            } 
            elseif($ServiceStatus -like 'Stopped')
            {
                #  проверяем тип запуска и если нужно ставим на "Отключено"  
                Write-Host "Проверяем тип запуска и если нужно ставим на Вручную"
                changeStartTypeService $Service Manual
            }
        }
}