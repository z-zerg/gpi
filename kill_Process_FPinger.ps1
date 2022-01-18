$comps=("geo-103","geo-120", "geo-134")
$procName = "FPinger"
filter compOnline { if (Test-Connection -ComputerName $_ -Delay 1 -Count 2 -ErrorAction SilentlyContinue) {$_} }
filter FPinger {if($_.Name -match $procName) {$_.Terminate()} }


$comps | compOnline | foreach { Get-WmiObject -ClassName Win32_Process -ComputerName $_ }| FPinger
