$userName = "lykov"
$userProfilesFolder = "\\geops.local\admin\UserProfiles\"
$path_dest = "D:\"
$dest_folder = $path_dest + $userName
$fromFolder = $userProfilesFolder+$userName

#New-Item -Name $userName -Path $path_dest -ItemType directory

$aclUser = Get-Acl $dest_folder

(Get-Item $fromFolder).GetAccessControl('Access') | Set-Acl $path_dest
Set-Acl -Path ($path_dest+$userName) -AclObject $acl

$proc = Start-Process -FilePath "C:\Windows\System32\xcopy.exe"  -ArgumentList "D:\000 D:\001 /E /H /I /O" -PassThru -Wait
$proc.HasExited