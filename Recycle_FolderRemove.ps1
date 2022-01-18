# Очистка корзины
$folder = $env:RedirectFolders + '\' + $env:USERNAME
$recycle = $env:DesktopDir + '\$RECYCLE.BIN'

#del $folder -Recurse -Force -Confirm
$folder
$desktop

Get-ChildItem $recycle | Remove-Item