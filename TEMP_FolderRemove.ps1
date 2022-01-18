# Очистка файлов в папке (передаётся параметром)
$folder = $env:TEMP
# $folder = args[0]
del $folder -Recurse -Force