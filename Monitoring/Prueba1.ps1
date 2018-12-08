$date = Get-Date -format "HH_mm_ss_dd_MM_yyyy"
$currentDir = Split-Path $MyInvocation.MyCommand.Path
$filetxt = New-Item $currentDir\$date"_Report_XenApp.csv" -ItemType file
$filetxt
write-host "Ruta:" $currentDir\$filetxt
