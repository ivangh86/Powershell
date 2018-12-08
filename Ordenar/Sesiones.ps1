#Ejecutar en un SDK de BS
Write-host "Consultando Granja Citrix BS"

$result = import-csv -Delimiter ";" -Path \\capelo\soft_citrix\00_Monitorizacion\Scripts\POWER_ENUM_XA6\Live_info\Live_Sessions_Stats.csv | select ICA_G_Total

write-host "Sesiones en la granja: $($result.ICA_G_total)"