

Write-host "***************SDI W2K12 PRO****************"
write-host "servidores sin sesiones"

(Get-BrokerMachine -MaxRecordCount 2000 -DesktopGroupName "SDI W2K12 PRO" -InMaintenanceMode $FALSE -PowerState On| Where-Object {$_.sessionCount -eq "0"}).count

write-host "sesiones Totales"

(Get-BrokerMachine -MaxRecordCount 2000 -DesktopGroupName "SDI W2K12 PRO" -InMaintenanceMode $FALSE -PowerState On| Where-Object {$_.sessionCount -ne "0"}).SessionCount | Measure-Object -Sum |Select-Object Sum |ft -AutoSize
 

Write-host "***************SDI W2K12 CONTACT CENTER PRO****************"
write-host "servidores sin sesiones"
(Get-BrokerMachine -MaxRecordCount 2000 -DesktopGroupName "SDI W2K12 CONTACT CENTER PRO" -InMaintenanceMode $FALSE -PowerState On| Where-Object {$_.sessionCount -eq "0"}).count

write-host "sesiones Totales"
(Get-BrokerMachine -MaxRecordCount 2000 -DesktopGroupName "SDI W2K12 CONTACT CENTER PRO" -InMaintenanceMode $FALSE -PowerState On| Where-Object {$_.sessionCount -ne "0"}).SessionCount | Measure-Object -Sum |Select-Object Sum |ft -AutoSize

