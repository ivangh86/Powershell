
asnp citrix*

$Sesiones = Get-XASession
write-host "************************************"
write-host "Usuarios conectados CBS"
($Sesiones | where-object {$_.servername -like "*BS*"} | Select-Object AccountName -Unique).count

write-host "Usuarios conectados SC"
($Sesiones | where-object {$_.servername -like "*SC*"} | Select-Object AccountName -Unique).count

write-host "Usuarios conectados SF"
($Sesiones | where-object {$_.servername -like "*IS*"} | Select-Object AccountName -Unique).count

write-host "Usuarios conectados CO"
($Sesiones | where-object {$_.servername -like "*CO*"} | Select-Object AccountName -Unique).count

write-host "Usuarios conectados CAV"
($Sesiones | where-object {$_.servername -like "*CAV*"} | Select-Object AccountName -Unique).count



write-host "************************************"

write-host "Sesiones Activas CBS"
($Sesiones| where-object {($_.State -eq "Active") -and ($_.servername -like "*BS*")}).count

write-host "Sesiones Activas SC"
($Sesiones| where-object {($_.State -eq "Active") -and ($_.servername -like "*SC*")}).count

write-host "Sesiones Activas SF"
($Sesiones| where-object {($_.State -eq "Active") -and ($_.servername -like "*IS*")}).count

write-host "Sesiones Activas COV"
($Sesiones| where-object {($_.State -eq "Active") -and ($_.servername -like "*COV*")}).count

write-host "Sesiones Activas CAV"
($Sesiones| where-object {($_.State -eq "Active") -and ($_.servername -like "*CAV*")}).count



write-host "************************************"

write-host "Sesiones Desconectadas CBS"
($Sesiones | where-object {($_.State -eq "Disconnected") -and ($_.servername -like "*BS*") }).count

write-host "Sesiones Desconectadas SC"
($Sesiones | where-object {($_.State -eq "Disconnected") -and ($_.servername -like "*SC*") }).count

write-host "Sesiones Desconectadas SF "
($Sesiones | where-object {($_.State -eq "Disconnected") -and ($_.servername -like "*IS*") }).count

write-host "Sesiones Desconectadas COV"
($Sesiones | where-object {($_.State -eq "Disconnected") -and ($_.servername -like "*COV*") }).count

write-host "Sesiones Desconectadas CAV"
($Sesiones | where-object {($_.State -eq "Disconnected") -and ($_.servername -like "*CAV*") }).count

write-host "************************************"