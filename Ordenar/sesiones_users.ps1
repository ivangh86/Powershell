

write-host "Usuarios conectados"
Get-XASession | Select-Object AccountName -Unique


write-host "Sesiones Activas"
(Get-XASession | where-object {$_.State -eq "Active"}).count


write-host "Sesiones Desconectadas"
(Get-XASession | where-object {$_.State -eq "Disconnected"}).count