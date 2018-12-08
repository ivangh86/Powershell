# Autor:       Iván Gómez
# Fecha:       09/08/2017
# Descripción: Contador Sesiones Usuarios "Letra U" contra PROTEO4UK
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 

asnp Citrix*

$FechaActual = '{0:yyyy}' -f (Get-Date) + '_' + '{0:MM}' -f (Get-Date) + '_' + '{0:dd}' -f (Get-Date)
$Logfile = "\\capelo\soft_wintel\Citrix\Scripts\TSBUAT\UsersTSB_UAT_Proteo4UK_"+$FechaActual+".txt"


Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

$date = Get-Date -UFormat %Hh:%Mm:%Ss
$SesionesTotal =  (Get-XASession | where-object {$_.BrowserName -eq "TSB UAT Proteo4UK"}).count
$SesionesActivas =  (Get-XASession | where-object {($_.BrowserName -eq "TSB UAT Proteo4UK") -and ($_.state -eq "Active") }).count
$SesionesDiscon =  (Get-XASession | where-object {($_.BrowserName -eq "TSB UAT Proteo4UK")  -and ($_.state -ne "Active") }).count
$SesionesTotalU =  (Get-XASession | where-object {($_.BrowserName -eq "TSB UAT Proteo4UK") -and ($_.AccountName -like "ADGBS\U*")}).count
$SesionesActivasU =  (Get-XASession | where-object {($_.BrowserName -eq "TSB UAT Proteo4UK") -and ($_.AccountName -like "ADGBS\U*") -and ($_.state -eq "Active") }).count
$SesionesDisconU =  (Get-XASession | where-object {($_.BrowserName -eq "TSB UAT Proteo4UK") -and ($_.AccountName -like "ADGBS\U*") -and ($_.state -ne "Active") }).count


#Write-Host "Fecha;Total;Sesiones Activas;Sesiones Disconected"

LogWrite "$date;$SesionesTotal;$SesionesActivas;$SesionesDiscon;$SesionesTotalU;$SesionesActivasU;$SesionesDisconU" 