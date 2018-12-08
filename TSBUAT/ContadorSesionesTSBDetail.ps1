##########################################################################
#
# Autor: Iván Gomez
# Fecha: 21/09/2017
# Descripcion: Obtener Sesiones contra el entorno de TSB para posterior Report
#
#
#########################################################################


# Load only the snap-ins, which are used
if 
(

(Get-PSSnapin "Citrix.Broker.Admin.*" -EA silentlycontinue) -eq $null -or
(Get-PSSnapin "Citrix.Common.*" -EA silentlycontinue) -eq $null -or 
(Get-PSSnapin "Citrix.XenApp.*" -EA silentlycontinue) -eq $null
)
        
{
try { Add-PSSnapin Citrix* -ErrorAction Stop }
catch { write-error "Error Get-PSSnapin Citrix.Broker.Admin.* Powershell snapin"; Return }
}

$FechaActual = '{0:yyyy}' -f (Get-Date) + '_' + '{0:MM}' -f (Get-Date) + '_' + '{0:dd}' -f (Get-Date)
$Logfile = "\\capelo\soft_wintel\Citrix\Scripts\TSBUAT\UsersTSB_UAT_"+$FechaActual+".txt"


Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

#Servidores Entorno TSB UAT
$Servidores = "XATBSV0001D"#,"XATBSV0002D","XATBSV0003D","XATBSV0004D","XATBSV0005D","XATBSV0006D","XATBSV0007D","XATBSV0008D","XATBSV0101D","XATBSV0102D","XATBSV0103D","XATBSV0104D","XATBSV0105D","XATBSV0106D","XATBSV0107D","XATBSV0108D","XATBSV0109D","XATBSV0110D","XATBSV0111D","XATBSV0112D","XATBSV0113D","XATBSV0114D","XATBSV0115D","XATBSV0116D"
$Estados = "Active","Disconnected"

$Sesiones = Get-XASession | where-object {$Servidores -contains $_.ServerName -and $Estados -contains $_.State -and $_.SessionName -ne "Console"} | Sort-Object $_.ServerName


LogWrite "Estado ; LogOnTime ; AccountName ; ServerName ; ClientName ; ClientAddress ; App"

foreach ($Sesion in $Sesiones){


$Estado = $Sesion.State 
$Logon  = $Sesion.LogOnTime
$Cuenta = $Sesion.AccountName
$Server = $Sesion.ServerName
$Client = $Sesion.ClientName
$APP    = $Sesion.BrowserName
$IPClient = Get-XASession  -account $Cuenta  -Full | Where-Object {$_.ServerName -eq $Server -and $_.BrowserName -eq $APP} | select -expand "ClientIPv4" 

 #"$Sesion.State;$Sesion.LogOnTime;$Sesion.AccountName;$Sesion.ServerName;$Sesion.ClientName;$Sesion.BrowserName" 
 LogWrite "$Estado;$Logon;$Cuenta;$Server;$Client;$IPClient;$APP"
 "$Estado;$Logon;$Cuenta;$Server;$Client;$IPClient;$APP"

}
