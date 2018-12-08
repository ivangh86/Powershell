
# Autor:       Iván Gómez
# Fecha:       09/08/2017
# Descripción: Envio de informe Sesiones Usuarios "Letra U" contra PROTEO4UK
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # 

$FechaActual = '{0:yyyy}' -f (Get-Date) + '_' + '{0:MM}' -f (Get-Date) + '_' + '{0:dd}' -f (Get-Date)



$Logfile = "\\capelo\soft_wintel\Citrix\Scripts\TSBUAT\UsersTSB_UAT_"+$FechaActual+".txt"
$CSVfile = "C:\informe_usuarios\TSBUAT\UsersTSB_UAT_"+$FechaActual+".csv"


#$Header = "Hora;Total;SesionesActivas;SesionesDisconnected;TotalU;SesionesActivasU;SesionesDisconnectedU"


echo "sep=;" > $CSVfile
#$Header >> $CSVfile


$Content = Get-Content $Logfile

$Content | Out-File -Append $CSVfile


# Enviar el resultado per correo
# -cc bssitogocargwinIberial1@hpe.com
#BS_Wintel_CTX@dxc.com
#PUJOL.VICENC@sabis.tech

C:\informe_usuarios\MailSend\mailsend.exe -f 0901stservers@bancsabadell.com -d adgbs.com -smtp relay.adgbs.com -t ramon.gobern-carrillo@hpe.com -cc ivan.gomez-hernandez@dxc.com -sub "UsersTSB_UAT_$FechaActual" -a  "$CSVfile" -y "csv"
#C:\informe_usuarios\MailSend\mailsend.exe -f 0901stservers@bancsabadell.com -d adgbs.com -smtp relay.adgbs.com -t  PUJOL.VICENC@sabis.tech,P4UK_Level2@sabadell.co.uk,GOMEZRI@sabis.tech  -cc BS_Wintel_CTX@dxc.com -sub "UsersTSB_UAT_Proteo4UK_$FechaActual" -a  "$CSVfile" -y "csv"



