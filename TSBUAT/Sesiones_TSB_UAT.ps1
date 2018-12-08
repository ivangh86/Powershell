# Recuperar la fecha y hora actual 
 $date = Get-date -Format yyyyMMdd 
 $date2 = Get-date -uformat "%d/%m/%Y %Hh %Mm"
 $ElapsedTime = [System.Diagnostics.Stopwatch]::Startnew()


 $smtpserver = "relay.adgbs.com"
 $from = "0901stservers@bancsabadell.com"
 $recipient = "ivan.gomez-hernandez@dxc.com"
 $subject = "Check_Users_TSB_$date2"
  



$Sesiones = import-csv -Delimiter ";" -Path \\capelo\soft_citrix\00_Monitorizacion\Scripts\POWER_ENUM_XA6\Live_info\Yesterday_Sessions_Stats.csv | where-object {($_.Date -like "*:00") -or ($_.Date -like "*:01")} | select Date,TSBU_PLT_CBS_Total,TSBU_PLT_CBS_Active,TSBU_PLT_CBS_Discon



#$Sesiones | Out-File -FilePath "C:\informe_usuarios\TSBUAT\Check_Users_TSB_UAT.txt" | export-csv -Path "C:\informe_usuarios\TSBUAT\Check_Users_TSB_UAT.csv"

$Sesiones | export-csv -Path "C:\informe_usuarios\TSBUAT\Check_Users_TSB_UAT.csv" -NoTypeInformation -Delimiter ";"

$attach = "C:\informe_usuarios\TSBUAT\Check_Users_TSB_UAT.csv"


#$attach =  "C:\informe_usuarios\TSBUAT\Check_Users_TSB_UAT.csv"

# Enviar el resultado per correo
# -cc bssitogocargwinIberial1@hpe.com
#BS_Wintel_CTX@dxc.com
#PUJOL.VICENC@sabis.tech
#C:\informe_usuarios\.\mailsend -f 0901stservers@bancsabadell.com -d adgbs.com -smtp relay.adgbs.com -t ivan.gomez-hernandez@dxc.com  -sub "Check_Users_TSB_$date2" -a "$HTML_File"
C:\informe_usuarios\.\mailsend -f 0901stservers@bancsabadell.com -d adgbs.com -smtp relay.adgbs.com -t  PUJOL.VICENC@sabis.tech,P4UK_Level2@sabadell.co.uk,GOMEZRI@sabis.tech -cc BS_Wintel_CTX@dxc.com -sub "Check_Users_TSB_UAT_$date2" -attach "$attach"




#Send-MailMessage -smtpServer $smtpserver -from $from -to $recipient -subject $subject  -body "$UltimaConsulta"   #-Attachments $file

#$destino = "\\capelo\soft_wintel\Citrix\Informes_Publicos\XA6\TSBUAT"
#move-Item "C:\informe_usuarios\uptime\Check_Devices_PVS.html" -Destination $destino -force