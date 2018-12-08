###############################################################################
#
#
# Autor: Ivan Gomez
# Fecha: 20/04/2018
# Descripcion: Revisa las sesiones totales y a que usuarios corresponden.
#
###############################################################################

asnp Citrix*

$ApplicationName = "Production\Others\3270 TSB FirstData PRO"



#Usuarios Conectados
write-host "Usuarios Conectados"

$Users = (Get-BrokerSession -ApplicationInUse "$ApplicationName" |Select-Object UserName ).count
$Users 

#Sesiones Conectadas
write-host "Sesiones Conectadas"


#$Instances = Get-BrokerApplicationInstance -ApplicationName "*3270 TSB FirstData PRO*" | Select-Object Instances | ft -HideTableHeaders -AutoSize
Get-BrokerApplicationInstance -ApplicationName "$ApplicationName" | Select-Object Instances | ft -HideTableHeaders -AutoSize > \\serverhqr1\soft_ctx\Team\Ivan\Instances.txt

$Sesiones = Get-Content \\serverhqr1\soft_ctx\Team\Ivan\Instances.txt | ForEach-Object -Begin {[int]$Sum = 0} -Process {$Sum += $_} -End {$Sum}
$Sesiones




Function FEnviaCorreo{

    $vHora = get-date -Format H:mm
    $vMessage = "Hi, I attached the total sessions and users connected to the application APPv 'TSB FirstData 3270 PRO'
    -
    -
    -
    Hour,Users,Sessions
    -
    $vHora ,  $Users , $Sesiones
    -
    -
    -
    "
    $vDominio = "adgbs.com"
    $vSmtpServer = "relay.adgbs.com"
    #$vDe = "jose-john.carranza.lazaro@hpe.com"
    $vDe = "0901stservers@bancsabadell.com"
    $vPara = "miquel.micolau.mila@everis.com,incidentbs@hpe.com,incidentsabistsb@sabis.tech,VALERORAQ@sabis.tech,UROZJ@sabis.tech,MARTINEZJOSEAN@sabis.tech,LOVERIDGEA@sabis-tech.co.uk"
    #$vPara = "gomez-hernandez@hpe.com"
    $vCC = "gomez-hernandez@hpe.com"
    #$vCC = "israel.egea@hpe.com,sergio.aragon@hpe.com,carlos.rojano@hpe.com,eloi.andreu@hpe.com,guillermo-luis.oliver@hpe.com,miguel.alberich-plans@hpe.com"
    $vFecha = Get-Date -format "yyyy/MM/dd"
    $vAsunto = "Users & Sessions APPv 3270 TSB FirstData PRO " + $vHora

    $vMsg = new-object Net.Mail.MailMessage
    $vSmtp = new-object Net.Mail.SmtpClient($vSmtpServer)
    $vControl = 0

    $vMsg.From = $vDe
    $vMsg.CC.Add($vCC)
    #$vAtt1 = new-object Net.Mail.Attachment($vReportHTML)
    #$vMsg.Attachments.Add($vAtt1)
    $vControl = 1
	
    $vMsg.To.Add($vPara)
    $vMsg.Body = $vMessage
    $vMsg.Subject = $vAsunto

    if ($vControl = 1){
	
	   Write-Host "Enviando Correo a $vPara y $vCC"
	   $vSmtp.Send($vMsg) 
	   #$vAtt1.Dispose()
	}

}

#Production\Others\3270 TSB FirstData PRO





FEnviaCorreo