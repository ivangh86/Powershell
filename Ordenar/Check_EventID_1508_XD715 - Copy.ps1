########################################################################################################################
# Script que busca el evento 1048 "Citrix Desktop Service" en los Xenapps del delivery group de CORE PRO               #
# Permite enviar correo del report pasando el parámetro -MAIL  (Recuerda Revisar destinatarios del correo              #
# Ricardo Zamora 29/03/2018                                                                                            #
# Ivan Gomez 15/10/2018
########################################################################################################################


#### Email ################################################################################

Param(
	[Parameter(Mandatory=$false)] 
	[switch]$MAIL
)


Function FEnviaCorreo($PathAttach){

    $vHora = get-date -Format H:mm
    
    $vMessage = $vMessageResult
    $vDominio = "adgbs.com"
    $vSmtpServer = "relay.adgbs.com"
    
    
    $vDe = "0901stservers@bancsabadell.com" 
    #$vPara = "Enrique.Cordero@citrix.com"
    $vPara = "bs_wintel_ctx@dxc.com"
    $vCC = "cuerdar@bancsabadell.com,delcastilloj@sabis.tech"
    $vFecha = Get-Date -format "yyyy/MM/dd"
    $vAsunto = "BS - Report Evento $ID $ProviderName - " + $vHora

    $vMsg = new-object Net.Mail.MailMessage
    $vSmtp = new-object Net.Mail.SmtpClient($vSmtpServer)
    $vControl = 0

    $vMsg.From = $vDe
    $vMsg.CC.Add($vCC)
    $vAtt1 = new-object Net.Mail.Attachment($PathAttach)
    $vMsg.Attachments.Add($vAtt1)
    $vControl = 1
	
    $vMsg.To.Add($vPara)
    $vMsg.Body = $vMessage
    $vMsg.Subject = $vAsunto

    if ($vControl = 1){
	
	   Write-Host "Enviando Correo a $vPara y $vCC"
	   $vSmtp.Send($vMsg) 
	   $vAtt1.Dispose()
	}

}


asnp Citrix*
if ((Get-PSSnapin "Citrix.Broker.Admin.*" -EA silentlycontinue) -eq $null) {
try { Add-PSSnapin Citrix* -ErrorAction Stop }
catch { write-error "Error loading Citrix PowerShell snapin"; Return }
}


[String]$today = get-date -format dd/MM/yyyy
$end =get-date -format ddMM_hhmmss
$ID1 = "1015"
#$ID2 = "1502"
$DG = "SDI W2K8R2 OFI PRO"
$Source = "Application"
$ProviderName= "Citrix Desktop Service"
#$ProviderName= "TerminalServices-RemoteConnectionManager"

$currentPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
$PathAttach = $currentPath + "\Report\Error_EventsID_" + $ID1 + "_" + "Date_" + $end + ".csv"
"Machine ;DeliveryGroup ; Event ID ; Message" | Out-File $PathAttach

$Serverlist = Get-BrokerMachine -DesktopGroupName $DG  -MaxRecordCount 10000 | Select DNSName,CatalogName,DesktopGroupName,SessionCount,InMaintenanceMode,PowerState,RegistrationState,Tags


$days=((Get-Date).AddDays(-1)) 
$total = @()
$countID = 0

$vCount = $Serverlist.count
$vi = 0
Foreach ($Server in $Serverlist) 
{
    $vi ++
    Write-host "$vi / $vCount"
    $Machine = $Server.DNSName
    
    $events = Get-WinEvent -Computername $Machine -FilterHashtable @{Logname=$Source ;ID=@($ID1); providername= $ProviderName ; StartTime=$days} -EA 0
    
    if($events){
        ForEach($event in $events){

            write-host "$machine  ; $($Server.DesktopGroupName) ; $($event.ProviderName) ; $($event.ID) ; $($event.TimeCreated); $($event.Message)" | ft -autosize

            "$machine  ; $($Server.DesktopGroupName) ; $($event.ID) ; $($event.TimeCreated)"  | Out-file -append $PathAttach

            $countID++
        }
       
            
    
    }
}
Write-Host " "

Write-host "El evento $ID '$ProviderName' se ha producido $countID veces, en el Delivery Group $DG durante las últimas 24 horas." -ForegroundColor Green
$vMessageResult= "El evento $ID '$ProviderName' se ha producido $countID veces en las últimas 24h."


If($MAIL -eq $true){
    FEnviaCorreo $PathAttach
}



