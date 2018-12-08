Param
(
[switch]$sendmail = $false
)

#Funcion que envia el correo informativo
Function FEnviaCorreo($vReportHTMLOK){

    $vHora = get-date -Format H:mm 
    $vMessage = "Buenos días, archivo con Xenapps. `r"
    $vDominio = "adgbs.com"
    $vSmtpServer = "relay.adgbs.com"   
    $vDe = "0901stservers@bancsabadell.com"
    
    $vPara = "serrajo@sabis.tech,incidentmanagementbs@dxc.com,bs_governance_desktop@dxc.com"
    
    $vCC = "bs_wintel_ctx@dxc.com"
    $vFecha = Get-Date -format "yyyy/MM/dd"
    $vAsunto = "Xenapps TSB Desarrolladores con unidades llenas -" + $vHora

    $vMsg = new-object Net.Mail.MailMessage
    $vSmtp = new-object Net.Mail.SmtpClient($vSmtpServer)
    $vControl = 0

    $vMsg.From = $vDe
    $vMsg.CC.Add($vCC)
    $vAtt1 = new-object Net.Mail.Attachment($vReportHTMLOK)
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


######     
#MAIN

# Cargar modulo Citrix
write-host "Cargando modulo de Citrix..."
if ((GET-PSSnapin -name Citrix* -ErrorAction SilentlyContinue) -eq $Null) {Add-PSSnapin Citrix*}

$end =get-date -format ddMM_hhmmss

$currentPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
$PathAttach = $currentPath + "\Reports\log_TSB_DES_Llenos_" + $end + ".log"

$list = Get-XAServer -ServerName "XATISV0*"
$list += Get-XAServer -ServerName "XATBSV0*"

Write-Host  "SERVIDOR UNIDAD FREESPACE" -ForegroundColor Yellow
$llenos = @()
"Servidor;Unidad;GB_libres"| Out-file $PathAttach

Foreach ($server_full in $list)
{
     $server = $server_full.servername
     
     if (test-connection $server -count 1 -EA 0)
     {
         $dp = Get-WmiObject win32_logicaldisk -ComputerName $server |  Where-Object {$_.drivetype -eq 3}
         foreach ($item in $dp)
         {
            $totSpace = $item.Size
            $frSpace = $item.FreeSpace
            $totSpace=[math]::Round(($totSpace/1073741824),2)
            $frSpace=[Math]::Round(($frSpace/1073741824),2)
            $freePercent = ($frspace/$totSpace)*100
            $freePercent = [Math]::Round($freePercent,0)
       
            if ($freePercent -lt 3 )
            {
                Write-Host  "$server $($item.DeviceID) $frSpace GB libres" -ForegroundColor red
                $llenos += $server
                "$server;$($item.DeviceID);$frSpace" | Out-file -append $PathAttach 
            }
            else
            {
                Write-Host  "$server $($item.DeviceID) $frSpace GB libres" -ForegroundColor green
                #Write-Host  "$server $($item.DeviceID) $($item.Size)  $frSpace $freePercent%_libre" -ForegroundColor green
            }
        
         }
     }
}

write-host "Servidores con unidades llenas:" -ForegroundColor yellow
write-host $llenos


#Se envia el correo con el resultado de la actualizacion de Proteo
if ($sendmail) { if ( test-path $PathAttach ) { FEnviaCorreo $PathAttach } }


