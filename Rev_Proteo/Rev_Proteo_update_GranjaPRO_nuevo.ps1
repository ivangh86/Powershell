# Script que revisa el update de Proteo en los servidores de Produccion TSB UK  
# Si se especifica el parametro -sendmail el script envia mail con el resultado
# Si se especifica el parametro -correction reinicia los servers que tienen uptime incorrecto (sin usuarios conectados) y relanza la tarea SibisUpdate en los que tienen el Uptime correcto y errores de actualizacion (sin usuarios conectados)
# Si se especifica el parametro -serverlist solo revisa los servidores en el archivo $Archivo_serverlist

#v2.0
#	*Que el log de resultado se guarde en carpeta Reports 
#	*Que se genere un log OK 
#	*Pause al final si no hay ningún parámetro 
#	*Optimización test-connection (implementada por Enric) 
#	*Permisos en carpetas log y binarios 
#	*Parámetro para revisar solo una lista de XenApps 
#	*Exclusión list
#	*Consulta si se ejecuta en un DDC y si lo hace no Invoke-command



[cmdletbinding()] 

Param
(
[switch]$sendmail = $false,
[switch]$correction = $false,
[switch]$serverlist = $false
)


#Variables
[String]$today = get-date -format dd/MM/yyyy
$end =get-date -format ddMM_hhmmss

$currentPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path

$Archivo_serverlist = $currentPath + "\Server_list.txt"
#write-host $Archivo_serverlist
$Archivo_exclusionlist = $currentPath + "\Exclusion_list.txt"
$archivo_errores = $currentPath + "\Report\Rev_Proteo_update_GranjaPRO_NOK.log"
$archivo_errores0 = $currentPath + "\Report\0_Rev_Proteo_update_GranjaPRO_serversNOK.log"
$archivo_errores_sin_usuarios = $currentPath + "\Report\Rev_Proteo_srv_PRO_tmp.log"
$archivo_errores_sin_usuarios0 = $currentPath + "\Report\0_Rev_Proteo_srv_PRO_tmp.log"
$archivo_errores_reboot = $currentPath + "\Report\Rev_Proteo_srv_PRO_uptime_tmp.log"
$archivo_errores_reboot_sin_usuarios = $currentPath + "\Report\Rev_Proteo_srv_PRO_uptime_tmp_su.log"
$archivo_historico = $currentPath + "\Report\historico\Rev_Proteo_update_GranjaPRO_NOK_" + $end + ".log"

$vXenDDC ="WCPNHIDDC001"
 
$Archivo_0_TOT = "d$\temp\sibis_ab_Update.txt"                                        #Archivo update sibis

$Archivo_1_PRO = "D$\usr\local\eb\JCDEPROD\lotes_pilotados\PRODECLIPSE_defectos.lst"  #Archivo update Eclipse
$Archivo_2_PRO = "D$\usr\local\eb\JCDEPROD\lotes_pilotados\PRODEBRANCH_defectos.lst"  #Archivo update Ebranch
$Archivo_3_PRO = "d$\USR\LOCAL\eb\JCDEPROD\lots\histo\eBranchHisto"                   #Archivo update Ebranch

$Archivo_1_FOR = "D$\usr\local\eb\JCDEFORM\lotes_pilotados\FORMECLIPSE_defectos.lst"  #Archivo update Eclipse
$Archivo_2_FOR = "D$\usr\local\eb\JCDEFORM\lotes_pilotados\FORMEBRANCH_defectos.lst"  #Archivo update Ebranch
$Archivo_3_FOR = "d$\USR\LOCAL\eb\JCDEFORM\lots\histo\eBranchHisto"                   #Archivo update Ebranch


Function Run_Command_with_TimeOut {
Param(
	[Parameter(Mandatory=$true)] 
	$Code,
	[Parameter(Mandatory=$false)] 
	[int]$TimeOut = 2,
	[Parameter(Mandatory=$false)] 
	[int]$refresh_interval = 200	
)


$Runspace = [runspacefactory]::CreateRunspace()
$PowerShell = [System.Management.Automation.PowerShell]::Create()
$PowerShell.runspace = $Runspace
$Runspace.Open()

#Add as many commands as we want to be executed
[void]$PowerShell.AddScript($Code)

#Launch the commands
$timer = [Diagnostics.Stopwatch]::StartNew()
$AsyncObject = $PowerShell.BeginInvoke()

#wait for the Launch commands to complete and kill if timeout reached
$timetaken = $timer.Elapsed.TotalSeconds
while (-Not $AsyncObject.Iscompleted ) {
	$timetaken = $timer.Elapsed.TotalSeconds
	#write-host "runspace job commands not completed, please wait..." -foregroundcolor yellow	
	start-sleep -Milliseconds $refresh_interval
	#Kill when time-out reached!!
	if ($timetaken -ge $timeout) {
		$timer.Stop()
		[switch]$Timeout_reached = $true
		write-host "runspace job time-out reached!, time-out is $timeout, KO" -foregroundcolor red
		break
	}
}
#OK msg if no time-out
#if (!$Timeout_reached) {
#	write-host "runspace job commands finished in $timetaken, no time-out reached, OK" -foregroundcolor green
#}

#Check if commands executed had errors and set result variable if so
if (-Not $PowerShell.HadErrors) {
	$Result = $PowerShell.EndInvoke($AsyncObject)
} else {
	$Result = $PowerShell.Streams.Error
}

#Return Result
$Result

#Clean-Up Runspace
$PowerShell.Dispose()

}


#Funcion que envia el correo informativo
Function FEnviaCorreo($vReportHTMLOK){

    $vHora = get-date -Format H:mm 
    $vMessage = "Buenos días, Adjunto servidores de TSB con error en el update de Proteo. `r"
    $vDominio = "adgbs.com"
    $vSmtpServer = "relay.adgbs.com"   
    $vDe = "0901stservers@bancsabadell.com"
    $vPara = "bs_wintel_ctx@dxc.com"
    $vCC = "alejandro.blanco.alonso@everis.com, PETCHAME.JORDI@sabis.tech"
    $vFecha = Get-Date -format "yyyy/MM/dd"
    $vAsunto = "TSB Update Proteo -" + $vHora

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



#Funcion que realiza la revision de archivos
function Revision_archivo([string]$servidor, [string]$archivo, [string]$sessions ) 
{
	$path = "\\" + $servidor + "\" + $archivo
    
    if (test-path $path)
    {

        [String]$UPD_time=(get-item $path).lastwritetime.ToString("dd/MM/yyyy  hh:mm:ss")
            
        if ($UPD_time -match $today) 
        { 
            write-host $servidor $archivo $UPD_time -foregroundcolor "green" 
            "$servidor $archivo $UPD_time " | out-file -append $archivo_historico
        }
        else 
        {
             write-host $servidor $archivo $UPD_time -foregroundcolor "red" 
             "$servidor $archivo $UPD_time " | out-file -append $archivo_historico
             $servidor | out-file -append $archivo_errores0
             if ($sessions -like 0) { $servidor | out-file -append $archivo_errores_sin_usuarios0 }
        }
         
    }
    else
    {
        write-host $servidor $archivo "NO EXISTE" -foregroundcolor "red" 
        "$servidor $archivo NO EXISTE" | out-file -append $archivo_historico
        $servidor | out-file -append $archivo_errores0
        if ($sessions -like 0) { $servidor | out-file -append $archivo_errores_sin_usuarios0 }
    }
	
}


#Funcion que realiza la revision de un Server
function Revision_Server([string]$server, [string]$sesconc, [String]$Archivo_0, [String]$Archivo_1, [String]$Archivo_2, [String]$Archivo_3, [boolean]$Maintenance, [String]$Registred, [String]$Power ) 
{
    
    if(test-connection $server -count 1 -ErrorAction SilentlyContinue)
    {
        #Calcula el Uptime
        $date = Run_Command_with_TimeOut -Code "Get-WmiObject -ComputerName $server -Class Win32_OperatingSystem -ErrorAction Silentlycontinue" -TimeOut 5 -refresh_interval 500
        #$date = Get-WmiObject -ComputerName $server -Class Win32_OperatingSystem -ErrorAction Silentlycontinue       
        $Lastboot = ($date.ConvertToDateTime($date.LastBootUpTime)).toshortDatestring()
        $hoy = get-date -format d
        
        <#If ($Maintenance)
         {
            write-host $server "Servidor en mantenimiento" -ForegroundColor "yellow"
            "$server Servidor en mantenimiento"| out-file -append $archivo_historico
        }
        else
        {
#>
            if ($Lastboot -ne $hoy)
            {
                write-host $server "Servidor con Uptime incorrecto" -ForegroundColor "red"
                "$server Servidor con Uptime incorrecto"| out-file -append $archivo_historico
                $server | out-file -append $archivo_errores_reboot
                if ($sesconc -like 0) { $server | out-file -append $archivo_errores_reboot_sin_usuarios }           
            }
            else
            {       
                Revision_archivo $server $Archivo_0 $sesconc
                Revision_archivo $server $Archivo_1 $sesconc
                Revision_archivo $server $Archivo_2 $sesconc
            
                $Archivo_3c = $Archivo_3 +"_" + $hostname + ".xml"
                Revision_archivo $server $Archivo_3c $sesconc
            } 
   #     }  
           
        write-host $server "Sesiones en el servidor:" $sesconc 
        "$server Sesiones en el servidor: $sesconc" | out-file -append $archivo_historico      
        write-host $server "Uptime:"$Lastboot 
        "$server Uptime: $Lastboot" | out-file -append $archivo_historico
        write-host $server "Maintenance:" $Maintenance 
        "$server Maintenance: $Maintenance" | out-file -append $archivo_historico
        
          
            
    }
    else
    {
        if ($Power -eq "On") 
        { 
            
            if ($Registred -eq "Unregistered") 
            {
                write-host $server "servidor Unregistered" -foregroundcolor "red" 
                "$server servidor Unregistered" | out-file -append $archivo_historico
                $server | out-file -append $archivo_errores_reboot
                if ($sesconc -like 0) { $server | out-file -append $archivo_errores_reboot_sin_usuarios }  
            } 
            Else
            {
                write-host $server "Problema de conectividad" -foregroundcolor "yellow" 
                "$server Problema de conectividad" | out-file -append $archivo_historico
            } 
        }
        else 
        { 
            write-host $server "Servidor Apagado" -foregroundcolor "yellow" 
            "$server Servidor Apagado" | out-file -append $archivo_historico 
        }
    }
    write-host 
    " " | out-file -append $archivo_historico
}


#MAIN
if(test-path $Archivo_exclusionlist) { $Exclusion = get-content $Archivo_exclusionlist }
if(test-path $archivo_errores ) { remove-item $archivo_errores }

# Si hay se ha pasado el parametro -serverlist solo se revisan los servers en el archivo $Archivo_serverlist
#if ($serverlist) 
#{
#    $Servers_a_revisar = get-content $Archivo_serverlist
#    $Result = @()
#    foreach ($server in $Servers_a_revisar)
  #  {
       #if ($env:computername -match "DDC")
        #{
 #           if ((GET-PSSnapin -name Citrix* -ErrorAction SilentlyContinue) -eq $Null) {Add-PSSnapin Citrix*}
 #           $Result += Get-BrokerMachine -machinename "TSB\$server" | Select DNSName,CatalogName,DesktopGroupName,SessionCount,InMaintenanceMode,PowerState,RegistrationState
        #}
        #else
        #{
        #    write-host "El parametro -serverlist solo funciona correctamente cuando el script se lanza desde un DDC"
            #$Result = Invoke-Command -ComputerName $vXenDDC -ScriptBlock {if ((GET-PSSnapin -name Citrix* -ErrorAction SilentlyContinue) -eq $Null) {Add-PSSnapin Citrix*} ; Get-BrokerMachine -machinename "TSB\$server" } | Select DNSName,CatalogName,DesktopGroupName,SessionCount,InMaintenanceMode,PowerState,RegistrationState  
        #}
       
  #  }
  #  $vBrokerDesktops = $Result 
#}
#else
#{ 
    #if ($env:computername -match "DDC")
    #{
        if ((GET-PSSnapin -name Citrix* -ErrorAction SilentlyContinue) -eq $Null) {Add-PSSnapin Citrix*}
        $Result = Get-BrokerMachine -MaxRecordCount 10000 | Select DNSName,CatalogName,DesktopGroupName,SessionCount,InMaintenanceMode,PowerState,RegistrationState
        
    #}
    #else
    #{
    #    $Result = Invoke-Command -ComputerName $vXenDDC -ScriptBlock {if ((GET-PSSnapin -name Citrix* -ErrorAction SilentlyContinue) -eq $Null) {Add-PSSnapin Citrix*} ; Get-BrokerMachine -MaxRecordCount 10000 } | Select DNSName,CatalogName,DesktopGroupName,SessionCount,InMaintenanceMode,PowerState,RegistrationState  
    #}
    $vBrokerDesktops = $Result | Sort-Object DNSName -unique
    
#}

foreach ($vBrokerDesktop in $vBrokerDesktops)
{
    #$sesiones = $vBrokerDesktop.SessionCount
    $hostname = (($vBrokerDesktop.DNSName).SubString(0,12))
    write-host $hostname
    
    #if (-not ($hostname -match $Exclusion))
    #{
        # Revisa los archivos en los XenApps del DeliveryGroup de PRODUCCION
        if ($vBrokerDesktop.DesktopGroupName -eq "SDI W2K12 PRO") { Revision_Server $hostname $vBrokerDesktop.SessionCount $Archivo_0_TOT $Archivo_1_PRO $Archivo_2_PRO $Archivo_3_PRO $vBrokerDesktop.InMaintenanceMode $vBrokerDesktop.RegistrationState $vBrokerDesktop.PowerState }

        # Revisa los archivos en los XenApps del DeliveryGroup de UAT
        if ($vBrokerDesktop.DesktopGroupName -eq "SDI W2K12 UAT") { Revision_Server $hostname $vBrokerDesktop.SessionCount $Archivo_0_TOT $Archivo_1_FOR $Archivo_2_FOR $Archivo_3_FOR $vBrokerDesktop.InMaintenanceMode $vBrokerDesktop.RegistrationState $vBrokerDesktop.PowerState }
   
        # Revisa los archivos en los XenApps de TEST
        #if ($vBrokerDesktop.CatalogName -eq "SDI W2K12 TEST APPS") { Revision_Server $hostname $vBrokerDesktop.SessionCount $Archivo_0_TOT $Archivo_1_FOR $Archivo_2_FOR $Archivo_3_FOR $vBrokerDesktop.InMaintenanceMode $vBrokerDesktop.RegistrationState $vBrokerDesktop.PowerState }
    #}
}


#Deja solo registros unicos en los archivos de errores
if(test-path $archivo_errores0 )
{
    gc $archivo_errores0 | sort | get-unique > $archivo_errores
    remove-item $archivo_errores0
}

if(test-path $archivo_errores_sin_usuarios0)
{
    gc $archivo_errores_sin_usuarios0 | sort | get-unique >  $archivo_errores_sin_usuarios
    remove-item $archivo_errores_sin_usuarios0
}

#Write-host "CORRECTION" -foregroundColor "yellow"
#"CORRECTION" | out-file -append $archivo_historico  

# Lanza la tarea programada SibisUpdate de los servidores con errores en el update y sin usuarios

if (test-path $archivo_errores_reboot)
{           
        write-host
        write-host "Servidores con uptime incorrecto" -foregroundColor "yellow"
        gc $archivo_errores_reboot
        $srv_uptime = (gc $archivo_errores_reboot).count
        
        if (test-path $archivo_errores_reboot_sin_usuarios)
        { 
            write-host
            write-host "Servidores con uptime incorrecto y sin usuarios conectados:" -foregroundColor "yellow"
            gc $archivo_errores_reboot_sin_usuarios
           
            if ($correction)  
            {
                write-host
                Write-host "Lanzando rebbot en los servidores con uptime incorrecto y sin usuarios..." -foregroundColor "yellow"
                "Lanzando rebbot en los servidores con uptime incorrecto y sin usuarios..." | out-file -append $archivo_historico  
                $servers_KO_uptime = Get-Content $archivo_errores_reboot_sin_usuarios

                foreach ($server_KO_uptime in $servers_KO_uptime)
                {                     
                    restart-computer $server_KO_uptime
                    write-host $server_KO_uptime " reiniciada"
                    "$server_KO_uptime  reiniciada" | out-file -append $archivo_historico
                }
            }
        }
}
        
if (test-path $archivo_errores)
{
    write-host
    write-host "Servidores con error en la actualizacion de Proteo:" -foregroundColor "yellow"
    gc $archivo_errores
    $srv_actual = (gc $archivo_errores).count
    
    if (test-path $archivo_errores_sin_usuarios)
    {       
        write-host
        write-host "Servidores con error en la actualizacion de Proteo y sin usurios conectados:" -foregroundColor "yellow"
        gc $archivo_errores_sin_usuarios

        if ($correction) 
        {
            write-host
            Write-host "Lanzando la tarea SibisUpdate en los servidores con error en la actualizacion y sin usuarios..." -foregroundColor "yellow"
            "Lanzando la tarea SibisUpdate en los servidores con error en la actualizacion y sin usuarios..." | out-file -append $archivo_historico 
            $servers_KO = Get-Content $archivo_errores_sin_usuarios 

            foreach ($server_KO in $servers_KO)
            {       
                    schtasks /run /s $server_KO /TN "SibisUpdate" 
                    write-host $server_KO " Tarea programada SibisUpdate lanzada"
                    "$server_KO  Tarea programada SibisUpdate lanzada" | out-file -append $archivo_historico 
            }
        }
    }

}
else
{
    write-host
    write-host "No hay servidores con error en el update de Proteo" -ForegroundColor "yellow"
    "No hay servidores con error en el update de Proteo" | out-file -append $archivo_historico
    "No hay servidores con error en el update de Proteo" | out-file -append $archivo_errores
}


write-host "Servidores con error de uptime $srv_uptime "
write-host "Servidores con error de actualizacion $srv_actual "

#Se envia el correo con el resultado
if ($sendmail) { FEnviaCorreo $archivo_errores }

# Borrado de archivos de control
if (test-path $archivo_errores_sin_usuarios) { remove-item $archivo_errores_sin_usuarios }
if (test-path $archivo_errores_reboot) { remove-item $archivo_errores_reboot }
if (test-path $archivo_errores_reboot_sin_usuarios) { remove-item $archivo_errores_reboot_sin_usuarios }


# Si no hay parametros, hacer una pausa al finalizar la ejecucion
if (( -not $sendmail) -or ( -not $correction)) 
{
    Write-Host "Press any key to continue ..."
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}