####################################################
#
#      Autor        Iván Gómez Hernández  
#      Date:        10/05/2017
#      Script       SearchKB_XA_v1.ps1
#      Descripción  Consulta Parches Seguridad KB
#      
#
#      Modificaciones:
# 
#      18/05/2017 - Filtrado para descartar maquinas apagadas
#      18/05/2017 - Añadir Estado Apagado y Falta KB
#      18/05/2017 - Formatear colores estados APAGADOS y FALTA KB
#      19/05/2017 - Modificado para consultas VDI
#      19/05/2017 - Cambio de la consulta WMI por si la version de powershell no conoce el alias wmi
#      22/05/2017 - Se cambia el tipo de consulta secuencial a jobs en pararelo para obtener el informe mucho más rapido en contra del diseño
#
#####################################################
#Declaración de constantes
##Contenido HTML
$HTML_Report = ""
$Fecha = get-Date
####################################################

####################################################
#Modulos necesarios

Add-PSSnapin citrix.*
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path



####################################################
#Generamos LOG

#echo "" > "$scriptDir\searchkbv1.log"

#$Log = "$scriptDir\searchkbv1.log"

###########################################################################

#Podemos pasarle un fichero o una consulta más generica de todos los serviores de la granja
#
#$Servidores = Get-XAServer "XAVBSV1005D"
#$Servidores = Get-BrokerDesktop | Where-Object {$_.DNSName -like "*vdnbpv0001d*" } | select-object DNSName
$Servidores = Get-XAServer 
#$Servidores = "localhost","XAC1","XAC","XAC2","XAC3","XAC4","XAC5","XAC6","XAC7","XAC8"

#Logs

get-date > "$scriptDir\SearchKB_Falta.html"
echo "<br>" >> "$scriptDir\SearchKB_Falta.html"
echo "SERVIDORES" >> "$scriptDir\SearchKB_Falta.html"
echo "<br>" >> "$scriptDir\SearchKB_Falta.html"

get-date > "$scriptDir\SearchKB_HotFix.html"
echo "<br>" >> "$scriptDir\SearchKB_HotFix.html"
echo "SERVIDORES" >> "$scriptDir\SearchKB_HotFix.html"
echo "<br>" >> "$scriptDir\SearchKB_HotFix.html"

get-date > "$scriptDir\Searchkb_Apagada.html"
echo "<br>" >> "$scriptDir\SearchKB_Apagada.html"
echo "SERVIDORES" >> "$scriptDir\SearchKB_Apagada.html"
echo "<br>" >> "$scriptDir\SearchKB_Apagada.html"

$check_hotfix =  {
param ($Servidor)


<# .SYNOPSIS Tests if a file is locked. #>
function Test-FileLocked
{
    param
    (
        [string]$FilePath
    )
    try { [IO.File]::OpenWrite($FilePath).close();return $false }
    catch {return $true}
}

       # filtramos si la máquina esta apagada para descartarla de la busqued 
        $MaquinaEncendida = Test-Connection -BufferSize 32 -Count 1 -ComputerName $Servidor -Quiet

        #Maquina Encendida
        if ($MaquinaEncendida -eq "True") 
        {
            $date = Get-date
            
            # Escogemos los KB que queramos consultar (KB4012212,KB4012213,KB4012214,KB4012598)
            $Hotfixes =  wmic /node:$Servidor qfe get hotfixid | Select-String 'KB4012212','KB4012213','KB4012214','KB4012598' 
            #write-host "Consulta finalizada en $Servidor - $date"
                
        #Maquina Sin HotFix
                if (!$Hotfixes)

                {
                #<B><FONT COLOR="red">Texto ROJO </FONT>
                $Serv= "" + '<B><FONT COLOR="red">' + "FALTA KB" + "</FONT>" + ""               
                Write-host "Realizando Informe para $Servidor"
                while ((Test-FileLocked  "\\capelo\soft_wintel\Citrix\Temporal\Ivan_Gomez\Scripts\Searchkb_Falta.html") -eq $true)
                {
                    Write-host  "fichero bloqueado esperando..."
                    Start-Sleep -Seconds 3
                    
                      
                }
                
                TRY {
                echo "<br>" >> "\\capelo\soft_wintel\Citrix\Temporal\Ivan_Gomez\Scripts\SearchKB_Falta.html"
                echo "$Servidor","$Serv" >> "\\capelo\soft_wintel\Citrix\Temporal\Ivan_Gomez\Scripts\SearchKB_Falta.html"
                }
                CATCH { write-host ""}

                }

                else {
                    
                    Foreach($hotfix in $Hotfixes)
                    {
                    
                    $Serv += ""+ '<B><FONT COLOR="green">'+ $hotfix +"</FONT>"
                    Write-host "Realizando Informe para $Servidor"
                        
                    }
                    
                    while ((Test-FileLocked  "\\capelo\soft_wintel\Citrix\Temporal\Ivan_Gomez\Scripts\Searchkb_HotFix.html") -eq $true)
                    {
                    Write-host  "fichero bloqueado esperando..."
                    Start-Sleep -Seconds 3
                    
                      
                    }
                    TRY {
                    
                    echo "<br>" >> "\\capelo\soft_wintel\Citrix\Temporal\Ivan_Gomez\Scripts\SearchKB_HotFix.html"
                    echo "$Servidor","$Serv" >> "\\capelo\soft_wintel\Citrix\Temporal\Ivan_Gomez\Scripts\SearchKB_HotFix.html"
                    }
                    CATCH {write-host ""}

                }
        }

        #Maquina Apagada
        else {
        $Serv =""+ "<b>" + "APAGADA" + "</b>" + ""
        while ((Test-FileLocked  "\\capelo\soft_wintel\Citrix\Temporal\Ivan_Gomez\Scripts\Searchkb_Apagada.html") -eq $true)
            {
                    Write-host  "fichero bloqueado esperando..."
                    Start-Sleep -Seconds 3
                    
                      
                }
                         
        Write-host "Realizando Informe para $Servidor"
        #Write-Host "$Servidor",$Serv
        TRY {
        echo "<br>" >> "\\capelo\soft_wintel\Citrix\Temporal\Ivan_Gomez\Scripts\Searchkb_Apagada.html"
        echo "$Servidor","$Serv" >> "\\capelo\soft_wintel\Citrix\Temporal\Ivan_Gomez\Scripts\Searchkb_Apagada.html"
        }

        CATCH {Write-host ""}

        }
           



}




# Array con los campos y valores para el Report


Foreach ($Servidor in $Servidores)

{

write-host "Lanzando tarea en:" $Servidor
Start-Job -ScriptBlock $check_hotfix -ArgumentList $Servidor | Out-Null

 
}
Write-Host ""
Write-Host ""
Write-Host "Esperando Informe..."
Write-Host ""
Write-Host ""


#Get-Job | Receive-Job 

#WAIT-JOB

#############################
# Get all the running jobs
$jobs = get-job | ? { $_.state -eq "running" }
$total = $jobs.count
$runningjobs = $jobs.count

# Loop while there are running jobs
while($runningjobs -gt 0) {
# Update progress based on how many jobs are done yet.
$percent=[math]::Round((($total-$runningjobs)/$total * 100),2)
write-progress -activity "Starting Provisioning Modules Instances" -status "Progress: $percent%" -percentcomplete (($total-$runningjobs)/$total*100)

# After updating the progress bar, get current job count
$runningjobs = (get-job | ? { $_.state -eq "running" }).Count
}

#############################


$Falta = get-content "$scriptDir\SearchKB_Falta.html"
$HotFix = get-content "$scriptDir\SearchKB_HotFix.html"
$Apagada = get-content "$scriptDir\Searchkb_Apagada.html"


$count_Falta = ($Falta | select-string -pattern "FALTA KB").length
$count_HotFix =  ($HotFix | select-string -pattern "KB*").length
$count_Apagada = ($Apagada | select-string -pattern "APAGADA").length


#Contadores Falta
echo "<br>" >> "$scriptDir\SearchKB_Falta.html"
echo "<br>" >> "$scriptDir\SearchKB_Falta.html"
echo "Existen" $count_Falta "máquinas sin actualizar" >> "$scriptDir\SearchKB_Falta.html"

#Contadores Hotfix
echo "<br>" >> "$scriptDir\SearchKB_HotFix.html"
echo "<br>" >> "$scriptDir\SearchKB_HotFix.html"
echo "Existen" $count_HotFix "máquinas Actualizadas" >> "$scriptDir\SearchKB_HotFix.html"

#Contadores Apagada
echo "<br>" >> "$scriptDir\Searchkb_Apagada.html"
echo "<br>" >> "$scriptDir\Searchkb_Apagada.html"
echo "Existen" $count_Apagada "máquinas Apagadas" >> "$scriptDir\Searchkb_Apagada.html"