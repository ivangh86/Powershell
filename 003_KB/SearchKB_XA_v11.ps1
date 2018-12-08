####################################################
#
#      Autor        Iván Gómez Hernández  
#      Date:        10/05/2017
#      Script       SearchKB_XA_v1.ps1
#      Descripción  Consulta Parches Seguridad KB
#      
#
#      Modificaciones: 
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

####################################################
#Funciones para gestionar la salida HTML




###########################################################################

#Podemos pasarle un fichero o una consulta más generica de todos los serviores de la granja
#
#$Servidores = Get-XAServer XATISV0001R
#$Servidores = Get-BrokerDesktop | Where-Object {$_.DNSName -like "*vdnbpv0001d*" } | select-object DNSName
#$Servidores = Get-XAServer | Sort-Object -Property ServerName
$Servidores = "localhost","XAC"

#Logs

get-date > "$scriptDir\SearchKB_Falta.html"
get-date > "$scriptDir\SearchKB_HotFix.html"
get-date > "$scriptDir\Searchkb_Apagada.html"

$check_hotfix =  {
param ($Servidor)

       # filtramos si la máquina esta apagada para descartarla de la busqued 
        $MaquinaEncendida = Test-Connection -BufferSize 32 -Count 1 -ComputerName $Servidor -Quiet

        #Maquina Encendida
        if ($MaquinaEncendida -eq "True") 
        {
            $date = Get-date
            write-host "Buscando parches en $Servidor - $date"
            # Escogemos los KB que queramos consultar (KB4012212,KB4012213,KB4012214,KB4012598)
            $Hotfixes =  get-wmiobject -class win32_quickfixengineering -computer $Servidor | select-object HotFixID | Select-String 'KB4012212','KB4012213','KB4012214','KB4012598' 
            
                
        #Maquina Sin HotFix
                if (!$Hotfixes)

                {
                #<B><FONT COLOR="red">Texto ROJO </FONT>
                $Serv= "" + '<B><FONT COLOR="red">' + "FALTA KB" + "</FONT>" + ""               
                Write-Host "$Servidor","$Serv"
                echo "$Servidor","$Serv" |Out-File "$scriptDir\SearchKB_Falta.html"
                }

                else {
                    
                    Foreach($hotfix in $Hotfixes)
                    {
                    $Serv += ""+ $hotfix +""
                    Write-Host "$Servidor","$Serv"
                    echo "$Servidor","$Serv" | Out-File "$scriptDir\SearchKB_HotFix.html"
                    }
                }
        }

        #Maquina Apagada
        else {
        $Serv =""+ "<b>" + "APAGADA" + "</b>" + ""
      
        Write-Host "$Servidor",$Serv
        echo "$Servidor","$Serv" | Out-File "$scriptDir\Searchkb_Apagada.html"
        }   



}




# Array con los campos y valores para el Report


Foreach ($Servidor in $Servidores)

{

write-host "Lanzando tarea en:" $Servidor
Start-Job -ScriptBlock $check_hotfix -ArgumentList $Servidor | Out-Null

 
}




Get-Job | Wait-Job| Receive-Job 


$Falta = get-content "$scriptDir\SearchKB_Falta.html"
$HotFix = get-content "$scriptDir\SearchKB_HotFix.html"
$Apagada = get-content "$scriptDir\Searchkb_Apagada.html"


$count_Falta = ($Falta | select-string -pattern "FALTA KB").length
$count_HotFix =  ($HotFix | select-string -pattern "KB*").length
$count_Apagada = ($Apagada | select-string -pattern "APAGADA").length


#Contadores Falta
echo "<br>" >> "$scriptDir\SearchKB_Falta.html"
echo "<br>" >> "$scriptDir\SearchKB_Falta.html"
echo "Existen" "$count_Falta" "máquinas sin actualizar" >> "$scriptDir\SearchKB_Falta.html"

#Contadores Hotfix
echo "<br>" >> "$scriptDir\SearchKB_HotFix.html"
echo "<br>" >> "$scriptDir\SearchKB_HotFix.html"
echo "Existen" "$count_HotFix" "máquinas Actualizadas" >> "$scriptDir\SearchKB_HotFix.html"

#Contadores Apagada
echo "<br>" >> "$scriptDir\Searchkb_Apagada.html"
echo "<br>" >> "$scriptDir\Searchkb_Apagada.html"
echo "Existen" "$count_Apagada" "máquinas Apagadas" >> "$scriptDir\Searchkb_Apagada.html"