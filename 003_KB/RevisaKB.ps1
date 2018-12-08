############################################################
#
#      Autor:        Iván Gómez Hernández  
#      Date:         22/06/2017
#      Script:       RevisaKB.ps1
#      Descripción:  Consulta Parches Seguridad KB
#      
#      Resultados:   "$scriptDir\ReportKB.html" "$scriptDir\ReportKB.csv" 
#
#############################################################

#Ruta Dinámica
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

#Encabezado para nuestro log csv
echo "Server;Estado;Parche" > "$scriptDir\ReportKB.csv"
$Log = "$scriptDir\ReportKB.csv"

#Posibles Fuentes de Datos
$Servers = Get-Content "$scriptDir\ServerList.txt"
#$servers = "localhost"
#$servers = Get-XAServer
#$Servers = Get-BrokerDesktop | Where-Object {$_.DNSName -like "*vdnbpv0001d*" } |Select-Object DNSName,AgentVersion,OSType,DesktopGroupname,DesktopKind,RegistrationState

#Función que revisa los KB
$check_hotfix = {
    param ($Server)
    $MaquinaEncendida = Test-Connection -BufferSize 32 -Count 1 -ComputerName $Server -Quiet
    $Parches = 'KB4012212','KB4012213','KB4012214','KB4012598'

   if ($MaquinaEncendida -eq "True") {
    foreach ($Parche in $Parches) {
    $is_patched = (get-wmiobject -class win32_quickfixengineering -computer $Server | select-object HotFixID | Select-String $Parche) -ne $null 
    if ($is_patched) { 
        Write-Output ($Server+ ";Exist;" +$Parche) 
    } 
    else { 
        Write-Output ($Server+ ";Falta;"+ $Parche)
    }
    }
    }
    else {
    Write-Output ($Server+ ";NoConecta")
    }
    }

foreach ($Server in $Servers) {
    Start-Job -ScriptBlock $check_hotfix -ArgumentList $Server | Out-Null
    write-host "Lanzando tarea en:" $Server
}

write-host ""
write-host ""
write-host ""
write-host "esperando resultados"
Get-Job | Wait-Job | Receive-Job | out-file -filepath $Log -Append

Write-host "PROCESO FINALIZADO, REVISAR" $Log


#Pasarlo a HTML
Get-Content $Log | ConvertFrom-Csv -Delimiter ';' | ConvertTo-Html | Out-String | out-file -filepath "$scriptDir\ReportKB.html"
Invoke-Expression  "$scriptDir\ReportKB.html"