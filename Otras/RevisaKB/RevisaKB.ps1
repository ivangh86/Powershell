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
    $Parches = 'KB3203392','KB3203461','KB3203392','KB3203461','KB3114400', 'KB3118310','KB3203461','KB3114340','KB3172458','KB3203461','KB3114340','KB3172458','KB3203392'
   if ($MaquinaEncendida -eq "True") {
    foreach ($Parche in $Parches) {
    $is_patched = (get-wmiobject -class win32_quickfixengineering -computer $Server -ErrorAction STOP | select-object HotFixID | Select-String $Parche) -ne $null 
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
write-host "Esperando Resultados"
write-host ""

# Get all the running jobs
$jobs = get-job | ? { $_.state -eq "running" }
$total = $jobs.count
$runningjobs = $jobs.count

# Loop while there are running jobs
while($runningjobs -gt 0) {
# Update progress based on how many jobs are done yet.
$percent=[math]::Round((($total-$runningjobs)/$total * 100),2)
write-progress -activity "Realizando consultas WMI" -status "Progreso: $percent%" -percentcomplete (($total-$runningjobs)/$total*100)

# After updating the progress bar, get current job count
$runningjobs = (get-job | ? { $_.state -eq "running" }).Count
}
Get-Job | Wait-Job | Receive-Job | out-file -filepath $Log -Append




Write-host "PROCESO FINALIZADO, REVISAR Html o el siguiente log" $Log

Pause

#Pasarlo a HTML
Get-Content $Log | ConvertFrom-Csv -Delimiter ';' | ConvertTo-Html | Out-String | out-file -filepath "$scriptDir\ReportKB.html"
Invoke-Expression  "$scriptDir\ReportKB.html"

