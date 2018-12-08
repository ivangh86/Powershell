###################################################
# Autor: Ivan Gomez
# Fecha: 30/11/2018
#
# Descripción: Reinicio controlado de los DDC. Paramos el Servicio "CitrixHighAvailabilityService" y a continuación revisamos si existen los ficheros HAImport en la ruta:
# C:\Windows\ServiceProfiles\NetworkService
# Dejamos log de control y revision "MachineName-RestartService.log" en C:\temp (Se genera en cada Reinicio) 
#
###################################################


# Declaramos variables logs y ficheros
$Logfile = "C:\temp\$(gc env:computername)-RestartService.log"
$ubicacion = "C:\Windows\ServiceProfiles\NetworkService"
$File_HaImportDatabaseName = "C:\temp\New1.txt"
$File_HaImportDatabaseName_log = "C:\temp\New2.txt"

#Borramos Log antiguo si existe
If (Test-Path $Logfile) {

Del $Logfile
}

#Funcion para escribir en Log
Function LogWrite
{
Param ([string]$logstring)
$timestamp = Get-Date -Format "dd/MM/yyyy hh:mm:ss"
$logMessage = $timestamp + " " + $logstring
 
Add-content $Logfile -value $logMessage
}

#Funcion parar servicio con Timeout
function Stop-ServiceWithTimeout ([string] $name, [int] $timeoutSeconds) {
    $timespan = New-Object -TypeName System.Timespan -ArgumentList 0,0,$timeoutSeconds
    $svc = Get-Service -Name $name
    if ($svc -eq $null) { return "Servicio no encontrado"}
    if ($svc.Status -eq "Stopped") { write-host "Servicio ya estaba parado" }
    $svc.Stop()
    try {
        $svc.WaitForStatus([ServiceProcess.ServiceControllerStatus]::Stopped, $timespan)
    }
    catch [ServiceProcess.TimeoutException] {
        Write-Verbose "Timeout stopping service $($svc.Name)"
        return "No se ha podido parar en el tiempo indicado"
    }
    return "Servicio Parado"
}


#Comenzamos el Script 
LogWrite "Iniciando Script Reinicio DDC"

#Servicio a tener en cuenta Citrix High Availability
$servicename =  "ctxProfile"
$Service_running = get-service $servicename | where {$_.status -eq 'running'}
 
if ($Service_running) {
    LogWrite "El servicio $servicename esta arrancado. Paramos el servicio..."
    write-host "El servicio $servicename esta arrancado. Paramos el servicio..."
    Stop-ServiceWithTimeout  $servicename "300"
    $status = get-service $servicename | where {$_.status -eq 'Stopped'}
    if (!$status) { LogWrite "El servicio se ha parado correctamente" }
    else {LogWrite "El servicio ha sido parado correctamente"}
    } 

else {
LogWrite "El servicio estaba parado. Se finaliza la ejecucion del script"
write-host "El servicio estaba parado. Se finaliza la ejecucion del script"
Exit
}


#Comprobamos si existen los ficheros HAImport y en ese caso los eliminamos.
LogWrite "Comprobamos que no existan los ficheros HAImport en: $ubicacion" 

If (Test-Path $File_HaImportDatabaseName){
    LogWrite "Existe fichero $File_HaImportDatabaseName , procedemos a borrarlo"
    Del $File_HaImportDatabaseName
    Start-Sleep 2
    If ((Test-Path $File_HaImportDatabaseName) -eq $False){LogWrite "Fichero $File_HaImportDatabaseName borrado"}
    Else {LogWrite "No se ha prodido borrar el fichero $File_HaImportDatabaseName"}
    }

Else {LogWrite "No existe $File_HaImportDatabaseName"}

If (Test-Path $File_HaImportDatabaseName_log){
    LogWrite "Existe fichero $File_HaImportDatabaseName_log , procedemos a borrarlo"
    Del $File_HaImportDatabaseName_log
    Start-Sleep 2
    If ((Test-Path $File_HaImportDatabaseName_log) -eq $False){LogWrite "Fichero $File_HaImportDatabaseName_log borrado"}
    Else {LogWrite "No se ha prodido borrar el fichero $File_HaImportDatabaseName_log"}
    }

Else {LogWrite "No existe $File_HaImportDatabaseName_log"}


