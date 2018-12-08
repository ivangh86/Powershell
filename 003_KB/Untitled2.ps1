
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path


get-date > "$scriptDir\prueba2.txt"

$Log = "$scriptDir\prueba2.txt"
$Servers = Get-Content "$scriptDir\ServerList.txt"
#$Servers = Get-BrokerDesktop | Where-Object {$_.DNSName -like "*vdnbpv0001d*" } |Select-Object DNSName,AgentVersion,OSType,DesktopGroupname,DesktopKind,RegistrationState



$check_hotfix = {
    param ($Server)
    $Parches = 'KB4012212','KB4012213','KB4012214','KB4012598'
    foreach ($Parche in $Parches) {
    $is_patched = (get-wmiobject -class win32_quickfixengineering -computer $Server | select-object HotFixID | Select-String $Parche) -ne $null 
    if ($is_patched) { 
        Write-Output ($Server+ " Exist " +$Parche) 
    } else { 
        Write-Output ($Server+ " Falta " +$Parche)
    }
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
Get-Job | Wait-Job | Receive-Job | out-file -filepath $Log

Write-host "PROCESO FINALIZADO, REVISAR" $Log




#Invoke-Expression  $scriptDir\prueba2.html