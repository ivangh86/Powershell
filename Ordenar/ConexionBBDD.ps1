
if ((Get-PSSnapin "Citrix.Broker.Admin.*" -EA silentlycontinue) -eq $null) {
try { Add-PSSnapin Citrix* -ErrorAction Stop }
catch { write-error "Error loading Citrix PowerShell snapin"; Return }
}


$StringBBDD = Get-BrokerDBConnection

$ServiceStatus = Test-BrokerDBConnection "$StringBBDD"  | Select-Object -ExpandProperty "ServiceStatus"
 

$Brokers = (Get-BrokerController).DNSName


foreach ($Broker in $Brokers) {


if ($ServiceStatus -eq "OK")  
    {

    write-host -ForegroundColor Green "$Broker : Conexión con la BBDD Correcta"
    write-host ""
    }

else 
    { 
    write-host -ForegroundColor Red "$Broker Problemas con la conexión a la BBDD"
    write-host ""
    Write-host -ForegroundColor Red "Puede probar el comando Get-BrokerDBConnection: para sacar el String de la conexión"  
    write-host -ForegroundColor Red "y pasarselo como parametro al comando Test-BrokerDBConnection 'String' "
    write-host ""
    }

        
}

read-host "Press Enter to continue..."
