 
 $Servers = Get-Content \\serverhqr1\soft_ctx\Team\Ivan\BrokerHostingPowerAction\ServerList.txt
 
 foreach ($Server in $Servers) {

 write-host "Arrancando $Server"
 New-BrokerHostingPowerAction -MachineName $Server -Action TurnOn
 
  }

   foreach ($Server in $Servers) {

 write-host "Comprobando $Server"
 Get-BrokerHostingPowerAction -HostedMachineName $Server
 
  }



