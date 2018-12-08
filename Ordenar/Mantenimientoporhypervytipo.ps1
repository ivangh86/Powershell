Get-BrokerMachine -MaxRecordCount 9999 | where-object 
 { ($_.HostedmachineName -like "wcpnhidvp*") -and 
 (($_.HostingServerName -eq "WFPNHIDHY063.tsb.banks.adgbs.com")  -or 
 ($_.HostingServerName -eq "WFPNHIDHY064.tsb.banks.adgbs.com")  -or 
 ($_.HostingServerName -eq "WFPNHIDHY065.tsb.banks.adgbs.com")  -or  
 ($_.HostingServerName -eq "WFPNHIDHY066.tsb.banks.adgbs.com"))} | Set-BrokerMachine -InMaintenanceMode $True #|select HostedMachineName,HostingServerName,InMaintenanceMode | ft -AutoSize > "C:\temp\MantenimientoTSB.txt"

