#asnp Citrix*

#Get-AcctDBConnection
Get-ConfigRegisteredServiceInstance -ServiceType 'Broker' -Version 1 -InterfaceType 'SDK' -MaxRecordCount 2147483647 -AdminAddress 'localhost' | Test-ConfigServiceInstanceAvailability -MaxDelaySeconds 1 -ForceWaitForOneOfEachType -AdminAddress 'localhost'



#Get-BrokerServiceAddedCapability

#Get-ConfigRegisteredServiceInstance

Get-ConfigServiceStatus

