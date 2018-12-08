notepad C:\temp\mcf\addtoLB.csv

#Add machines to list to add them to the load balancing policy.
#Remove them from the list to remove them from the policy.

Write-Host "Edit notepad file and add new computer names. Press any key to continue ..."
Write-Host "Do not remove machines unless they no longer need to be part of Load Balancer..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

#import csv
$fichero = Import-CSV C:\temp\mcf\addtoLB.csv
#default entry needed for formatting of array
$machinelist = 'zzz-DefaultEntry'
#Add machines from list to variable machinelist
ForEach ($record in $fichero){
$machinelist += ", "+ $record.AddMachinesToList

}
$machinelistadd = $machinelist -split ", "
#add machines from list to Load Balancing Policy

if 
Set-XALoadBalancingPolicyFilter -PolicyName "BalanceoEquipo" -AllowedClientNames $machinelistadd