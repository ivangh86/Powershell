$RGs = Get-AzureRmResourceGroup

$VMList= @()
foreach ($RG in $RGs)
{
    $VMs = Get-AzureRmVM -ResourceGroupName $RG.ResourceGroupName
    foreach ($VM in $VMs)
    {
        $VMDetail = Get-AzureRmVM -ResourceGroupName $RG.ResourceGroupName -Name $VM.Name -Status
        foreach ($VMStatus in $VMDetail.Statuses)
        {
            if ($VMStatus.Code.CompareTo("ProvisioningState/succeeded") -ne 0 ) # don't want to the provsioning status
            {
                $VMStatusDetail = $VMStatus.DisplayStatus

            }
        }
    }   
    $VMList += [PSCustomObject]@{
        "Name"=$VM.Name; "Status"=$VMStatusDetail; "ResourceGroup" = $VM.ResourceGroupName
    }

}
$VMList


#############################################

#View all Azure locations
Get-AzureRmLocation

$azurelocations = Get-AzureRmLocation
$resources = Get-AzureRmResourceProvider -ProviderNamespace 'Microsoft.Compute'
$resources.ResourceTypes.Where{($_.ResourceTypeName -eq 'virtualMachines')}.locations

Get-AzureRmVMSize -Location "North Europe"
#############################################

# Actions on all VMs in RG
$rgname = 'RGNEInfra'
$VMs = Get-AzureRMVm -ResourceGroupName $rgname

foreach ($VM in $VMs)   
    {
        Stop-AzureRmVM -Name $VM.Name -ResourceGroupName $rgname -Force
    }   
