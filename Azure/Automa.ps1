$RGs = Get-AzureRmResourceGroup

$VMList= @ ()

foreach ($RG in $RGs)
{

    $VMs = Get-AzureRmVM -ResourceGroupName $RG.ResourceGroupName
    foreach ($VM in VMs)

    {
        $VMDetail = Get-AzureRmVM -ResourceGroupName $RG.ResourceGroupName -Name $VM.Name -Status


    }


}
