#Create a VM @azurermvmcreate

$rgname = 'RGNEInfra'
$vmsize = 'Standard_A2';
$vmname = 'savazudc01';
$virtnetname = 'Vnet115'
#Setup Storage
$stoname = 'sanegrsinfra';
$stotype = 'Standard_GRS';
$loc = 'northeurope'

#New-AzureRmResourceGroup -Name $rgname -Location $loc -Verbose -Force -ErrorAction Stop

#New-AzureRmVirtualNetwork -ResourceGroupName $rgname -Name $virtnetname -Location $loc -AddressPrefix 10.1.0.0/24 

#New-AzureRmStorageAccount -ResourceGroupName $rgname -Name $stoname -Location $loc -Type $stotype
#$stoaccount = Get-AzureRmStorageAccount -ResourceGroupName $rgname -Name $stoname 

#Create VM Object
#$vm = New-AzureRmVMConfig -VMName $vmname -VMSize $vmsize

# Setup Networking
# $subnet = New-AzureRmVirtualNetworkSubnetConfig -Name ('subnet' + $rgname)  -AddressPrefix "10.0.0.0/24"
# New-AzureRmVirtualNetwork -Force -Name ('vnet'+ $rgname) -ResourceGroupName $rgname -Location $loc -AddressPrefix 10.0.0.0/16 -DnsServer 10.1.1.1 -Subnet $subnet
$vnet = Get-AzureRmVirtualNetwork -Name $virtnetname -ResourceGroupName $rgname
$subnetID = $vnet.Subnets[0].Id

$pip = New-AzureRmPublicIpAddress -ResourceGroupName $rgname -Name "vip1" `
 -Location $loc -AllocationMethod Dynamic -DomainNameLabel $vmname.ToLower()


 $nic = New-AzureRmNetworkInterface -Force -Name ('nic' + $vmname) -ResourceGroupName $rgname`
 -Location $loc -SubnetID $subnetID




 