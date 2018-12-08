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
#New-AzureRmStorageAccount -ResourceGroupName $rgname -Name $stoname -Location $loc -Type $stotype
#$stoaccount = Get-AzureRmStorageAccount -ResourceGroupName $rgname -Name $stoname 

#Create VM Object
$vm = New-AzureRmVMConfig -VMName $vmname -VMSize $vmsize

# Setup Networking
#$subnet = New-AzureRmVirtualNetworkSubnetConfig -Name ('subnet' + $rgname)  -AddressPrefix "10.0.0.0/24"
New-AzureRmVirtualNetwork -ResourceGroupName $rgname -Name $virtnetname -Location $loc -AddressPrefix 10.0.0.0/16 -DnsServer "10.1.1.1" -Subnet $subnet
$vnet = New-AzureRmVirtualNetwork -Force -Name ('vnet'+ $rgname) -ResourceGroupName $rgname -Location $loc -AddressPrefix 10.0.0.0/16 -DnsServer 10.1.1.1 -Subnet $subnet
$vnet = Get-AzureRmVirtualNetwork -Name $virtnetname -ResourceGroupName $rgname
$subnetID = $vnet.Subnets[0].Id

$pip = New-AzureRmPublicIpAddress -ResourceGroupName $rgname -Name "vip1" `
-Location $loc -AllocationMethod Dynamic -DomainNameLabel $vmname.ToLower()

$nic = New-AzureRmNetworkInterface -Force -Name ('nic' + $vmname) `
-ResourceGroupName $rgname -Location $loc -SubnetId $subnetID -PrivateIPAddress 10.0.0.13 -DnsServer 10.0.0.13 `
-PublicIPAddressID $pip.Id


# Add NIC to VM
$vm = Add-AzureRmVMNetworkInterface -VM $vm -Id $nic.Id

# Using an Azure image
$osDiskName = $vmname+'osDisk'
$osDiskCaching = 'ReadWrite'
$osDiskVhdUri = "https://$stoname.blob.core.windows.net/vhds/"+$vmname+"-OS.vhd"
# Setup OS & Image
$user = "localadmin"
$password = 'Pa55word5'
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($user,$securePassword)

#Get the latest image using version
$loc = 'northeurope'
$AzureImageSku = Get-AzureRmVMImage -Location 'northeurope' -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2012-R2-Datacenter" 
# Cogemos la mas nueva
$AzureImageSku = $AzureImageSku |Sort-Object Version -Descending 
$AzureImageSku[0]
$AzureImage= $AzureImageSku[0]

# Ejemplo de imagenes ya publicadas
# Get-AzureRmVMImagePublisher -Location $loc
# Get-AzureRmVMImageOffer -Location $loc -PublisherName "MicrosoftWindowsServer" 
# Get-AzureRmVMImageSku -Location $loc -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer"

$vm = Set-AzureRmVMOperatingSystem -VM $vm -Windows -ComputerName $vmname -Credential $cred 
$vm = Set-AzureRmVMSourceImage -VM $vm -PublisherName $AzureImage.PublisherName -Offer $AzureImage.Offer -Skus $AzureImage.Skus -Version $AzureImage.Version
$vm = Set-AzureRmVMOSDisk -VM $vm -VhdUri $osDiskVhdUri -Name $osDiskName -CreateOption fromImage -Caching $osDiskCaching


# To create VM from a custom VHD replace with https://azure.microsoft.com/en-en/documentation/articles/virtual-machines-windows
#$sourcevhd = "https://$stoname.blob.core.windows.net/vhds/win2012r2custom-OS.vhd"
#$osDiskVhdUri = "https://$stoname.blob.core.windows.net/vhds/"+$vmname+"-OS.vhd"
#$vm = Set-AzureRmVMOSDisk -VM $vm -Name $osDiskName -VhdUri $osDiskVhdUri - -CreateOption 'fromImage' -SourceImageUri $sourcevhd -Windows -Caching $osDiskCaching
 
# To create VM from an existing disk replace Using an Azure image commands with:
#$vm = Set-AzureRmVMOSDisk -VM $vm -VhdUri $osDiskVhdUri -Name $osDiskName -CreateOption 'attach' -Windows -Caching $osDiskCaching


#Add a data disk
#$dataDiskVhdUri = "https://$stoname.blob.core.windows.net/vhds/"+$vmname+"-Data.vhd"
#$dataDiskName = $vmname+'_dataDisk'
#$vm = Add-AzureRmVMDataDisk -VM $vm -Name $dataDiskName -Caching None -CreateOption 'attach' -DiskSizeInGB 1023 -VhdUri $dataDiskVhdUri


# Create virtual machine

New-AzureRmVM -ResourceGroupName $rgname -Location $loc -VM $vm