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
    
# Azure Store Info

$storacts = Get-AzureRmStorageAccount
foreach ($storact in $storacts)
{
    Write-Output "$($storact.StorageAccountName), $($storact.sku.Name),  $($storact.PrimaryLocation), $($storact.SecondaryLocation)"
}


# View Availability Set information for VMs
# Debes tener un conjunto de disponibilidad creado (Se crea al crear la maquina)
Get-AzureRmVM | Where-Object {$_.ResourceGroupName -eq 'RGNEInfra' } | Format-Table Name,ResourceGroupName -AutoSize
$AS = Get-AzureRmAvailabilitySet -ResourceGroupName 'RGNEInfra' 
$AS.VirtualMachinesReferences | ForEach-Object {$VMResource = (Get-AzureRmResource -Id $_.Id);
    $VM = Get-AzureRMVM -Name $VMResource.Name -ResourceGroup $VMResource.ResourceGroupName -Status; 
    [PSCustomObject]@{ "Name" = $VM.Name; "FaultDomain"=$VM.PlatformFaultDomain; "UpateDomain"=$VM.PlatformUpdateDomain;
    }
}

# Capture an image @azcaptimage

Stop-AzureRmVM -ResourceGroupName "RGNEInfra" -Name 'savazudc01' -Force
Set-AzureRmVM -ResourceGroupName "RGNEInfra" -Name 'savazudc01' -Generalized
Save-AzureRmVMImage -ResourceGroupName "RGNEInfra" -Name "savazudc01" -DestinationContainerName "custtemp" -VHDNamePrefix "template" -Path "e:\temp\CapTemple.json" 

# Inject Powershell into a VM
$rgname = 'RGNEInfra'
$vmname = 'savazudc01';
$loc = 'northeurope'

$ScriptBlobAccount = 'sanegrsinfra'
$ScriptBlobKey = "U9tKrSQ1iK8UpVZsZb/1sBqeD3SeOxyH0QNTizNhkNZXudha+CQDU4RWUSZKf88jRAVHB/l28p3oPXipDOM6MQ=="
$ScriptBlobURL =  "https://sanegrsinfra.blob.core.windows.net/scripts/"

$ScriptName = "outcomputername.ps1"
$ExtensionName = "outcomputername"
$ExtensionType = 'CustomScriptExtension'
$Publisher = 'Microsoft.Compute'
$Version = '1.8'
$timestamp = (Get-Date).Ticks

$ScriptLocation = $ScriptBlobURL + $ScriptName
$ScriptExe = ".\$ScriptName"

$PrivateConfiguration = @{"storageAccountName" = "ScriptBlobAccount"; "storageAccountKey" = "$ScriptBlobKey"}
$PublicConfiguration = @{"fileUris" = [Object[]]"$ScriptLocation";"timestamp = $timestamp"; "commandToExecute"  = "powershell.exe -ExecutionPolicy Unrestricted"

Set-AzureRmVMExtension -ResourceGroupName "RGNEInfra" -VMName $vmname -Location $loc `
-Name $ExtensionName -Publisher $Publisher -ExtensionType $ExtensionType -TypeHandlerVersion $Version `
-Settings $PublicConfiguration #-ProtectedSettings $PrivateConfiguration

# Should see its name as the Message as thats what the PowerShell runs
((Get-AzureRmVM -Name $vmname -ResourceGroupName $rgname -Status).Extensions |Where-Object {$_.Name -eq $ExtensionName}).Substatuses