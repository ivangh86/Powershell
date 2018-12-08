# FindShutnotDeprovVMs.ps1
# John Savill 7/27/2017
#
# Script should be run every 20 minutes and is designed to deprovision VMs from Azure that have been stopped for more
# than 20 minutes. 
# On each execution it loads in the previous executions list of VMs that were stopped and had not been restarted/stopped within
# the last 20 minutes. If the VM on this execution is still stopped and has not been restarted/stopped within that time
# the VM is stopped so it is no longer charged.

#Modify this as needed. Note there are different ways to store the credential such as at
#http://windowsitpro.com/development/save-password-securely-use-powershell
$user = "account"
$password = 'password'
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($user, $securePassword)  
Login-AzureRmAccount -Credential $cred #ARM logon
Select-AzureRmSubscription -SubscriptionId <subscription ID>


$VMStateFile = 'D:\temp\AzureStoppedVMs.csv' 
$RGs = Get-AzureRMResourceGroup
$VMList = @()
#read in previous list to $VMListPrev
if(Test-Path $VMStateFile) #if file exists
{
    $VMListPrev = Import-Csv $VMStateFile
}
else
{
    $VMListPrev = $null
}
foreach($RG in $RGs)
{
    $VMs = Get-AzureRmVM -ResourceGroupName $RG.ResourceGroupName
    foreach($VM in $VMs)
    {
        $VMDetail = Get-AzureRmVM -ResourceGroupName $RG.ResourceGroupName -Name $VM.Name -Status
        foreach ($VMStatus in $VMDetail.Statuses)
        { 
            if($VMStatus.Code.CompareTo("PowerState/stopped") -eq 0) #if VM is stopped but not deprovisioned
            {
                $tarvmname = $VM.Name.ToLower()
                $tarrg = $VM.ResourceGroupName
                $tarVMObj = [PSCustomObject]@{"Name"=$tarvmname; "ResourceGroup" = $tarrg}
                Write-Output "VM $tarvmname in RG $tarrg is stopped but not deprovisioned"
                
                #Check the VM has not been started in last 20 minutes as it could have been started then stopped again
                #in which case we should not take action
                $StartedRecent = $false
                #Get all start logs. Could technically optimize this and move out of loop if group all VMs by RG
                $logs = Get-AzureRmLog -ResourceGroup $tarrg -StartTime (Get-Date).AddMinutes(-20) | 
                    Where-Object OperationName -eq Microsoft.Compute/virtualMachines/start/action
                #Need to check to only do this if there are logs
                if($logs -ne $null)
                {
                    foreach($log in $logs)
                    {
                        if($log.Authorization.Scope.ToLower().EndsWith($tarvmname)) #if its this VM
                        {
                            $StartedRecent = $true
                            Write-Output "VM $tarvmname in RG $tarrg has been started in time interval. No action will be taken."
                        }
                    }
                }
                if(!$StartedRecent)
                {
                    #Check if the VM is in the list of VMs that were down previously and if so stop it
                    if(($VMListPrev | ? {$_.Name -eq $tarvmname –and $_.ResourceGroup  -eq $tarrg}) -ne $null)  #if found a match
                    {
                        Write-Output "VM $tarvmname in RG $tarrg was found in previous execution and will be stopped"
                        Stop-AzureRmVM -Name $tarvmname -ResourceGroupName $tarrg -Force
                    }
                    else
                    {
                        #If not add to new list which will be checked next time
                        $VMList += $tarVMObj
                    }
                }    
            }
        }
    }
}
#Save the current list to file for checking next time
$VMList | Export-csv $VMStateFile -notypeinformation 