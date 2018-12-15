#region workflow 
workflow MyWorkflow {Write-Output -InputObject "Hello from workflow!"}
MyWorkflow


workflow LongWorkflow {
    
    Write-Output -InputObject "Loading some information..."
    Start-Sleep -Seconds 10
    Checkpoint-Workflow
    Write-Output -InputObject "Performing process list..."
    Get-Process -PSPersist $true #thisadds checkpoint
    Start-Sleep -Seconds 10
    Checkpoint-Workflow
    Write-Output -InputObject "Cleaning up..."
    Start-Sleep -Seconds 10
}


LongWorkflow -AsJob -JobName LongWF -PSPersist $true
Suspend-Job LongWF
Get-Job LongWF
Receive-Job LongWF -Keep
Resume-Job LongWF
Get-Job LongWF
Receive-Job LongWF -keep
Remove-Job LongWF # removed the saved state of the job