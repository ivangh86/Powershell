$encryptedPassword = ConvertFrom-SecureString (ConvertTo-SecureString -AsPlainText -Force "Password123")
$encryptedPassword
$passwordAsSecureString = ConvertTo-SecureString $encryptedPassword
$passwordAsSecureString

$cred = ""

$credpath = "c:\temp\MyCredential.xml"
New-Object System.Management.Automation.PSCredential("gomez-hernandez@dxc.com", (ConvertTo-SecureString -AsPlainText -Force "Eeupdc13!")) | Export-Clixml $credpath


$credpath = "c:\temp\MyCredential.xml"
$cred = import-clixml -path $credpath

Login-AzureRmAccount -Credential $cred



$ID = (Get-AzureRmSubscription).Id

Select-AzureRmSubscription -SubscriptionID $ID

Get-AzureRmStorageAccount |ft StorageAccountName,ResourceGroupName
