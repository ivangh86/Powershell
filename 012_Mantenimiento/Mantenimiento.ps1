asnp Citrix*

$ServerList = Get-Content "C:\Temp\Listserver.txt"

foreach ($Server in $Serverlist){

(
Get-XAServer $Server).ServerName
Get-XAServer $Server | Set-XAServerLoadEvaluator -LoadEvaluatorName Mantenimiento
Write-Host "Puesto en Mantenimeinto"
Get-XAServerLoad  $Server | Get-XALoadEvaluator 

}