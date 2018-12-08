$Servers_Solvia_CBS = Get-XAServer -WorkerGroupName "SB_PROD_SOLVIA_CBS" |Select-Object -ExpandProperty "ServerName"  | Sort-Object


$Servers_Solvia_SF =Get-XAServer -WorkerGroupName "SB_PROD_SOLVIA_SF"  |Select-Object -ExpandProperty "ServerName" | Sort-Object 


Write-host "Servidores SOLVIA CBS"
Foreach ($ServersCBS in $Servers_Solvia_CBS){
write-host $ServersCBS ":"(Get-XALoadEvaluator -ServerName $ServersCBS | Select-Object -ExpandProperty LoadEvaluatorName)
}

Write-host  ""
Write-host  ""

Write-host "Servidores SOLVIA SF"
Foreach ($ServersSF in $Servers_Solvia_SF){
write-host $ServersSF ":" (Get-XALoadEvaluator -ServerName $ServersSF | Select-Object  -ExpandProperty LoadEvaluatorName)
}