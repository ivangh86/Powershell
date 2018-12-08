$Servers = Get-XAServer




foreach ($Server in $Servers){
    
    $Row1 = (Get-XAServer $Server).ServerName
    $Row2 = (Get-XAServer $Server | Get-XALoadEvaluator).LoadEvaluatorName
    $Row1,$Row2 | format-table >> "C:\temp\ListadoLoadEvaluator.txt"
    
   
}

#foreach ($item in $Consulta) {

#$Consulta.LoadEvaluatorName

#} 