
asnp Citrix*

$servers = get-content "\\cortex\soft_citrix\Temporal\Ivan_Gomez\serverlist.txt"



Function Revisa_Servicio {

Param (
  [Parameter(Mandatory=$True)]
   [Array]$servidores
)


foreach ($server in $servers){

Write-Host "*********" $server "**********"


#Primer analisis de la maquina Bloque 1
write-Host "Estado Inicial del Servicio IMA y peso del archivo imalhc.mdb, asi como el evaluador de carga (Bloque 1)" 
$servicio = Invoke-Command -ComputerName $server -scriptBlock {Get-Service imaservice }
$servicio
$fichero = Invoke-Command -ComputerName $server -scriptBlock {Get-ChildItem "C:\Program Files (x86)\Citrix\Independent Management Architecture\imalhc.mdb"}
$fichero
$Load = (Get-XALoadEvaluator -ServerName $server).LoadEvaluatorName
write-host "Load evaluator"
$Load

write-host "**************************************"



#Revisamos si la maquina esta en Mantenimiento para reiniciar el IMA, en caso contrario damos la opcion de ponerlo en mantenimiento.

If ($Load -eq "Mantenimiento_Error4001"){

# Comprobamos si el servicio esta arrancado  bloque 2

    if ($servicio.status -eq "Running"){
    write-host "Paramos Servicio IMA"
    Invoke-Command -ComputerName $server -scriptBlock {Stop-Service imaservice -force }

    write-host "Comprobamos que esta parado el Servicio IMA"
    Invoke-Command -ComputerName $server -scriptBlock {Get-Service imaservice }

    write-host "Arrancamos el Servicio IMA"
    Invoke-Command -ComputerName $server -scriptBlock {Start-Service imaservice }

    write-host "Comprobamos que esta arrancado el Servicio IMA y el peso del archivo imalhc.mdb (Bloque 1)"
    Invoke-Command -ComputerName $server -scriptBlock {Get-Service imaservice }
    Invoke-Command -ComputerName $server -scriptBlock {Get-ChildItem "C:\Program Files (x86)\Citrix\Independent Management Architecture\imalhc.mdb"}

    write-host "Presione cualquier tecla para continuar con el siguiente Servidor de la lista" -BackgroundColor red -ForegroundColor Black

    pause

    }

# Si no esta arrancado directamente lo intentamos levantar

    else {

    write-host "El servicio IMA ya estaba parado"

    
    write-host "Arrancamos el Servicio IMA"
    Invoke-Command -ComputerName $server -scriptBlock {Start-Service imaservice }

    write-host "Comprobamos que esta arrancado el Servicio IMA y el peso del archivo imalhc.mdb asi como el Evaluador de Carga (Bloque 3)"
    Invoke-Command -ComputerName $server -scriptBlock {Get-Service imaservice }
    Invoke-Command -ComputerName $server -scriptBlock {Get-ChildItem "C:\Program Files (x86)\Citrix\Independent Management Architecture\imalhc.mdb"}

    write-host "Presione cualquier tecla para continuar con el siguiente Servidor de la lista" -BackgroundColor Blue -ForegroundColor Black

    pause

    }

}


else { write-host $Server "No esta en Mantenimiento_Error4001 "}



#$optionSelected = Read-Host "El servidor no esta en mantenimiento. Quiere ponerlo Mantenimiento_Error4001 para correr el Script 'Yes(y)/No(n)/Exit(e)', default = No)"}


#switch ($optionSelected) {
#
#  { 'Yes', 'y', '1', 'True', 'T' -contains $_ } {   write-host "Has elegido SI, pocredemos a ponerlo en mantenimiento y revisar de nuevo" -BackgroundColor Magenta -ForegroundColor Black
#   Set-XAServerLoadEvaluator -ServerName $server -LoadEvaluatorName "Mantenimiento_Error4001" 
#    $server

    
   #Revisa_Servicio  -servidores $server}

#  { 'No', 'n', '0', 'False', 'F' -contains $_ } { 
#  write-host "Has elegido NO, pasamos al siguiente Servidor de la lista"}

#  { 'Exit', 'e' -contains $_ } { 
#  write-host "Has elegido Exit, Salimos del proceso" 
#  break}

#  default { $boolOption = $false }
#  }

#}



}

}


Revisa_Servicio  -servidores $servers





