############################################################################### 
#        Author: Iván Gómez 
#        Date: 25/04/2017 
#               
#        imported in SQL Table 
#               Modified: 
############################################################################### 
########################Add Quest Shell########################################

#Modulos Necesarios
#
asnp Citrix*
Import-Module ActiveDirectory

$Group = "Grupo Thinclients Oficina Piloto Load Balancing DES"
$today = (Get-Date).AddHours(-1)
$today
$fichero = "C:\prueba\equipos.txt"
$Contfichero = Get-Content $fichero
$PropFichero = dir $fichero 

$FarmDes = Get-XAFarm
#$FarmDes.FarmName


(Get-ADGroupMember $Group | select-object).name > C:\prueba\equipos.txt



if (($PropFichero.LastWriteTime -gt $today) -and ($FarmDes.FarmName -eq "Sabadell Desarrollo"))

{

write-host "Es cierto, existe fichero con fecha de hoy y estas en la granja Desarrollo"

    $machine = ""
    $machinelist =""
    ForEach ($record in $Contfichero){
    $machinelist +=$record
    $machinelist +=", "
    }
    
    $machinelist > C:\prueba\MachineList.txt
    $machinelist = $machinelist -replace ".$"
    $machinelist = $machinelist -replace ".$"
    $machinelist
     
    Set-XALoadBalancingPolicyFilter -PolicyName "BalanceoEquipo" -AllowedClientNames $machinelist

 }


Else {
write-host "no es cierto, no existe fichero con fecha de hoy o no estas en granja Desarrollo"
}


#Select-Object name | Export-csv C:\prueba\equipos.txt -NoTypeInformation
