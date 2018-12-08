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

# Obtengo las máquinas del grupo 
(Get-ADGroupMember $Group | select-object).name > C:\prueba\equipos.txt

#Variable del fichero con los miembros
$fichero = "C:\prueba\equipos.txt"

#Propiedades del fichero
$PropFichero = dir $fichero 

#Obtenemos el contenido del fichero
$Contfichero = Get-Content $fichero

#Revisamos si estamos en la granja de Desarrollo
$FarmDes = Get-XAFarm
#$FarmDes.FarmName

#Revisamos que el fichero tenga fecha escritura de hoy y que estemos en la granja de Desarollo

if (($PropFichero.LastWriteTime -gt $today) -and ($FarmDes.FarmName -eq "Sabadell Desarrollo"))

    {

    write-host "Es cierto, existe fichero con fecha de hoy y estas en la granja Desarrollo"

    #Añadimos las máquinas del grupo al Load Balancing Policy indicado en el apartado de Nombre de Clientes
    Set-XALoadBalancingPolicyFilter -PolicyName "BalanceoEquipo" -AllowedClientNames $Contfichero
    
    }


Else {
    write-host "no es cierto, no existe fichero con fecha de hoy o no estas en granja Desarrollo"
    }


#Select-Object name | Export-csv C:\prueba\equipos.txt -NoTypeInformation
