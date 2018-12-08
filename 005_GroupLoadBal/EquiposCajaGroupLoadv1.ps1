﻿############################################################################### 
#        
#        
#        Name:          EquiposCajaGroupLoad1.ps1
#        Author:        Iván Gómez 
#        Date:          25/04/2017 
#        Description:   Añadimos las máquinas del grupo "Grupo Thinclients Oficina Piloto Load Balancing DES" 
#                       al Load Balancing Policy indicado en el apartado de Nombre de Clientes 
#                           
#
############################################################################### 
###############################################################################

#Modulos Necesarios
#
asnp Citrix*
Import-Module ActiveDirectory

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path 
$scriptDir >> "$scriptDir\equiposCajaGroupLoad.log"

#Definimos el grupo de Ad donde estan las máquinas
$Group = "Grupo Thinclients Oficina Piloto Load Balancing DES"

#Definimos la fecha de hoy (Como mucho una hora antes)
$today = (Get-Date).AddHours(-1)

$today

# Obtengo las máquinas del grupo 
(Get-ADGroupMember $Group | select-object).name > "$scriptDir\equipos.txt"

#Variable del fichero con los miembros
$fichero = "$scriptDir\equipos.txt"

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

        try {
            Set-XALoadBalancingPolicyFilter -PolicyName "BalanceoEquipo" -AllowedClientNames $Contfichero
            }
            
        Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            Get-Date | Add-Content "$scriptDir\ErrorequiposCajaGroupLoad.log"
            echo "" | Add-Content "$scriptDir\ErrorequiposCajaGroupLoad.log"
            echo "Se ha producido un error al añadir máquinas del grupo al Load Balancing Policy" | Add-Content "$scriptDir\ErrorequiposCajaGroupLoad.log"
            echo "" | Add-Content "$scriptDir\ErrorequiposCajaGroupLoad.log"
            echo "Mensaje:$ErrorMessage" | Add-Content "$scriptDir\ErrorequiposCajaGroupLoad.log"
            echo "Objeto:$FailedItem"    | Add-Content "$scriptDir\ErrorequiposCajaGroupLoad.log"
            echo "" | Add-Content "$scriptDir\ErrorequiposCajaGroupLoad.log"
            
               }
      }

Else {
    write-host "no es cierto, no existe fichero con fecha de hoy o no estas en granja Desarrollo"
    }


#Select-Object name | Export-csv C:\prueba\equipos.txt -NoTypeInformation
