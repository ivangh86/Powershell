###################################################
#
#   Autor:        IVAN GOMEZ
#   Fecha:        24/10/2017
#   Departamento: Citrix 
#
#
#
################################################## 



$Fecha = Get-Date

$fecha > "\\capelo\soft_wintel\Citrix\Scripts\Grupos\Log.txt"
$fecha > "\\capelo\soft_wintel\Citrix\Scripts\Grupos\LogErrores.txt"

$Logfile = "\\capelo\soft_wintel\Citrix\Scripts\Grupos\Log.txt"
$LogfileError = "\\capelo\soft_wintel\Citrix\Scripts\Grupos\LogErrores.txt"



Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

Function LogWriteError
{
   Param ([string]$logstring)

   Add-content $LogfileError -value $logstring
}




Function RegulaGruposGTS
{
   Param 
   
   (
   
   [string]$App,
   [string]$GTS,
   [string]$GICONOS
   #$GTS_Users = (Get-ADGroupMember "$GTS").SamAccountName ,
   #$GICONOS_Users = (Get-ADGroupMember "$GICONOS").SamAccountName 
   

   )
   
   
   $GTS_Users = (Get-ADGroupMember "$GTS").SamAccountName
   $GICONOS_Users = (Get-ADGroupMember "$GICONOS").SamAccountName 
   write-host "Aplicacion::::" $App
   write-host "Grupo GTS_Icono::::" $GTS
   write-host "Grupo GICONOS::::" $GICONOS
   write-host "Usuarios GTS_Icono::::" $GTS_Users
   write-host "Usuarios GICONOS::::" $GICONOS_Users
   

write-host "************ INICIO $GTS Aplicación $App *********************"
LogWrite   "************ INICIO $GTS Aplicación $App *********************"


  foreach ($user in $GICONOS_Users ) {

           
TRY {
           
   if ($GTS_Users -notcontains $user )

        {
        
        LogWrite "$user"
        Write-Host "$user no existe, lo añadimos"
        
        #Add-ADGroupMember $GTS –Member $user
        write-host "Modo Test no añade"
        #LogWrite "$user añadido al grupo $GTS"
        write-host "$user añadido al grupo $GTS"

        #LogWrite "Comprobación:"
        
        #LogWrite ((GET-ADUSER –Identity $user –Properties MemberOf | Select-Object MemberOf).MemberOf -match  $GTS)
        #LogWrite "Modo Test no Comprueba"
        
        }

    else {  

        #LogWrite "$user Ya existe"
        Write-Host "$user Ya existe"
    
        }
   
    }

CATCH 

    {

    Write-host "Se ha producido un problema con el usuario $user"
    LogWriteError  "Se ha producido un problema con el usuario $user"
    $ErrorMessage = $_.Exception.Message
    LogWriteError $ErrorMessage
    $FailedItem = $_.Exception.ItemName
    LogWriteError $FailedItem
    }

}

write-host "************FIN $GTS Aplicación $App *********************"
write-host "********************************************************************************************************************"
write-host "********************************************************************************************************************"
LogWrite   "************FIN $GTS Aplicación $App *********************"
LogWrite   "********************************************************************************************************************"
LogWrite   ""
LogWrite   ""
LogWrite   ""



}

#$App3
RegulaGruposGTS  "Adaptiv Risk PRO" "GTS_Icono_Adaptiv_Risk_PRO" "GICONOS SSCC Adaptiv Risk PRO"
#$App4
RegulaGruposGTS  "Putty" "GTS_Icono_putty" "GICONOS SSCC TS PUTTY"
#$App5
RegulaGruposGTS  "Rochade BS" "GTS_Icono_Rochade" "GICONOS SSCC TS Rochade"
#$App6
RegulaGruposGTS  "Lotus Notes 8 - Basic" "GTS_Icono_LotusNotes8" "GICONOS SSCC TS NOTES_6"
#$App7
RegulaGruposGTS  "Acceso ProteoLife PreProduccion" "GTS_Icono_SibisPre_ProteoLife" "GICONOS SSCC TS SIBIS Preproduccion ProteoLife"
#$App8
RegulaGruposGTS  "Emulador 3270 PL" "GTS_Icono_Emulacio_ProteoLife" "GICONOS SSCC TS Emulador 3270 ProteoLife"
#$App9
RegulaGruposGTS  "Reconciliation" "GTS_Icono_Reconciliation" "GICONOS SSCC TS Reconciliation"
#$App10
RegulaGruposGTS  "Soporte Incidencias BS" "GTS_Icono_SSCC_Equipo_Virtual_CAM" "GICONOS SSCC Equipo Virtual CAM"
#$App11
RegulaGruposGTS  "SIBIS-Proteo Preproduccion" "GTS_Icono_Sibis_Proteo_Preproduccion" "GICONOS SSCC TS SIBIS Pre-Produccion"
#$App12
RegulaGruposGTS  "SAGE" "GTS_Icono_SAGE" "GICONOS SSCC TS SAGE"
#$App13
RegulaGruposGTS  "Siebel Web" "GTS_Icono_Siebel" "GICONOS SSCC TS Siebel"
#$App14
RegulaGruposGTS  "SEDAS Web" "GTS_Icono_SedasWeb" "GICONOS SSCC TS Sedas"
#$App15
RegulaGruposGTS  "SIBIS-Proteo Preproduccion Extern" "GTS_Icono_Sibis_PRE_Extern" "GICONOS SSCC TS SIBIS Pre-Produccion"
#$App16
RegulaGruposGTS  "Swift Alliance" "GTS_Icono_Swift_Alliance" "GICONOS SSCC TS Swift"
#$App17
RegulaGruposGTS  "UltraEdit" "GTS_Icono_Ultraedit" "GICONOS SSCC TS UltraEdit32"
#$App18
RegulaGruposGTS  "RochadePL" "GTS_Icono_Rochade_ProteoLife" "GICONOS SSCC TS Rochade Proteo Life"
#$App19
RegulaGruposGTS  "Centro de Servicio Integrado" "GTS_Icono_Centro_de_Servicio_Integrado" "GICONOS SSCC TS Centro de Servicio Integrado"
#$App20
RegulaGruposGTS  "RecAdmin" "GTS_Icono_RecAdmin" "GICONOS SSCC TS RecAdmin"
#$App3
RegulaGruposGTS  "Adaptiv Risk PRO" "GTS_Icono_Adaptiv_Risk_PRO" "GICONOS SSCC Adaptiv Risk PRO"
#$App3
RegulaGruposGTS  "Adaptiv Risk PRO" "GTS_Icono_Adaptiv_Risk_PRO" "GICONOS SSCC Adaptiv Risk PRO"
#$App3
RegulaGruposGTS  "Adaptiv Risk PRO" "GTS_Icono_Adaptiv_Risk_PRO" "GICONOS SSCC Adaptiv Risk PRO"
#$App3
RegulaGruposGTS  "Adaptiv Risk PRO" "GTS_Icono_Adaptiv_Risk_PRO" "GICONOS SSCC Adaptiv Risk PRO"
#$App3
RegulaGruposGTS  "Adaptiv Risk PRO" "GTS_Icono_Adaptiv_Risk_PRO" "GICONOS SSCC Adaptiv Risk PRO"
#$App3
RegulaGruposGTS  "Adaptiv Risk PRO" "GTS_Icono_Adaptiv_Risk_PRO" "GICONOS SSCC Adaptiv Risk PRO"
#$App3
RegulaGruposGTS  "Adaptiv Risk PRO" "GTS_Icono_Adaptiv_Risk_PRO" "GICONOS SSCC Adaptiv Risk PRO"

 


#ok
#$App23 = "Adaptiv Risk PRO"
#$Giconos_23 = "GICONOS SSCC Adaptiv Risk PRO"
#$Giconos_23_Usuarios = (Get-ADGroupMember "$Giconos_23").SamAccountName
#$GTS_23 = "GTS_Icono_Adaptiv_Risk_PRO"
#$GTS_23_Usuarios = (Get-ADGroupMember "$GTS_23").SamAccountName