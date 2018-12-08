###################################################
#
#   Autor:        IVAN GOMEZ
#   Fecha:        24/10/2017
#   Departamento: Citrix 
#   Descripción:  Revisa los grupos que se le pasen por parametro compara y regulariza los usuarios en el grupo Destino
#
#
################################################## 



$Fecha = Get-Date

$fecha > "\\capelo\soft_wintel\Citrix\Scripts\Grupos\Log.txt"
$fecha > "\\capelo\soft_wintel\Citrix\Scripts\Grupos\LogErrores.txt"
$fecha > "\\capelo\soft_wintel\Citrix\Scripts\Grupos\LogFileValida.txt"

$Logfile = "\\capelo\soft_wintel\Citrix\Scripts\Grupos\Log.txt"
$LogfileError = "\\capelo\soft_wintel\Citrix\Scripts\Grupos\LogErrores.txt"
$LogfileValida = "\\capelo\soft_wintel\Citrix\Scripts\Grupos\LogFileValida.txt"


###################################################
#
#   Funcion Log Ejecucion

Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

###################################################
#
#   Funcion Log Errores 


Function LogWriteError
{
   Param ([string]$logstring)

   Add-content $LogfileError -value $logstring
}


###################################################
#
#   Funcion Log Comprobaciones 

Function LogWriteValida
{
   Param ([string]$logstring)

   Add-content $LogfileValida -value $logstring
}


###################################################
#
#   Funcion Regula Grupo GICONOS a GTS

Function RegulaGruposGTS
{
   Param 
   
   (
   [string]$ID,
   [string]$App,
   [string]$GTS,
   [string]$GICONOS
   

   )
   
###################################################   
#
#   Backup usuarios grupos

   $GTS_Users = (Get-ADGroupMember "$GTS").SamAccountName
   $GTS_Users > "\\capelo\soft_wintel\Citrix\Scripts\Grupos\$ID $GTS.txt"
   $GICONOS_Users = (Get-ADGroupMember "$GICONOS").SamAccountName 
   $GICONOS_Users > "\\capelo\soft_wintel\Citrix\Scripts\Grupos\$ID $GICONOS.txt"

   write-host "Aplicacion::::" $App
   write-host "Grupo GTS_Icono::::" $GTS
   write-host "Grupo GICONOS::::" $GICONOS
   write-host "Usuarios GTS_Icono::::" $GTS_Users
   write-host "Usuarios GICONOS::::" $GICONOS_Users
   

write-host "************ INICIO: $ID; APLICACION: $App; GRUPO GICONOS: $GICONOS; GRUPO GTS: $GTS ********************"
LogWrite   "************ INICIO: $ID; APLICACION: $App; GRUPO GICONOS: $GICONOS; GRUPO GTS: $GTS *********************"


  foreach ($user in $GICONOS_Users ) {

           
TRY {
           
   if ($GTS_Users -notcontains $user )

        {
        
        LogWrite "$user"
        Write-Host "$user no existe, lo añadimos"     
        Add-ADGroupMember $GTS –Member $user
        write-host "Modo Test no añade"
        #LogWrite "$user añadido al grupo $GTS"
        write-host "$user añadido al grupo $GTS"
        #LogWrite "Comprobación:"
        
        LogWriteValida ((GET-ADUSER –Identity $user –Properties MemberOf | Select-Object MemberOf).MemberOf -match  $GTS)
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


write-host "************ FIN: $ID; APLICACION: $App; GRUPO GICONOS: $GICONOS; GRUPO GTS: $GTS *********************"
write-host "********************************************************************************************************************"
write-host "********************************************************************************************************************"
LogWrite   "************ FIN: $ID; APLICACION: $App; GRUPO GICONOS: $GICONOS; GRUPO GTS: $GTS *********************"
LogWrite   "********************************************************************************************************************"
LogWrite   ""
LogWrite   ""
LogWrite   ""



}

######## FASE 1 ###########

######## GRANJA 6 ###########
<#
#$App3
RegulaGruposGTS  "App3" "BS Proteo 3" "GTS_Icono_Proteo3_SAU" "GICONOS SSCC TS Proteo3 SAU"
#$App4
RegulaGruposGTS  "App4" "Putty" "GTS_Icono_putty" "GICONOS SSCC TS PUTTY"
#$App5
RegulaGruposGTS  "App5" "Rochade BS" "GTS_Icono_Rochade" "GICONOS SSCC TS Rochade"
#$App6
RegulaGruposGTS  "App6" "Lotus Notes 8 - Basic" "GTS_Icono_LotusNotes8" "GICONOS SSCC TS NOTES_6"
#$App7
RegulaGruposGTS  "App7" "Acceso ProteoLife PreProduccion" "GTS_Icono_SibisPre_ProteoLife" "GICONOS SSCC TS SIBIS Preproduccion ProteoLife"
#$App8
RegulaGruposGTS  "App8" "Emulador 3270 PL" "GTS_Icono_Emulacio_ProteoLife" "GICONOS SSCC TS Emulador 3270 ProteoLife"
#$App9
RegulaGruposGTS  "App9" "Reconciliation" "GTS_Icono_Reconciliation" "GICONOS SSCC TS Reconciliation"
#$App10
RegulaGruposGTS  "App10" "Soporte Incidencias BS" "GTS_Icono_SSCC_Equipo_Virtual_CAM" "GICONOS SSCC Equipo Virtual CAM"
#$App11
RegulaGruposGTS  "App11" "SIBIS-Proteo Preproduccion" "GTS_Icono_Sibis_Proteo_Preproduccion" "GICONOS SSCC TS SIBIS Pre-Produccion"
#$App12
RegulaGruposGTS  "App12" "SAGE" "GTS_Icono_SAGE" "GICONOS SSCC TS SAGE"
#$App13
RegulaGruposGTS  "App13" "Siebel Web" "GTS_Icono_Siebel" "GICONOS SSCC TS Siebel"
#$App14
RegulaGruposGTS  "App14" "SEDAS Web" "GTS_Icono_SedasWeb" "GICONOS SSCC TS Sedas"
#$App15
RegulaGruposGTS  "App15" "SIBIS-Proteo Preproduccion Extern" "GTS_Icono_Sibis_PRE_Extern" "GICONOS SSCC TS SIBIS Pre-Produccion"
#$App16
RegulaGruposGTS  "App16" "Swift Alliance" "GTS_Icono_Swift_Alliance" "GICONOS SSCC TS Swift"
#$App17
RegulaGruposGTS  "App17" "UltraEdit" "GTS_Icono_Ultraedit" "GICONOS SSCC TS UltraEdit32"
#$App18
RegulaGruposGTS  "App18" "RochadePL" "GTS_Icono_Rochade_ProteoLife" "GICONOS SSCC TS Rochade Proteo Life"
#App18-1
RegulaGruposGTS  "App18-1" "RochadePL" "GTS_Icono_Rochade_ProteoLife" "GICONOS SSCC TS SIBIS Preproduccion ProteoLife"
#App18-2
RegulaGruposGTS  "App18-2" "RochadePL" "GTS_Icono_Rochade_ProteoLife" "GICONOS SSCC TS SIBIS Desarrollo ProteoLife"
#$App19
RegulaGruposGTS  "App19" "Centro de Servicio Integrado" "GTS_Icono_Centro_de_Servicio_Integrado" "GICONOS SSCC TS Centro de Servicio Integrado"
#$App20
RegulaGruposGTS  "App20" "RecAdmin" "GTS_Icono_RecAdmin" "GICONOS SSCC TS RecAdmin"
#$App21
RegulaGruposGTS  "App21" "PHabitat" "GTS_Icono_phabitat" "GICONOS SSCC TS phabitat"
#$App22
RegulaGruposGTS  "App22" "SIBIS-Proteo Desarrollo" "GTS_Icono_SIBIS_Proteo_D" "GICONOS SSCC TS SIBIS Desarrollo"
#$App23
RegulaGruposGTS  "App23" "Adaptiv Risk PRO" "GTS_Icono_Adaptiv_Risk_PRO" "GICONOS SSCC Adaptiv Risk PRO"

#>

#$App42
RegulaGruposGTS  "App42" "GOLF IIC Aceptacion" "GTS_Icono_GOLF_IIC_ACE" "GICONOS SSCC TS Golf IIC GOS"


RegulaGruposGTS  "App45-2" "CitrixPlus" "GTS_Icono_CitrixPlus" "GrpTSK_GrupoIconos_EUWM_L1" 

RegulaGruposGTS  "App45-3" "CitrixPlus" "GTS_Icono_CitrixPlus" "GrpTSK_GrupoIconos_EUWM_L2"

RegulaGruposGTS  "App28-5" "CitrixPlus" "GTS_Icono_CitrixPlus" "GrpMD_HP_EUWM_L1" 

RegulaGruposGTS  "App28-6" "CitrixPlus" "GTS_Icono_CitrixPlus" "GrpMD_HP_EUWM_L2"

######## FASE 2 ###########
<#


######### EXTRAS 1 ##################
RegulaGruposGTS  "App28-1" "Service Manager" "GTS_Icono_Service_Manager" "GrpMD_HP_WINTEL_L2" 

RegulaGruposGTS  "App28-2" "Service Manager" "GTS_Icono_Service_Manager" "GrpMD_HP_OPERACION_L1_ARG"

RegulaGruposGTS  "App28-3" "Service Manager" "GTS_Icono_Service_Manager" "GrpTSK_GrupoIconos_EUWM_L1" 

RegulaGruposGTS  "App28-4" "Service Manager" "GTS_Icono_Service_Manager" "GrpTSK_GrupoIconos_EUWM_L2"

RegulaGruposGTS  "App30-1" "Emulador 3270 BS" "GTS_Icono_PersonalComunication" "GICONOS SSCC TS Rochade"

RegulaGruposGTS  "App35" "Fronton Demonio IIC Exadata" "GTS_Icono_Fronton_demonio_IIC_Exadata" "GICONOS SSCC TS Fronton Demonio"

######### EXTRAS 2 ##################
RegulaGruposGTS  "App28-5" "Service Manager" "GTS_Icono_Service_Manager" "GrpMD_HP_EUWM_L1" 

RegulaGruposGTS  "App28-6" "Service Manager" "GTS_Icono_Service_Manager" "GrpMD_HP_EUWM_L2"

##################



#>


##################
<#

#$App24
RegulaGruposGTS  "App24" "DESARROLLO DE APLICACIONS EBRANCH" "GTS_Icono_ebranch_D" "GICONOS SSCC TS Ebranch" ok
#$App25
RegulaGruposGTS  "App25" "DESARROLLO DE APLICACIONS EBRANCH" "GTS_Icono_ebranch_D" "GICONOS EE TS Ebranch" ok
#$App26
RegulaGruposGTS  "App26" "SAP Logon" "GTS_Icono_SAP720" "GICONOS SSCC TS SAP" ok
#$App27
RegulaGruposGTS  "App27" "WorkFlow Accurate" "GTS_Icono_WorkFlow" "GICONOS SSCC TS Reconciliation" ok
#$App28
RegulaGruposGTS  "App28" "Service Manager" "GTS_Icono_Service_Manager" "GICONOS SSCC TS Service Manager" ok
#$App29
RegulaGruposGTS  "App29" "WebJetAdmin" "GTS_Icono_WebJetAdmin" "GICONOS SSCC TS WebJetAdmin" ok
#$App30
RegulaGruposGTS  "App30" "Emulador 3270 BS" "GTS_Icono_PersonalComunication" "GICONOS SSCC TS Emulador 3270" ok
#$App31 (Ya realizado Fase 1)
#RegulaGruposGTS  "App31" "SAGE" "GTS_Icono_SAGE" "GICONOS SSCC TS SAGE" ok
#$App32
RegulaGruposGTS  "App32" "CIDExadata" "GTS_Icono_Cid_exadata" "GICONOS SSCC TS CID" ok
#$App33
RegulaGruposGTS  "App33" "Fronton Exadata" "GTS_Icono_Fronton" "GICONOS SSCC TS Fronton" ok
#$App34
RegulaGruposGTS  "App34" "Fronton SBP Exadata" "GTS_Icono_Fronton_SBP" "GICONOS SSCC TS Fronton SBP" ok
#$App35
RegulaGruposGTS  "App35" "Fronton Demonio IIC Exadata" "GTS_Icono_Fronton_demonio_IIC_Exadata" "GICONOS SSCC TS Fronton Demonio" ok


#$App36
RegulaGruposGTS  "App36" "Golf IIC Exadata" "GTS_Icono_Golf_IIC" "GICONOS SSCC TS Golf IIC" ok
#$App37
RegulaGruposGTS  "App37" "Golf MV BS Exadata" "GTS_Icono_Golf_MV_Exadata" "GICONOS SSCC TS Golf MV"
#$App38
RegulaGruposGTS  "App38" "Explorar directorio local ebranch" "GTS_Icono_SSCC_TS_Ebranch" "GICONOS SSCC TS Ebranch" ok
#$App39
RegulaGruposGTS  "App39" "Explorar directorio local ebranch" "GTS_Icono_SSCC_TS_Ebranch" "GICONOS SSCC TS SIBIS Desarrollo" ok
#$App40
RegulaGruposGTS  "App40" "Explorar directorio local ebranch" "GTS_Icono_SSCC_TS_Ebranch" "GICONOS SSCC TS SIBIS Pre-Produccion" ok
#$App41
RegulaGruposGTS  "App41" "Fronton IIC Aceptacion" "GTS_Icono_Fronton_IIC_ACE" "GICONOS SSCC TS Fronton GOS" ok
#$App42
RegulaGruposGTS  "App42" "GOLF IIC Aceptacion" "GTS_Icono_GOLF_IIC_ACE" "GICONOS SSCC TS Golf IIC GOS" ok
#$App43
RegulaGruposGTS  "App43" "GOLF SBP ACEPTACION" "GTS_Icono_GOLF_SBP_ACE" "GICONOS SSCC TS Golf SBP GOS" ok
#$App44
RegulaGruposGTS  "App44" "SAGE_Treasury_P" "GTS_Icono_SAGEP" "GICONOS SSCC TS SAGE" ok


######## GRANJA 5 ###########

#$App45
RegulaGruposGTS  "App45" "CitrixPlus" "GTS_Icono_CitrixPlus" "GICONOS SSCC TS CitrixPlus" ok

#$App45-1
RegulaGruposGTS  "App45-1" "CitrixPlus" "GTS_Icono_CitrixPlus" "GICONOS SSCC TS 9800ceco" ok
#>