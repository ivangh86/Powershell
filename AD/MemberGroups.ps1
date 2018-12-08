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




#############################
#  
# Grupos y Usuarios a Migrar
#
#
#
######## GRANJA 6 ###########


#para pruebas son grupos de granja 5
#$Giconos_1 ="GICONOS SSCC TS Dameware"
#$Giconos_1_Usuarios = (Get-ADGroupMember "$Giconos_1").SamAccountName
#$GTS_1 ="GTS_Icono_Dameware"  
#$GTS_1_Usuarios = (Get-ADGroupMember "$GTS_1").SamAccountName


#$Giconos_2 ="GICONOS SSCC TS Rochade"
#$Giconos_2_Usuarios = (Get-ADGroupMember "$Giconos_2").SamAccountName
#$GTS_2 ="GTS_Icono_Rochade"
#$GTS_2_Usuarios = (Get-ADGroupMember "$GTS_2").SamAccountName  

#ok
$App23 = "Adaptiv Risk PRO"
$Giconos_23 = "GICONOS SSCC Adaptiv Risk PRO"
$Giconos_23_Usuarios = (Get-ADGroupMember "$Giconos_23").SamAccountName
$GTS_23 = "GTS_Icono_Adaptiv_Risk_PRO"
$GTS_23_Usuarios = (Get-ADGroupMember "$GTS_23").SamAccountName

<#

#ok
$App3 = "BS Proteo 3"
$Giconos_3 ="GICONOS SSCC TS Proteo3 SAU"
$Giconos_3_Usuarios = (Get-ADGroupMember "$Giconos_3").SamAccountName
$GTS_3 =          "GTS_Icono_Proteo3_SAU"
$GTS_3_Usuarios = (Get-ADGroupMember "$GTS_3").SamAccountName

#ok
$App4 = "Putty"
$Giconos_4 = "GICONOS SSCC TS PUTTY"
$Giconos_4_Usuarios = (Get-ADGroupMember "$Giconos_4").SamAccountName
$GTS_4 = "GTS_Icono_putty"
$GTS_4_Usuarios = (Get-ADGroupMember "$GTS_4").SamAccountName

#ok
$App5 = "Rochade BS"
$Giconos_5 = "GICONOS SSCC TS Rochade"
$Giconos_5_Usuarios = (Get-ADGroupMember "$Giconos_5").SamAccountName
$GTS_5 = "GTS_Icono_Rochade"
$GTS_5_Usuarios = (Get-ADGroupMember "$GTS_5").SamAccountName

#ok
$App6 = "Lotus Notes 8 - Basic"
$Giconos_6 = "GICONOS SSCC TS NOTES_6"
$Giconos_6_Usuarios = (Get-ADGroupMember "$Giconos_6").SamAccountName
$GTS_6 = "GTS_Icono_LotusNotes8"
$GTS_6_Usuarios = (Get-ADGroupMember "$GTS_6").SamAccountName

#ok
$App7 = "Acceso ProteoLife PreProduccion"
$Giconos_7 = "GICONOS SSCC TS SIBIS Preproduccion ProteoLife"
$Giconos_7_Usuarios = (Get-ADGroupMember "$Giconos_7").SamAccountName
$GTS_7 = "GTS_Icono_SibisPre_ProteoLife"
$GTS_7_Usuarios = (Get-ADGroupMember "$GTS_7").SamAccountName

#ok
$App8 = "Emulador 3270 PL"
$Giconos_8 = "GICONOS SSCC TS Emulador 3270 ProteoLife"
$Giconos_8_Usuarios = (Get-ADGroupMember "$Giconos_8").SamAccountName
$GTS_8 = "GTS_Icono_Emulacio_ProteoLife"
$GTS_8_Usuarios = (Get-ADGroupMember "$GTS_8").SamAccountName

#ok
$App9 = "Reconciliation"
$Giconos_9 = "GICONOS SSCC TS Reconciliation"
$Giconos_9_Usuarios = (Get-ADGroupMember "$Giconos_9").SamAccountName
$GTS_9 = "GTS_Icono_Reconciliation"
$GTS_9_Usuarios = (Get-ADGroupMember "$GTS_9").SamAccountName

#ok
$App10 = "Soporte Incidencias BS"
$Giconos_10 = "GICONOS SSCC Equipo Virtual CAM"
$Giconos_10_Usuarios = (Get-ADGroupMember "$Giconos_10").SamAccountName
$GTS_10 = "GTS_Icono_SSCC_Equipo_Virtual_CAM"
$GTS_10_Usuarios = (Get-ADGroupMember "$GTS_10").SamAccountName

#ok
$App11 = "SIBIS-Proteo Preproduccion"
$Giconos_11 = "GICONOS SSCC TS SIBIS Pre-Produccion"
$Giconos_11_Usuarios = (Get-ADGroupMember "$Giconos_11").SamAccountName
$GTS_11 = "GTS_Icono_Sibis_Proteo_Preproduccion"
$GTS_11_Usuarios = (Get-ADGroupMember "$GTS_11").SamAccountName


#ok
$App12 = "SAGE"
$Giconos_12 = "GICONOS SSCC TS SAGE"
$Giconos_12_Usuarios = (Get-ADGroupMember "$Giconos_12").SamAccountName
$GTS_12 = "GTS_Icono_SAGE"
$GTS_12_Usuarios = (Get-ADGroupMember "$GTS_12").SamAccountName


#ok
$App13 = "Siebel Web"
$Giconos_13 = "GICONOS SSCC TS Siebel"
$Giconos_13_Usuarios = (Get-ADGroupMember "$Giconos_13").SamAccountName
$GTS_13 = "GTS_Icono_Siebel"
$GTS_13_Usuarios = (Get-ADGroupMember "$GTS_13").SamAccountName



#ok
$App14 = "SEDAS Web"
$Giconos_14 = "GICONOS SSCC TS Sedas"
$Giconos_14_Usuarios = (Get-ADGroupMember "$Giconos_14").SamAccountName
$GTS_14 = "GTS_Icono_SedasWeb"
$GTS_14_Usuarios = (Get-ADGroupMember "$GTS_14").SamAccountName


#ok
$App15 = "SIBIS-Proteo Preproduccion Extern"
$Giconos_15 = "GICONOS SSCC TS SIBIS Pre-Produccion"
$Giconos_15_Usuarios = (Get-ADGroupMember "$Giconos_15").SamAccountName
$GTS_15 = "GTS_Icono_Sibis_PRE_Extern"
$GTS_15_Usuarios = (Get-ADGroupMember "$GTS_15").SamAccountName


#ok
$App16 = "Swift Alliance"
$Giconos_16 = "GICONOS SSCC TS Swift"
$Giconos_16_Usuarios = (Get-ADGroupMember "$Giconos_16").SamAccountName
$GTS_16 = "GTS_Icono_Swift_Alliance"
$GTS_16_Usuarios = (Get-ADGroupMember "$GTS_16").SamAccountName


#ok
$App17 = "UltraEdit"
$Giconos_17 = "GICONOS SSCC TS UltraEdit32"
$Giconos_17_Usuarios = (Get-ADGroupMember "$Giconos_17").SamAccountName
$GTS_17 = "GTS_Icono_Ultraedit"
$GTS_17_Usuarios = (Get-ADGroupMember "$GTS_17").SamAccountName


#ok
$App18 = "RochadePL"
$Giconos_18 = "GICONOS SSCC TS Rochade Proteo Life"
$Giconos_18_Usuarios = (Get-ADGroupMember "$Giconos_18").SamAccountName
$GTS_18 = "GTS_Icono_Rochade_ProteoLife"
$GTS_18_Usuarios = (Get-ADGroupMember "$GTS_18").SamAccountName


#ok
$App19 = "Centro de Servicio Integrado"
$Giconos_19 = "GICONOS SSCC TS Centro de Servicio Integrado"
$Giconos_19_Usuarios = (Get-ADGroupMember "$Giconos_19").SamAccountName
$GTS_19 = "GTS_Icono_Centro_de_Servicio_Integrado"
$GTS_19_Usuarios = (Get-ADGroupMember "$GTS_19").SamAccountName


#ok
$App20 = "RecAdmin"
$Giconos_20 = "GICONOS SSCC TS RecAdmin"
$Giconos_20_Usuarios = (Get-ADGroupMember "$Giconos_20").SamAccountName
$GTS_20 = "GTS_Icono_RecAdmin"
$GTS_20_Usuarios = (Get-ADGroupMember "$GTS_20").SamAccountName


#ok
$App21 = "PHabitat"
$Giconos_21 = "GICONOS SSCC TS phabitat"
$Giconos_21_Usuarios = (Get-ADGroupMember "$Giconos_21").SamAccountName
$GTS_21 = "GTS_Icono_phabitat"
$GTS_21_Usuarios = (Get-ADGroupMember "$GTS_21").SamAccountName


#ok
$App22 = "SIBIS-Proteo Desarrollo"
$Giconos_22 = "GICONOS SSCC TS SIBIS Desarrollo"
$Giconos_22_Usuarios = (Get-ADGroupMember "$Giconos_22").SamAccountName
$GTS_22 = "GTS_Icono_SIBIS_Proteo_D"
$GTS_22_Usuarios = (Get-ADGroupMember "$GTS_22").SamAccountName


#ok
$App23 = "Adaptiv Risk PRO"
$Giconos_23 = "GICONOS SSCC Adaptiv Risk PRO"
$Giconos_23_Usuarios = (Get-ADGroupMember "$Giconos_23").SamAccountName
$GTS_23 = "GTS_Icono_Adaptiv_Risk_PRO"
$GTS_23_Usuarios = (Get-ADGroupMember "$GTS_23").SamAccountName


#ok
$App24 = "DESARROLLO DE APLICACIONS EBRANCH"
$Giconos_24 = "GICONOS SSCC TS Ebranch"
$Giconos_24_Usuarios = (Get-ADGroupMember "$Giconos_24").SamAccountName
$GTS_24 = "GTS_Icono_ebranch_D"
$GTS_24_Usuarios = (Get-ADGroupMember "$GTS_24").SamAccountName


#ok
$App25 = "DESARROLLO DE APLICACIONS EBRANCH"
$Giconos_25 = "GICONOS EE TS Ebranch"
$Giconos_25_Usuarios = (Get-ADGroupMember "$Giconos_25").SamAccountName
$GTS_25 = "GTS_Icono_ebranch_D"
$GTS_25_Usuarios = (Get-ADGroupMember "$GTS_25").SamAccountName


#ok
$App26 = "SAP Logon"
$Giconos_26 = "GICONOS SSCC TS SAP"
$Giconos_26_Usuarios = (Get-ADGroupMember "$Giconos_26").SamAccountName
$GTS_26 = "GTS_Icono_SAP720"
$GTS_26_Usuarios = (Get-ADGroupMember "$GTS_26").SamAccountName


#ok
$App27 = "WorkFlow Accurate"
$Giconos_27 = "GICONOS SSCC TS Reconciliation"
$Giconos_27_Usuarios = (Get-ADGroupMember "$Giconos_27").SamAccountName
$GTS_27 = "GTS_Icono_WorkFlow"
$GTS_27_Usuarios = (Get-ADGroupMember "$GTS_27").SamAccountName


#ok
$App28 = "Service Manager"
$Giconos_28 = "GICONOS SSCC TS Service Manager"
$Giconos_28_Usuarios = (Get-ADGroupMember "$Giconos_28").SamAccountName
$GTS_28 = "GTS_Icono_Service_Manager"
$GTS_28_Usuarios = (Get-ADGroupMember "$GTS_28").SamAccountName


#ok
$App29 = "WebJetAdmin"
$Giconos_29 = "GICONOS SSCC TS WebJetAdmin"
$Giconos_29_Usuarios = (Get-ADGroupMember "$Giconos_29").SamAccountName
$GTS_29 = "GTS_Icono_WebJetAdmin"
$GTS_29_Usuarios = (Get-ADGroupMember "$GTS_29").SamAccountName


#ok
$App30 = "Emulador 3270 BS"
$Giconos_30 = "GICONOS SSCC TS Emulador 3270"
$Giconos_30_Usuarios = (Get-ADGroupMember "$Giconos_30").SamAccountName
$GTS_30 = "GTS_Icono_PersonalComunication"
$GTS_30_Usuarios = (Get-ADGroupMember "$GTS_30").SamAccountName


#ok
$App31 = "SAGE"
$Giconos_31 = "GICONOS SSCC TS SAGE"
$Giconos_31_Usuarios = (Get-ADGroupMember "$Giconos_31").SamAccountName
$GTS_31 = "GTS_Icono_SAGE"
$GTS_31_Usuarios = (Get-ADGroupMember "$GTS_31").SamAccountName


#ok
$App32 = "CIDExadata"
$Giconos_32 = "GICONOS SSCC TS CID"
$Giconos_32_Usuarios = (Get-ADGroupMember "$Giconos_32").SamAccountName
$GTS_32 = "GTS_Icono_Cid_exadata"
$GTS_32_Usuarios = (Get-ADGroupMember "$GTS_32").SamAccountName


#ok
$App33 = "Fronton Exadata"
$Giconos_33 =  "GICONOS SSCC TS Fronton"
$Giconos_33_Usuarios = (Get-ADGroupMember "$Giconos_33").SamAccountName
$GTS_33 = "GTS_Icono_Fronton"
$GTS_33_Usuarios = (Get-ADGroupMember "$GTS_33").SamAccountName


#ok
$App34 = "Fronton SBP Exadata"
$Giconos_34 = "GICONOS SSCC TS Fronton SBP"
$Giconos_34_Usuarios = (Get-ADGroupMember "$Giconos_34").SamAccountName
$GTS_34 ="GTS_Icono_Fronton_SBP"
$GTS_34_Usuarios = (Get-ADGroupMember "$GTS_34").SamAccountName


#ok
$App35 = "Fronton Demonio IIC Exadata"
$Giconos_35 = "GICONOS SSCC TS Fronton Demonio"
$Giconos_35_Usuarios = (Get-ADGroupMember "$Giconos_35").SamAccountName
$GTS_35 = "GTS_Icono_Fronton_demonio_IIC_Exada"
$GTS_35_Usuarios = (Get-ADGroupMember "$GTS_35").SamAccountName


#ok
$App36 = "Golf IIC Exadata"
$Giconos_36 = "GICONOS SSCC TS Golf IIC"
$Giconos_36_Usuarios = (Get-ADGroupMember "$Giconos_36").SamAccountName
$GTS_36 = "GTS_Icono_Golf_IIC"
$GTS_36_Usuarios = (Get-ADGroupMember "$GTS_36").SamAccountName


#ok
$App37 = "Golf MV BS Exadata"
$Giconos_37 = "GICONOS SSCC TS Golf MV"
$Giconos_37_Usuarios = (Get-ADGroupMember "$Giconos_37").SamAccountName
$GTS_37 = "GTS_Icono_Golf_MV_Exadata"
$GTS_37_Usuarios = (Get-ADGroupMember "$GTS_37").SamAccountName


#ok
$App38 = "Explorar directorio local ebranch"
$Giconos_38 = "GICONOS SSCC TS Ebranch"
$Giconos_38_Usuarios = (Get-ADGroupMember "$Giconos_38").SamAccountName
$GTS_38 = "GTS_Icono_SSCC_TS_Ebranch"
$GTS_38_Usuarios = (Get-ADGroupMember "$GTS_38").SamAccountName


#ok
$App39 = "Explorar directorio local ebranch"
$Giconos_39 = "GICONOS SSCC TS SIBIS Desarrollo"
$Giconos_39_Usuarios = (Get-ADGroupMember "$Giconos_39").SamAccountName
$GTS_39 = "GTS_Icono_SSCC_TS_Ebranch"
$GTS_39_Usuarios = (Get-ADGroupMember "$GTS_39").SamAccountName


#ok
$App40 = "Explorar directorio local ebranch"
$Giconos_40 = "GICONOS SSCC TS SIBIS Pre-Produccion"
$Giconos_40_Usuarios = (Get-ADGroupMember "$Giconos_40").SamAccountName
$GTS_40 = "GTS_Icono_SSCC_TS_Ebranch"
$GTS_40_Usuarios = (Get-ADGroupMember "$GTS_40").SamAccountName


#ok
$App41 = "Fronton IIC Aceptacion"
$Giconos_41 = "GICONOS SSCC TS Fronton GOS"
$Giconos_41_Usuarios = (Get-ADGroupMember "$Giconos_41").SamAccountName
$GTS_41 = "GTS_Icono_Fronton_IIC_ACE"
$GTS_41_Usuarios = (Get-ADGroupMember "$GTS_41").SamAccountName


#ok
$App42 = "GOLF IIC Aceptacion"
$Giconos_42 = "GICONOS SSCC TS Golf IIC GOS"
$Giconos_42_Usuarios = (Get-ADGroupMember "$Giconos_42").SamAccountName
$GTS_42 = "GTS_Icono_GOLF_IIC_ACE"
$GTS_42_Usuarios = (Get-ADGroupMember "$GTS_42").SamAccountName


#ok
$App43 = "GOLF SBP ACEPTACION"
$Giconos_43 = "GICONOS SSCC TS Golf SBP GOS"
$Giconos_43_Usuarios = (Get-ADGroupMember "$Giconos_43").SamAccountName
$GTS_43 = "GTS_Icono_GOLF_SBP_ACE"
$GTS_43_Usuarios = (Get-ADGroupMember "$GTS_43").SamAccountName


#ok
$App44 = "SAGE_Treasury_PL"
$Giconos_44 = "GICONOS SSCC TS SAGE"
$Giconos_44_Usuarios = (Get-ADGroupMember "$Giconos_44").SamAccountName
$GTS_44 = "GTS_Icono_SAGEP"
$GTS_44_Usuarios = (Get-ADGroupMember "$GTS_44").SamAccountName

######## GRANJA 5 ###########

#ok
$App45 = "CitrixPlus"
$Giconos_45 = "GICONOS SSCC TS CitrixPlus"
$Giconos_45_Usuarios = (Get-ADGroupMember "$Giconos_45").SamAccountName
$GTS_45 = "GTS_Icono_CitrixPlus"
$GTS_45_Usuarios = (Get-ADGroupMember "$GTS_45").SamAccountName


#>

################################

#$GTS_Grupos = ("$GTS_1","$GTS_2","$GTS_3","$GTS_4","$GTS_5","$GTS_6","$GTS_7","$GTS_8","$GTS_9","$GTS_10","$GTS_11","$GTS_12","$GTS_13","$GTS_14","$GTS_15","$GTS_16","$GTS_17",
#"$GTS_18","$GTS_19","$GTS_20","$GTS_21","$GTS_22","$GTS_23","$GTS_24","$GTS_25","$GTS_26","$GTS_27","$GTS_28","$GTS_29","$GTS_30","$GTS_31","$GTS_32","$GTS_33","$GTS_34","$GTS_35","$GTS_36",
#"$GTS_37","$GTS_38","$GTS_39","$GTS_40","$GTS_41","$GTS_42","$GTS_43","$GTS_44","$GTS_45")

#$Giconos_Grupos = ("$Giconos_1","$Giconos_2","$Giconos_3","$Giconos_4","$Giconos_5","$Giconos_6","$Giconos_7","$Giconos_8","$Giconos_9","$Giconos_10","$Giconos_11","$Giconos_12","$Giconos_13","$Giconos_14","$Giconos_15","$Giconos_16","$Giconos_17",
#"$Giconos_18","$Giconos_19","$Giconos_20","$Giconos_21","$Giconos_22","$Giconos_23","$Giconos_24","$Giconos_25","$Giconos_26","$Giconos_27","$Giconos_28","$Giconos_29","$Giconos_30","$Giconos_31","$Giconos_32","$Giconos_33","$Giconos_34","$Giconos_35","$Giconos_36",
#"$Giconos_37","$Giconos_38","$Giconos_39","$Giconos_40","$Giconos_41","$Giconos_42","$Giconos_43","$Giconos_44","$Giconos_45")

################################


#  
LogWrite  "Hacemos backup de los GTS_Icono"
# 
LogWrite "C:\temp\grupos\$GTS_23.txt"
#$GTS_1_Usuarios > "C:\temp\grupos\$GTS_1.txt"
#LogWrite $GTS_2 > "C:\temp\grupos\$GTS_2.txt"
#LogWrite $GTS_3 > "C:\temp\grupos\$GTS_3.txt"
#LogWrite $GTS_4 > "C:\temp\grupos\$GTS_4.txt"
#LogWrite $GTS_5 > "C:\temp\grupos\$GTS_5.txt"
#LogWrite $GTS_6 > "C:\temp\grupos\$GTS_6.txt"
#LogWrite $GTS_7 > "C:\temp\grupos\$GTS_7.txt"
#LogWrite $GTS_8 > "C:\temp\grupos\$GTS_8.txt"
#LogWrite $GTS_9 > "C:\temp\grupos\$GTS_9.txt"
#LogWrite $GTS_10 > "C:\temp\grupos\$GTS_10.txt"
#LogWrite $GTS_11 > "C:\temp\grupos\$GTS_11.txt"
#LogWrite $GTS_12 > "C:\temp\grupos\$GTS_12.txt"
#LogWrite $GTS_13 > "C:\temp\grupos\$GTS_13.txt"
#LogWrite $GTS_14 > "C:\temp\grupos\$GTS_14.txt"
#LogWrite $GTS_15 > "C:\temp\grupos\$GTS_15.txt"
#LogWrite $GTS_16 > "C:\temp\grupos\$GTS_16.txt"
#LogWrite $GTS_17 > "C:\temp\grupos\$GTS_17.txt"
#LogWrite $GTS_18 > "C:\temp\grupos\$GTS_18.txt"
#LogWrite $GTS_19 > "C:\temp\grupos\$GTS_19.txt"
#LogWrite $GTS_20 > "C:\temp\grupos\$GTS_20.txt"
#LogWrite $GTS_21 > "C:\temp\grupos\$GTS_21.txt"
#LogWrite $GTS_22 > "C:\temp\grupos\$GTS_22.txt"
$GTS_23_Usuarios > "C:\temp\grupos\$GTS_23.txt"
#LogWrite $GTS_24 > "C:\temp\grupos\$GTS_24.txt"
#LogWrite $GTS_25 > "C:\temp\grupos\$GTS_25.txt"
#LogWrite $GTS_26 > "C:\temp\grupos\$GTS_26.txt"
#LogWrite $GTS_27 > "C:\temp\grupos\$GTS_27.txt"
#LogWrite $GTS_28 > "C:\temp\grupos\$GTS_28.txt"
#LogWrite $GTS_29 > "C:\temp\grupos\$GTS_29.txt"
#LogWrite $GTS_30 > "C:\temp\grupos\$GTS_30.txt"
#LogWrite $GTS_31 > "C:\temp\grupos\$GTS_31.txt"
#LogWrite $GTS_32 > "C:\temp\grupos\$GTS_32.txt"
#LogWrite $GTS_33 > "C:\temp\grupos\$GTS_33.txt"
#LogWrite $GTS_34 > "C:\temp\grupos\$GTS_34.txt"
#LogWrite $GTS_35 > "C:\temp\grupos\$GTS_35.txt"
#LogWrite $GTS_36 > "C:\temp\grupos\$GTS_36.txt"
#LogWrite $GTS_37 > "C:\temp\grupos\$GTS_37.txt"
#LogWrite $GTS_38 > "C:\temp\grupos\$GTS_38.txt"
#LogWrite $GTS_39 > "C:\temp\grupos\$GTS_39.txt"
#LogWrite $GTS_40 > "C:\temp\grupos\$GTS_40.txt"
#LogWrite $GTS_41 > "C:\temp\grupos\$GTS_41.txt"
#LogWrite $GTS_42 > "C:\temp\grupos\$GTS_42.txt"
#LogWrite $GTS_43 > "C:\temp\grupos\$GTS_43.txt"
#LogWrite $GTS_44 > "C:\temp\grupos\$GTS_44.txt"
#LogWrite $GTS_45 > "C:\temp\grupos\$GTS_45.txt"


#
#Comenzamos con la revision de los grupos
#
write-host "************ INICIO $GTS_23 Aplicación $App23 *********************"
LogWrite   "************ INICIO $GTS_23 Aplicación $App23 *********************"



Function RegulaGruposGTS
{
   Param 
   
   (
   [string]$GTS,
   [string]$GICONOS,
   [string]$GTS_Users,
   [string]$GICONOS_Users,

   )


#ok
$App23 = "Adaptiv Risk PRO"
$Giconos_23 = "GICONOS SSCC Adaptiv Risk PRO"
$Giconos_23_Usuarios = (Get-ADGroupMember "$Giconos_23").SamAccountName
$GTS_23 = "GTS_Icono_Adaptiv_Risk_PRO"
$GTS_23_Usuarios = (Get-ADGroupMember "$GTS_23").SamAccountName


foreach ($user in $Giconos_23_Usuarios ) {


TRY{

    if ($GTS_23_Usuarios -notcontains $user )

        {
        LogWrite "$user no existe, lo añadimos"
        Write-Host "$user no existe, lo añadimos"
        
        #Add-ADGroupMember $GTS_1 –Member $user
        write-host "Modo Test no añade"
        LogWrite "$user añadido al grupo $GTS_23"
        write-host "$user añadido al grupo $GTS_23"

        LogWrite "Comprobación:"
        
        #LogWrite ((GET-ADUSER –Identity $user –Properties MemberOf | Select-Object MemberOf).MemberOf -match  $Giconos_Rochade)
        LogWrite "Modo Test no Comprueba"
        
        }

    else {  
    
    LogWrite "$user Ya existe"
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

write-host "************FIN $GTS_23 Aplicación $App23 *********************"
LogWrite   "************FIN $GTS_23 Aplicación $App23 *********************"


}