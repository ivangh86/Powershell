
##############################################
#
#
# Autor: Ivan Gomez
# Fecha: 14/03/2018
# Descripcion: Cambiar Configuracion Appv Citrix. tener en cuenta de desmarcar las opciones qeu queremos modificar y comentar el resto.
#
##############################################

if ((Get-PSSnapin "Citrix.*" -EA silentlycontinue) -eq $null) {
	try { Add-PSSnapin Citrix.* -ErrorAction Stop }
	catch { write-error "Error loading XenApp Powershell snapin"; Return }
}

#Variables Publisher y Manager

#PRO
####################################################################
#APPv FH
#$AppvPUblisher = "http://WCPFHIDAV001.tsb.banks.adgbs.com:8001"
#$AppvManager =  "http://WCPFHIDAM001.tsb.banks.adgbs.com"

#APPv NH
#$AppvPUblisher = "http://WCPNHIDAV001.tsb.banks.adgbs.com:8001"
#$AppvManager =  "http://WCPNHIDAM001.tsb.banks.adgbs.com"
####################################################################

#DEV
####################################################################
#APPv NH
$AppvPUblisher = "http://appvpdev.tsb.banks.adgbs.com:8001"
$AppvManager =  "http://appvmdev.tsb.banks.adgbs.com"

#APPv NH
#$AppvPUblisher = "http://WCDNHIDAV001.tsb.banks.adgbs.com:8001"
#$AppvManager =  "http://WCDNHIDAV001.tsb.banks.adgbs.com"
###################################################################

#Creamos nueva Politica de APPv
$policy = New-CtxAppVServer -ManagementServer $AppvManager -PublishingServer $AppvPUblisher

#Chequeamos el nuevo contenido de la Politica.
Get-CtxAppVServer -ByteArray $policy


#Cambiamos la configuracion de la Politica Existente:
Get-BrokerMachineConfiguration -Name appv\1 | Set-BrokerMachineConfiguration -Policy $policy 


#Herramientas de Revision

#$config = Get-BrokerMachineConfiguration -Name appv* 
#Get-BrokerMachineConfiguration -Name appv* 
#Get-CtxAppVServer -ByteArray $config[0].Policy

Write-Host "Testeando Publisher"
Test-CtxAppVServer -AppVPublishingServer $AppvPUblisher
Write-Host "Testeando Manager"
Test-CtxAppVServer -AppVManagementServer $AppvManager
