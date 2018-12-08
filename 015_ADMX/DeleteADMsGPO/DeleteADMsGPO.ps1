################################################################################################################################
#
# SCRIPT PARA ELIMINAR LA ADM EN CUESTION DE LA GPO QUE LE PASEMOS POR .TXT 
#
# Versiones: 
# 1.0 - 18/01/2017 - Oriol Agulló - Versión Inicial
#
################################################################################################################################

Import-Module grouppolicy

$GPOs = @(Get-Content "\\capelo\soft_wintel\Citrix\Temporal\Oriol_Agullo\Scripts\Proyecto adm a admx\DeleteADMsGPO\GPOs.txt")
$logDeleteADMsGPO = "\\capelo\soft_wintel\Citrix\Temporal\Oriol_Agullo\Scripts\Proyecto adm a admx\DeleteADMsGPO\logDeleteADMsGPO.txt"
$nameADM = (Get-Content "\\capelo\soft_wintel\Citrix\Temporal\Oriol_Agullo\Scripts\Proyecto adm a admx\DeleteADMsGPO\nameADM.txt")
$countErrors = 0

Write-Output "INICIO SCRIPT DeleteADMsGPO.PS1" | Out-File $logDeleteADMsGPO
Write-Output (Get-Date) | Out-File $logDeleteADMsGPO -Append
Write-Output "-------------------------------"`n | Out-File $logDeleteADMsGPO -Append

foreach ($GPO in $GPOs) {

    $objectGPO = Get-GPO $GPO
    $idGPO = $objectGPO.Id
    $pathADM = "\\adgbs.com\SYSVOL\adgbs.com\Policies\{$idGPO}\Adm\$nameADM.adm"
    
    if(Test-Path $pathADM) {

        try{

            Remove-Item -Path $pathADM -Force -WhatIf #DEJAR -WhatIf POR PRECAUCIÓN
            Write-Output "Archivo borrado: $pathADM" | Out-File $logDeleteADMsGPO -Append
    
        }catch{

            Write-Output "*** ERROR: Problemas al borrar el archivo: $pathADM" | Out-File $logDeleteADMsGPO -Append
            $countErrors += 1

        }

    }else{

        Write-Output "*** ERROR: NO existe el archivo: $pathADM" | Out-File $logDeleteADMsGPO -Append
        $countErrors += 1

    }
}

Write-Output "-------------------------------"`n | Out-File $logDeleteADMsGPO -Append
Write-Output "TOTAL ERRORES: $countErrors" | Out-File $logDeleteADMsGPO -Append
Write-Output (Get-Date) | Out-File $logDeleteADMsGPO -Append
Write-Output "FIN SCRIPT DeleteADMsGPO.PS1" | Out-File $logDeleteADMsGPO -Append