################################################################################################################################
#
# SCRIPT PARA COMPARAR GPOs QUE SE ENCUENTRAN EN 2 .TXT 
#
# Versiones: 
# 1.0 - 20/01/2017 - Oriol Agulló - Versión Inicial
#
################################################################################################################################

$pathFolderFirstGPOs = Get-Content "\\capelo\soft_wintel\Citrix\Temporal\Oriol_Agullo\Scripts\Proyecto adm a admx\CompareGPOs\pathFolderFirstGPOs.txt"
$pathFolderSecondGPOs = Get-Content "\\capelo\soft_wintel\Citrix\Temporal\Oriol_Agullo\Scripts\Proyecto adm a admx\CompareGPOs\pathFolderSecondGPOs.txt"
$logCompareGPOs = "\\capelo\soft_wintel\Citrix\Temporal\Oriol_Agullo\Scripts\Proyecto adm a admx\CompareGPOs\logCompareGPOs.txt"
$countErrors = 0
$countGPOsDifferences = 0
$countGPOsOK = 0

Write-Output "INICIO SCRIPT CompareGPOs.PS1" | Out-File $logCompareGPOs
Write-Output (Get-Date) | Out-File $logCompareGPOs -Append
Write-Output "-------------------------------"`n | Out-File $logCompareGPOs -Append

#Obtener todas las GPOs de la ruta $pathFolderFirstGPOs
$firstGPOs = Get-ChildItem $pathFolderFirstGPOs -File

#comparamos las GPOs con el mismo nombre en las diferentes carpetas
foreach($firstGPO in $FirstGPOs){
        
        try{

            $nameGPO = $firstGPO.name
            $firstGPO = Get-Content "$pathFolderFirstGPOs\$nameGPO"
            $secondGPO = Get-Content "$pathFolderSecondGPOs\$nameGPO"
            $resultCompare = Compare-Object -ReferenceObject $firstGPO -DifferenceObject $secondGPO 
            
        if($resultCompare -eq $null){

            Write-Output "$nameGPO no hay diferencias - OK" | Out-File $logCompareGPOs -Append
            $countGPOsOK += 1

        }else{

            Write-Output "*** $nameGPO SI que tiene diferencias - KO" | Out-File $logCompareGPOs -Append
            #Write-Output $resultCompare | Out-File $logCompareGPOs -Append
            $countGPOsDifferences += 1

        }

        }catch{

            Write-Output "***ERROR: Algún error al comparar la GPO: $nameGPO" | Out-File $logCompareGPOs -Append
            $countErrors += 1

        }
}

Write-Output "-------------------------------"`n | Out-File $logCompareGPOs -Append
Write-Output "TOTAL GPOs OK: $countGPOsOK" | Out-File $logCompareGPOs -Append
Write-Output "TOTAL GPOs CON DIFERENCIAS: $countGPOsDifferences" | Out-File $logCompareGPOs -Append
Write-Output "TOTAL ERRORES: $countErrors" | Out-File $logCompareGPOs -Append
Write-Output (Get-Date) | Out-File $logCompareGPOs -Append
Write-Output "FIN SCRIPT CompareGPOs.PS1" | Out-File $logCompareGPOs -Append