################################################################################################################################
#
# SCRIPT PARA ENCONTRAR EN QUE DIRECTORIOS DEL SYSVOL HAY LA .ADM EN CUESTIÓN 
#
# Versiones: 
# 1.0 - 17/01/2017 - Oriol Agulló - Versión Inicial
#
################################################################################################################################

$nameADM = Get-Content .\nameADM.txt

$filePath = ".\$nameADM.txt"

$pathsADM = (Get-ChildItem -Path "\\adgbs.com\SYSVOL\adgbs.com\Policies" -File -Recurse | Where-Object name -eq $nameADM | select DirectoryName)

foreach ($pathADM in $pathsADM){

    $pathADM.DirectoryName | Out-File $filePath -Append

}