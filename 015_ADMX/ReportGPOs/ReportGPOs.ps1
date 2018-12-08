################################################################################################################################
#
# SCRIPT PARA SACAR EL REPORT DE LAS GPOs QUE SE PASEN POR .TXT Y GUARDAR LOS REPORTES EN .XML EN UNA CARPETA 
#
# Versiones: 
# 1.0 - 19/01/2017 - Oriol Agulló - Versión Inicial
# 1.1 - 11/05/2017 - Oriol Agulló - Pasar todos los path absolutos a relativos
#
################################################################################################################################
Import-Module grouppolicy

$GPOs = @(Get-Content ".\GPOs.txt")
$logReportGPOs = ".\logReportGPOs.txt"
$countErrors = 0
$countExistReport = 0
$countCreateReports = 0
$folderDay = (get-date -Format dd_MM_yyyy)
$folderReports = ".\Reports\$folderDay"
Write-Output "INICIO SCRIPT ReportGPOs.PS1" | Out-File $logReportGPOs
Write-Output (Get-Date) | Out-File $logReportGPOs -Append
Write-Output "-------------------------------"`n | Out-File $logReportGPOs -Append

#Creamos la carpeta $folderDay dentro de .\Reports\ si NO existe.
if (-not(Test-Path "$folderReports")) {
    
    try{

    New-Item -Path $folderReports -ItemType "directory" | Out-Null
    Write-Output "Creación de directorio: $folderReports" | Out-File $logReportGPOs -Append

    }catch{

        Write-Output "***ERROR: NO se ha podido crear $folderReports" | Out-File $logReportGPOs -Append
        $countErrors += 1

    }

}else{

    Write-Output "Ya está creado el directorio: $folderReports" | Out-File $logReportGPOs -Append

}

foreach ($GPO in $GPOs) {

    try{
        
        if(-not(Test-Path "$folderReports\$GPO.html")){

            Get-GPOReport -Name "$GPO" -ReportType Html | Out-File -FilePath "$folderReports\$GPO.html"
            Write-Output "Report creado de la GPO: $GPO" | Out-File $logReportGPOs -Append
            $countCreateReports += 1

        }else{

            Write-Output "Ya existe el report de la GPO: $GPO" | Out-File $logReportGPOs -Append
            $countExistReport += 1

        }

    }catch{

        Write-Output "***ERROR: NO se ha podido crear el report de la GPO: $GPO" | Out-File $logReportGPOs -Append
        $countErrors += 1

    }

}

Write-Output "-------------------------------"`n | Out-File $logReportGPOs -Append
Write-Output "TOTAL REPORTS CREADOS: $countCreateReports" | Out-File $logReportGPOs -Append
Write-Output "TOTAL REPORTS QUE YA EXISTEN: $countExistReport" | Out-File $logReportGPOs -Append
Write-Output "TOTAL ERRORES: $countErrors" | Out-File $logReportGPOs -Append
Write-Output (Get-Date) | Out-File $logReportGPOs -Append
Write-Output "FIN SCRIPT DeleteADMsGPO.PS1" | Out-File $logReportGPOs -Append