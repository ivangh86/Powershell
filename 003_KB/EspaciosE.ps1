############################################################################### 
#        
#        
#        Name:          EspaciosE.ps1
#        Author:        Iván Gómez 
#        Date:          25/04/2017 
#        Description:   Revisamos la distribución del espacio en launidad E:\ para los Equipos XAV. 
#                           
#
############################################################################### 
###############################################################################

#Modulos Necesarios
#


Add-PSSnapin citrix.*
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

#Iniciando Html


$a = ""
$a = ConvertTo-Html -as Table  -Title "Desarrollo - Informe Espacio Unidad E:\ Equipos XAV" -PreContent "<p><h2>Desarrollo - Informe Espacio E: Equipos XAV </h2></p><br>" -PostContent "<br>Informe Realizado $(Get-Date)"  > C:\Prueba\EspacioE.html




$Servers = Get-XAServer | Where-Object { $_.ServerName -like "XAV*"} 
 
foreach ($Server in $Servers)
 
{
$ServerName = $Server.ServerName
#Write-Host $ServerName

$dirIP = (Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ServerName | ? {$_.IPAddress -ne $null}).ipaddress


################# Obtengo tamaño pagefile.sys y vdiskcache ############################## 

#Get-ChildItem -File -Hidden "E:" | Where-Object  {$_.name -Like "*.vdiskcache" -or "pagefile.sys"}

#Write-host Distribución Espacio Unidad E en GB
#Write-host
$vdiskcache = Get-ChildItem -File -Hidden "\\$ServerName\E$" | Where-Object  {$_.name -Like "*.vdiskcache"} 
$vdiskcachet = $vdiskcache.Length / 1GB
$vdiskcachetf = [math]::Round($vdiskcachet,2)



#write-host Espacio Ocupado .vdiskcache $vdiskcachet
$pagefile = Get-ChildItem -File -Hidden "\\$ServerName\E$" | Where-Object  {$_.name -Like "pagefile.sys"}
$pageFilet = $pagefile.Length / 1GB
$pageFiletf = [math]::Round($pageFilet,2)
#write-host Espacio Ocupado PageFile.sys $pageFilet


################# Obtengo tamaño de la Unidad E ############################## 

$disk = Get-WmiObject -class win32_logicaldisk -filter "DeviceID='E:'" -ComputerName $ServerName
#$freepercent = $disk.freespace / $disk.size * 100
##Write-Host "Free space on drive C: at $freepercent percent"
$TotalSize = ($disk.Size / 1GB)
$TotalSizef = [math]::Round($FreeSpace,2)
#write-host Espacio Total E $TotalSize 
$FreeSpace = ($disk.FreeSpace / 1GB)
$FreeSpacef = [math]::Round($FreeSpace,2) 
#write-host Espacio Libre E $FreeSpace
$UsageSpace = $TotalSize - $FreeSpace
$UsageSpacef = [math]::Round($UsageSpace,2)
#write-host Espacio Usado E $UsageSpace

$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@


$Table = @"
Servidor,IP,Espacio Ocupado vdiskcache (GB),Espacio Ocupado PageFile.sys (GB),Espacio Total E: (GB),Espacio Libre E: (GB),Espacio Usado E: (GB)

$ServerName,$dirIP,$vdiskcachetf,$pageFiletf,$TotalSizef,$FreeSpacef,$UsageSpacef
"@

$Table = $Table | ConvertFrom-Csv
$HTML = $Table | ConvertTo-Html -as Table  -Title "Informe Unidad E:" -head $Header <#-PreContent "<p><h3>$Server</h3></p><br>" -PostContent "<br>Informe Realizado $(Get-Date)"#>  >> $scriptDir\EspacioE.html

#$hash

#ConvertTo-Html -Body "$ServerName"  -CssUri c:\style.css >> \\XAPSDKBPV01D\C$\Prueba\EspacioE.html
#ConvertTo-Html -Body "Espacio Ocupado vdiskcache: $vdiskcachetf"  -CssUri \\XAPSDKBPV01D\C$\Prueba\style.css  >> \\XAPSDKBPV01D\C$\Prueba\EspacioE.html
#ConvertTo-Html -Body "Espacio Ocupado PageFile.sys: $pageFiletf" -CssUri \\XAPSDKBPV01D\C$\Prueba\style.css >> \\XAPSDKBPV01D\C$\Prueba\EspacioE.html
#ConvertTo-Html -Body "Espacio Total E: $TotalSizef" -CssUri \\XAPSDKBPV01D\C$\Prueba\style.css >> \\XAPSDKBPV01D\C$\Prueba\EspacioE.html
#ConvertTo-Html -Body "Espacio Libre E: $FreeSpacef" -CssUri \\XAPSDKBPV01D\C$\Prueba\style.css >> \\XAPSDKBPV01D\C$\Prueba\EspacioE.html
#ConvertTo-Html -Body "Espacio Usado E: $UsageSpacef" -CssUri \\XAPSDKBPV01D\C$\Prueba\style.css >> \\XAPSDKBPV01D\C$\Prueba\EspacioE.html
#ConvertTo-Html -Body "" -CssUri \\XAPSDKBPV01D\C$\Prueba\style.css >> \\XAPSDKBPV01D\C$\Prueba\EspacioE.html
#ConvertTo-Html -Body "" -CssUri \\XAPSDKBPV01D\C$\Prueba\style.css >> \\XAPSDKBPV01D\C$\Prueba\EspacioE.html

Write-host

}

Invoke-Expression $scriptDir\EspacioE.html