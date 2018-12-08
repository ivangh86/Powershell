Add-PSSnapin citrix.*



$HTML1 = $Header1 | select "Server,Espacio Ocupado vdiskcache,Espacio Ocupado PageFile.sys,Espacio Total E,Espacio Libre E,Espacio Usado E" | ConvertTo-Html -as Table -Title "Informe Unidad E:" <#-PreContent "<p><h3>$Server</h3></p><br>" -PostContent "<br>Informe Realizado $(Get-Date)"#>  > C:\Prueba\EspacioE.html
#ConvertTo-Html -Head $Header1  -Title "Distribución Espacio Unidad E en GB" -Body "Informe sobre la distribucíón del Espacio en la Unidad E: en los equipos de Desarrollo XAV" > \\XAPSDKBPV01D\C$\Prueba\EspacioE.html

$Servers = Get-XAServer | Where-Object { $_.ServerName -like "XAV*"} 
 
foreach ($Server in $Servers)
 
{
$ServerName = $Server.ServerName
#Write-Host $ServerName




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



$ServerName,$vdiskcachetf,$pageFiletf,$TotalSizef,$FreeSpacef,$UsageSpacef
"@

$Table = $Table | ConvertFrom-Csv
$HTML = $Table | select $ServerName,$vdiskcachetf,$pageFiletf,$TotalSizef,$FreeSpacef,$UsageSpacef | ConvertTo-Html -as Table -Title "Informe Unidad E:" -head $Header <#-PreContent "<p><h3>$Server</h3></p><br>" -PostContent "<br>Informe Realizado $(Get-Date)"#>  >> C:\Prueba\EspacioE.html

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

Invoke-Expression C:\Prueba\EspacioE.html