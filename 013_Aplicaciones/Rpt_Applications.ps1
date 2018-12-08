

#*************************************************************
#  Checklist Granja TSB - UK
#  Catalogo de Aplicaciones TSB UK
#  v 1.0 - 09/03/18 - John Carranza
#*************************************************************


$vElapsedTime = [System.Diagnostics.Stopwatch]::Startnew()

if ((Get-PSSnapin "Citrix.*" -EA silentlycontinue) -eq $null) {
	try { Add-PSSnapin Citrix.* -ErrorAction Stop }
	catch { write-error "Error loading XenApp Powershell snapin"; Return }
}

$vMyInvocation = (Get-Variable MyInvocation).Value
$vPathActual = Split-Path $vMyInvocation.MyCommand.Path

$vRpt_Applications = $vPathActual + "\Report\Rpt_Applications" + ".csv"
$vRpt_Applications_html = $vPathActual + "\Report\Rpt_Applications" + ".html"

$vRpt_Applications_Temp = "C:\Temp\Rpt_Applications" + ".csv"
$vRpt_Applications_html_Temp = "C:\Temp\Rpt_Applications" + ".html"

If (Test-Path $vRpt_Applications_Temp){
    Remove-Item $vRpt_Applications_Temp
}

If (Test-Path $vRpt_Applications_html_Temp){
    Remove-Item $vRpt_Applications_html_Temp
}


Write-Host "Procesando...!!!"
$vBrokerApplications = Get-BrokerApplication -MaxRecordCount 10000| Select ApplicationName,Enabled,AdminFolderName,CommandLineExecutable,AssociatedUserFullNames,AllAssociatedDesktopGroupUids

"ApplicationName;Enabled;AdminFolderName;CommandLineExecutable;AssociatedUserFullNames;DeliveryGroups" | Out-File $vRpt_Applications_Temp -Append -Encoding ASCII


"<p align=Center><font size=6 face=Bodoni MT color=#003333><B>Catálogo de Aplicaciones TSB UK</B></font></p>" | Out-File $vRpt_Applications_html_Temp -Append
"<p align=Center><font size=4 face=Bodoni MT color=#330033><B>Granja Producción</B></font></p>" | Out-File $vRpt_Applications_html_Temp -Append
"<p align=Center><table BORDER=1 width=100% cellspacing=0 cellpadding=3>" | Out-File $vRpt_Applications_html_Temp -Append
"<tr>" | Out-File $vRpt_Applications_html_Temp -Append
"<th bgcolor=#DDDDFF colspan=1 width=100><p align=center><b><font face=Verdana size=2 color=#000000>Aplicación</font></b></p></th>" | Out-File $vRpt_Applications_html_Temp -Append
"<th bgcolor=#DDDDFF colspan=1 width=50><p align=center><b><font face=Verdana size=2 color=#000000>Estado</font></b></p></th>" | Out-File $vRpt_Applications_html_Temp -Append
"<th bgcolor=#DDDDFF colspan=1 width=100><p align=center><b><font face=Verdana size=2 color=#000000>Ubicación</font></b></p></th>" | Out-File $vRpt_Applications_html_Temp -Append
"<th bgcolor=#DDDDFF colspan=1 width=400><p align=center><b><font face=Verdana size=2 color=#000000>Comando</font></b></p></th>" | Out-File $vRpt_Applications_html_Temp -Append
"<th bgcolor=#DDDDFF colspan=1 width=300><p align=center><b><font face=Verdana size=2 color=#000000>Grupo Usuarios</font></b></p></th>" | Out-File $vRpt_Applications_html_Temp -Append
"<th bgcolor=#DDDDFF colspan=1 width=300><p align=center><b><font face=Verdana size=2 color=#000000>Delivery Group</font></b></p></th>" | Out-File $vRpt_Applications_html_Temp -Append

"</tr>" | Out-File $vRpt_Applications_html_Temp -Append

$vcolor = ""

Foreach($vBrokerApplication in $vBrokerApplications){
    $vApplicationName = $vBrokerApplication.ApplicationName
    $vEnabled = $vBrokerApplication.Enabled
    $vAdminFolderName = $vBrokerApplication.AdminFolderName
    $vCommandLineExecutable = $vBrokerApplication.CommandLineExecutable
    $vAssociatedUserFullNames = $vBrokerApplication.AssociatedUserFullNames
    $vAllAssociatedDesktopGroupUids = $vBrokerApplication.AllAssociatedDesktopGroupUids
    

    If ($vEnabled -eq "True"){
        $vcolor = "Black"
    }else{
        $vcolor = "Red"
    }

    If(!($vAdminFolderName)){
        $vAdminFolderName = "-"
    }

    $vDeliveryGroup = ""
    $vDeliveryGroups = ""

    Foreach($vAllAssociatedDesktopGroupUid in $vAllAssociatedDesktopGroupUids){
        $vDeliveryGroup = Get-BrokerDesktopGroup | Where-Object {($_.Uid -eq $vAllAssociatedDesktopGroupUid)} | Select Name
        $vDeliveryGroups = $vDeliveryGroups + " - " + $vDeliveryGroup.Name
    }
    
    "$vApplicationName;$vEnabled;$vAdminFolderName;$vCommandLineExecutable;$vAssociatedUserFullNames;$vDeliveryGroups" | Out-File $vRpt_Applications_Temp -Append -Encoding ASCII

    "<tr><td><p align=center><font face=Verdana color= $vcolor  size=2>" + $vApplicationName + "</font></TD>" | Out-File $vRpt_Applications_html_Temp -Append
    "<td><font face=Verdana color= $vcolor size=2>" + $vEnabled + "</font></TD>" | Out-File $vRpt_Applications_html_Temp -Append
    "<td><font face=Verdana color= $vcolor size=2>" + $vAdminFolderName + "</font></TD>" | Out-File $vRpt_Applications_html_Temp -Append
    "<td><font face=Verdana color= $vcolor size=2>" + $vCommandLineExecutable + "</font></TD>" | Out-File $vRpt_Applications_html_Temp -Append
    "<td><font face=Verdana color= $vcolor size=2>" + $vAssociatedUserFullNames + "</font></TD>" | Out-File $vRpt_Applications_html_Temp -Append
    "<td><font face=Verdana color= $vcolor size=2>" + $vDeliveryGroups + "</font></TD>" | Out-File $vRpt_Applications_html_Temp -Append
    
}

"</table>" | Out-File $vRpt_Applications_html_Temp -Append

If (Test-Path $vRpt_Applications){
    Remove-Item $vRpt_Applications
}

If (Test-Path $vRpt_Applications_html){
    Remove-Item $vRpt_Applications_html
}

copy-item $vRpt_Applications_Temp -Destination $vRpt_Applications -force
copy-item $vRpt_Applications_html_Temp -Destination $vRpt_Applications_html -force

If (Test-Path $vRpt_Applications_Temp){
    Remove-Item $vRpt_Applications_Temp
}

If (Test-Path $vRpt_Applications_html_Temp){
    Remove-Item $vRpt_Applications_html_Temp
}

$vTimeExecution= $vElapsedTime.Elapsed.ToString()

Write-Host $vTimeExecution