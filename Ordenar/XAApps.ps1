# ==============================================================================================
# 
# 
# NAME: XAApps.ps1
# 
# AUTHOR: Kevin McLeman 
# DATE  : 4/25/2011
# 
# COMMENT: 
#       Used for importing and exporting XenApp published resources. Works with XenApp 5 and up.
#
# Application Types:
#	 ServerInstalled
#	 ServerDesktop
#	 Content
#	 StreamedToServer
#	 StreamedToClient
#	 StreamedToClientOrInstalled
#	 StreamedToClientOrStreamedToServer
#
# SWITCHES:
#	 -CSV 			 Required Name of CSV file to import or export from/to
#	 -XA5			 Required for exporting from XenApp 5 using XenApp Commands CTP v3
#	 -Import 		 Include this parameter to import applications from CSV
#	 -Export 		 Include this parameter to export applications from CSV
#	 -WorkerGroups 	 Include to export workergroups with published resources
#	 -Servers		 Include to export servers with published resources
#	 -Help			 Displays the help topic
#		
#    	<example>
#	    		powershell.exe XAApps.ps1 -CSV "C:\Test.csv" -Export -Workergroups -Servers
#               powershell.exe XAApps.ps1 -CSV "C:\Test.csv" -Import
#			<script language="PowerShell">
# ==============================================================================================

param(
	  $CSV,
	  [switch]$Import,
	  [switch]$Export,
	  [switch]$Servers,
	  [switch]$WorkerGroups,
	  [switch]$XA5,
	  [switch]$help
	  )
 
	  
$erroractionpreference = "SilentlyContinue" #"Stop" "Inquire"
cls

## Add Citrix cmdlets
Add-PSSnapin Citrix* 

	function log-XenApps ($XAData)
	{Out-File -FilePath $LogXAApps -InputObject $XAData -Append} 

###########################################################################################################
##                        Export XenApp Applications
###########################################################################################################

function XAExport
{
	$LogXAApps = $CSV
	

	## Create CSV Headers
	log-XenApps "DisplayName,BrowserName,Enabled,HideWhenDisabled,Workergroup,Description,AccountDisplayName,AddToClientDesktop,AddToClientStartMenu,WaitOnPrinterCreation,ClientFolder,ApplicationType,ServerNames,CommandLineExecutable,WorkingDir,ContentAddress,ProfileLocation,ProfileProgramName,FolderPath,WindowType,OfflineAccessAllowed,EncodedIconData"
	
	$XAAppArray = @{}
	
	Write-Host "Exporting XenApp Applications to: " $CSV -ForegroundColor Green
	
	$XAApps = get-xaapplication
	
	foreach($XAApp in $XAApps)
	{
		## Retrieve CTP v3 cmdlet specific data that is different from XenApp 6 PowerShell SDK
		If ($XA5 -eq $True)
		{	
			$Account = Get-xaaccount -BrowserName $($XAApp.browsername) | Select-Object $_.AccountDisplayName
			$XAIcon = Get-xaapplicationIconStream -BrowserName $($XAApp.browsername) | Select-Object EncodedIconData
			If ($($XAIcon.EncodedIconData.length) -gt "25000")
				{$XenAppIcon = " ,"}
			Else
				{$XenAppIcon = $XAIcon.EncodedIconData}
				
			$Accounts = [string]::join(",",$Account)
			$Accts = '"' + $Accounts + '"'
			
			If (($Servers -eq $True) -and ($WorkerGroups -eq $True))
			{
				If ($($XAApp.ApplicationType) -eq "ServerInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToServer" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrStreamedToServer" -or $($XAApp.ApplicationType) -eq "ServerDesktop")
				{
					$XAServerNames = Get-xaserver -BrowserName $($XAApp.browsername) | Select-Object $_.ServerName
					$XAServer = [string]::join(",",$XAServerNames)
					$ServerNames = '"' + $XAServer + '"'
					$WorkerGroup = ""
				}
				Else
				{ 
					$ServerNames = ""
					$WorkerGroup = ""
				}
			}
			
			ElseIf (($WorkerGroups -eq $True) -and ($Servers -eq $False))
			{
				$ServerNames = ""
				$WorkerGroup = ""
			}
				
			ElseIf (($Servers -eq $True) -and ($WorkerGroups -eq $False))
			{
				If ($($XAApp.ApplicationType) -eq "ServerInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToServer" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrStreamedToServer" -or $($XAApp.ApplicationType) -eq "ServerDesktop")
				{
					$XAServerNames = Get-xaserver -BrowserName $($XAApp.browsername) | Select-Object $_.ServerName
					$XAServer = [string]::join(",",$XAServerNames)
					$ServerNames = '"' + $XAServer + '"'
					$WorkerGroup = ""
				}
				Else
				{ 
					$ServerNames = ""
					$WorkerGroup = ""
				}
			}	
			
			$AppDescription = '"' + $($XAApp.Description) + '"'
			If ($($XAApp.CommandLineExecutable) -like '"*')
				{
				$XACommandLine = '"' + $($XAApp.CommandLineExecutable) + " " + $($XAApp.CommandLineArguments)
				}
			Else
				{
				$XACommandLine = $($XAApp.CommandLineExecutable) + " " + $($XAApp.CommandLineArguments) 
				}

			If ($($XAApp.ContentAddress) -like '"*')
				{
				$XAContent = '"' + $($XAApp.ContentAddress)
				}
			Else
				{
				$XAContent = $($XAApp.ContentAddress) 
				}
				
			If ($($XAApp.WorkingDirectory) -like '"*')
				{
				$XAWorkingDir = '"' + $($XAApp.WorkingDirectory)
				}
			Else
				{
				$XAWorkingDir = $($XAApp.WorkingDirectory)
				}
		}
		Else
		{
			
			$Account = Get-xaapplication -BrowserName $($XAApp.browsername) | Get-xaaccount | Select-Object $_.AccountDisplayName
			$XAIcon = Get-xaapplicationIcon -BrowserName $($XAApp.browsername) | Select-Object EncodedIconData
			If ($($XAIcon.EncodedIconData.length) -gt "25000")
				{$XenAppIcon = " ,"}
			Else
				{$XenAppIcon = $XAIcon.EncodedIconData}
				
			$Accounts = [string]::join(",",$Account)
			$Accts = '"' + $Accounts + '"'
			
			If (($Servers -eq $True) -and ($WorkerGroups -eq $True))
			{
				If ($($XAApp.ApplicationType) -eq "ServerInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToServer" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrStreamedToServer" -or $($XAApp.ApplicationType) -eq "ServerDesktop")
				{
					$WorkerGroupName = Get-xaapplication -BrowserName $($XAApp.browsername) | Get-xaworkergroup | Select-Object $_.WorkerGroupName
					$XAServerNames = Get-xaapplication -BrowserName $($XAApp.browsername) | Get-xaserver | Select-Object $_.ServerName
					$XAServer = [string]::join(",",$XAServerNames)
					$ServerNames = '"' + $XAServer + '"'
					$WrkrGrp = [string]::join(",",$WorkerGroupName)
					$WorkerGroup = '"' + $WrkrGrp + '"'
				}
				Else
				{ 
					$ServerNames = ""
					$WorkerGroup = ""
				}
			}
			
			ElseIf (($WorkerGroups -eq $True) -and ($Servers -eq $False))
			{
				If ($($XAApp.ApplicationType) -eq "ServerInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToServer" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrStreamedToServer" -or $($XAApp.ApplicationType) -eq "ServerDesktop")
				{
					$WorkerGroupName = Get-xaapplication -BrowserName $($XAApp.browsername) | Get-xaworkergroup | Select-Object $_.WorkerGroupName
					$WrkrGrp = [string]::join(",",$WorkerGroupName)
					$WorkerGroup = '"' + $WrkrGrp + '"'
					$ServerNames = ""
				}
				Else
				{ 
					$ServerNames = ""
					$WorkerGroup = ""
				}
			}
				
			ElseIf (($Servers -eq $True) -and ($WorkerGroups -eq $False))
			{
				If ($($XAApp.ApplicationType) -eq "ServerInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrInstalled" -or $($XAApp.ApplicationType) -eq "StreamedToServer" -or $($XAApp.ApplicationType) -eq "StreamedToClientOrStreamedToServer" -or $($XAApp.ApplicationType) -eq "ServerDesktop")
				{
					$XAServerNames = Get-xaapplication -BrowserName $($XAApp.browsername) | Get-xaserver | Select-Object $_.ServerName
					$XAServer = [string]::join(",",$XAServerNames)
					$ServerNames = '"' + $XAServer + '"'
					$WorkerGroup = ""
				}
				Else
				{ 
					$ServerNames = ""
					$WorkerGroup = ""
				}
			}	
			
			$AppDescription = '"' + $($XAApp.Description) + '"'
			If ($($XAApp.CommandLineExecutable) -like '"*')
				{
				$XACommandLine = '"' + $($XAApp.CommandLineExecutable)
				}
			Else
				{
				$XACommandLine = $($XAApp.CommandLineExecutable)
				}

			If ($($XAApp.ContentAddress) -like '"*')
				{
				$XAContent = '"' + $($XAApp.ContentAddress)
				}
			Else
				{
				$XAContent = $($XAApp.ContentAddress) 
				}
				
			If ($($XAApp.WorkingDirectory) -like '"*')
				{
				$XAWorkingDir = '"' + $($XAApp.WorkingDirectory)
				}
			Else
				{
				$XAWorkingDir = $($XAApp.WorkingDirectory)
				}
		}
		
		## Write data to CSV file
		log-XenApps "$($XAApp.DisplayName),$($XAApp.BrowserName),$($XAApp.Enabled),$($XAApp.HideWhenDisabled),$($WorkerGroup),$($AppDescription),$($Accts),$($XAApp.AddToClientDesktop),$($XAApp.AddToClientStartMenu),$($XAApp.WaitOnPrinterCreation),$($XAApp.ClientFolder),$($XAApp.ApplicationType),$($ServerNames),$($XACommandLine),$($XAWorkingDir),$($XAContent),$($XAApp.ProfileLocation),$($XAApp.ProfileProgramName),$($XAApp.FolderPath),$($XAApp.WindowType),$($XAApp.OfflineAccessAllowed),$($XenAppIcon)"
	}
	
	## Count Application Type ServerInstalled exported
	$ExportSI = $XAApps | Where-Object {$_.ApplicationType -eq "ServerInstalled"}
	if ($ExportSI.Count -eq $null)
		{$OnlineExp = 0}
	else
		{$OnlineExp = $ExportSI.Count}
	
	## Count Application Type StreamedToClientOrInstalled exported	
	$ExportSCSI = $XAApps | Where-Object {$_.ApplicationType -eq "StreamedToClientOrInstalled"}
	if ($ExportSCSI.Count -eq $null)
		{$STCOnlineExp = 0}
	else
		{$STCOnlineExp = $ExportSCSI.Count}
	
	## Count Application Type StreamedToClientOrStreamedToServer exported	
	$ExportSCSS = $XAApps | Where-Object {$_.ApplicationType -eq "StreamedToClientOrStreamedToServer"}
	if ($ExportSCSS.Count -eq $null)
		{$STCSTSExp = 0}
	else
		{$STCSTSExp = $ExportSCSS.Count}
		
	## Count Application Type StreamedToServer exported	
	$ExportSTS = $XAApps | Where-Object {$_.ApplicationType -eq "StreamedToServer"}
	if ($ExportSTS.Count -eq $null)
		{$STSExp = 0}
	else
		{$STSExp = $ExportSTS.Count}
		
	## Count Application Type StreamedToClient exported	
	$ExportSTC = $XAApps | Where-Object {$_.ApplicationType -eq "StreamedToClient"}
	if ($ExportSTC.Count -eq $null)
		{$STCExp = 0}
	else
		{$STCExp = $ExportSTC.Count}
		
	## Count Application Type Content exported	
	$ExportContent = $XAApps | Where-Object {$_.ApplicationType -eq "Content"}
	if ($ExportContent.Count -eq $null)
		{$ContentExp = 0}
	else
		{$ContentExp = $ExportContent.Count}
		
	## Count Application Type ServerDesktop exported	
	$ExportServerDesktop = $XAApps | Where-Object {$_.ApplicationType -eq "ServerDesktop"}
	if ($ExportServerDesktop.Count -eq $null)
		{$SrvrDesktopExp = 0}
	else
		{$SrvrDesktopExp = $ExportServerDesktop.Count}
	
	## Add all exported resources for total	
	$ExportAll = $OnlineExp + $ContentExp + $STCOnlineExp + $STCExp + $STSExp + $STCSTSExp + $SrvrDesktopExp
	
	## Write number of exported resources by type to PowerShell console
	Write-Host "XA Server Installed Apps Exported: " $OnlineExp -ForegroundColor Yellow
	Write-Host "XA Published Content Exported: " $ContentExp -ForegroundColor Yellow
	Write-Host "XA Streamed to client Apps Exported: " $STCExp -ForegroundColor Yellow
	Write-Host "XA Streamed to server Apps Exported: " $STSExp -ForegroundColor Yellow
	Write-Host "XA Server Desktops Exported: " $SrvrDesktopExp -ForegroundColor Yellow
	Write-Host "XA Streamed to client Fallback to XA Server Installed Exported: " $STCOnlineExp -ForegroundColor Yellow
	Write-Host "XA Streamed to client Fallback to XA Streamed to Server Apps Exported: " $STCSTSExp -ForegroundColor Yellow
	Write-Host "XA Apps Exported Total: " $ExportAll -ForegroundColor White 
	
}

###########################################################################################################
##                        Import XenApp Applications
###########################################################################################################

function XAImport
{
	$ImportXA = Import-Csv $CSV
	
	Write-Host "Importing XenApp Applications From: " $CSV -ForegroundColor Green
	
	foreach($Import in $ImportXA)
	{
	 	$AccountArray = ($($Import.AccountDisplayName)).split(”,“)
		If ($($Import.CommandLineExecutable) -like '*"*')
			{
			$XACommandLine = '"' + $($Import.CommandLineExecutable)
			}
		Else
			{
			$XACommandLine = $($Import.CommandLineExecutable)
			}
		If ($($Import.ContentAddress) -like '*"*')
			{
			$ImpXAContent = '"' + $($Import.ContentAddress)
			}
		Else
			{
			$ImpXAContent = $($Import.ContentAddress)
			}
		If ($($Import.WorkingDir) -like '*"*')
			{
			$XAWorkingDirectory = '"' + $($Import.WorkingDir)
			}
		Else
			{
			$XAWorkingDirectory = $($Import.WorkingDir)
			}
		If ($($Import.OfflineAccessAllowed) -eq "TRUE")
			{
			$OfflineAccess = $TRUE
			}
		Else
			{
			$OfflineAccess = $FALSE
			}
		If ($($Import.Enabled) -eq "TRUE")
			{
			$PubEnabled = $TRUE
			}
		Else
			{
			$PubEnabled = $FALSE
			}
		If ($($Import.HideWhenDisabled) -eq "TRUE")
			{
			$Hide = $TRUE
			}
		Else
			{
			$Hide = $FALSE
			}
		If ($($Import.AddToClientDesktop) -eq "TRUE")
			{
			$AddToDesktop = $TRUE
			}
		Else
			{
			$AddToDesktop = $FALSE
			}
		If ($($Import.AddToClientStartMenu) -eq "TRUE")
			{
			$AddToStartMenu = $TRUE
			}
		Else
			{
			$AddToStartMenu = $FALSE
			}
		If ($($Import.WaitOnPrinterCreation) -eq "TRUE")
			{
			$PrinterWait = $TRUE
			}
		Else
			{
			$PrinterWait = $FALSE
			}
			
		## The following code imports published resources based on Application Type. 
		## Resources published by Server Name assignment if applicable, Worker Group assignment if applicable, both or neither in the case of streamed to client etc.
	 	If ($($Import.EncodedIconData) -eq "")
	 	{
	 		If ($($Import.ServerNames) -eq "" -and $($Import.WorkerGroup) -eq "")
	 		{
				$WorkerGroupArray = ($($Import.WorkerGroup)).split(”,“)
				$ServerArray = ($($Import.ServerNames)).split(”,“)
				If ($($Import.ApplicationType) -eq "ServerInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "Content")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ContentAddress $($ImpXAContent) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "ServerDesktop")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrStreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClient")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -OfflineAccessAllowed $OfflineAccess -Force -SkipPassThru -Accounts $AccountArray
				}
			}
	 		ElseIf ($($Import.ServerNames) -eq "" -and $($Import.WorkerGroup) -gt "")
	 		{
				$WorkerGroupArray = ($($Import.WorkerGroup)).split(”,“)
				$ServerArray = ($($Import.ServerNames)).split(”,“)
				If ($($Import.ApplicationType) -eq "ServerInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "Content")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ContentAddress $($ImpXAContent) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "ServerDesktop")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrStreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClient")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -OfflineAccessAllowed $OfflineAccess -Force -SkipPassThru -Accounts $AccountArray
				}
			}
	 		ElseIf ($($Import.ServerNames) -gt "" -and $($Import.WorkerGroup) -eq "")
	 		{
				$WorkerGroupArray = ($($Import.WorkerGroup)).split(”,“)
				$ServerArray = ($($Import.ServerNames)).split(”,“)
				If ($($Import.ApplicationType) -eq "ServerInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "Content")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ContentAddress $($ImpXAContent) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "ServerDesktop")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrStreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClient")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -AddToClientDesktop $FALSE -AddToClientStartMenu $AddToDesktop -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -OfflineAccessAllowed $OfflineAccess -Force -SkipPassThru -Accounts $AccountArray
				}
			}
	 		ElseIf ($($Import.ServerNames) -gt "" -and $($Import.WorkerGroup) -gt "")
	 		{
				$WorkerGroupArray = ($($Import.WorkerGroup)).split(”,“)
				$ServerArray = ($($Import.ServerNames)).split(”,“)
				If ($($Import.ApplicationType) -eq "ServerInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "Content")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ContentAddress $($ImpXAContent) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "ServerDesktop")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrStreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClient")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -OfflineAccessAllowed $OfflineAccess -Force -SkipPassThru -Accounts $AccountArray
				}
			}
		}
		Else
		{
	 		If ($($Import.ServerNames) -eq "" -and $($Import.WorkerGroup) -eq "")
	 		{
				$WorkerGroupArray = ($($Import.WorkerGroup)).split(”,“)
				$ServerArray = ($($Import.ServerNames)).split(”,“)
				If ($($Import.ApplicationType) -eq "ServerInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "Content")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ContentAddress $($ImpXAContent) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "ServerDesktop")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrStreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClient")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -OfflineAccessAllowed $OfflineAccess -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)
				}
			}
	 		ElseIf ($($Import.ServerNames) -eq "" -and $($Import.WorkerGroup) -gt "")
	 		{
				$WorkerGroupArray = ($($Import.WorkerGroup)).split(”,“)
				$ServerArray = ($($Import.ServerNames)).split(”,“)
				If ($($Import.ApplicationType) -eq "ServerInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "Content")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ContentAddress $($ImpXAContent) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "ServerDesktop")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrStreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClient")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -OfflineAccessAllowed $OfflineAccess -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)
				}
			}
	 		ElseIf ($($Import.ServerNames) -gt "" -and $($Import.WorkerGroup) -eq "")
	 		{
				$WorkerGroupArray = ($($Import.WorkerGroup)).split(”,“)
				$ServerArray = ($($Import.ServerNames)).split(”,“)
				If ($($Import.ApplicationType) -eq "ServerInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "Content")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ContentAddress $($ImpXAContent) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "ServerDesktop")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrStreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClient")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -OfflineAccessAllowed $OfflineAccess -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)
				}
			}
	 		ElseIf ($($Import.ServerNames) -gt "" -and $($Import.WorkerGroup) -gt "")
	 		{
				$WorkerGroupArray = ($($Import.WorkerGroup)).split(”,“)
				$ServerArray = ($($Import.ServerNames)).split(”,“)
				If ($($Import.ApplicationType) -eq "ServerInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "Content")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ContentAddress $($ImpXAContent) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "ServerDesktop")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrInstalled")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -CommandLineExecutable $($XACommandLine) -WorkingDirectory $($XAWorkingDirectory) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClientOrStreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -OfflineAccessAllowed $OfflineAccess -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToServer")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -WorkerGroupNames $($WorkerGroupArray) -ServerNames $($ServerArray) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -WaitOnPrinterCreation $PrinterWait -WindowType $($Import.WindowType) -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)				
				}
				ElseIf ($($Import.ApplicationType) -eq "StreamedToClient")
				{
				New-xaapplication -ApplicationType $($Import.ApplicationType) -DisplayName $($Import.DisplayName) -BrowserName $($Import.BrowserName) -Enabled $PubEnabled -HideWhenDisabled $Hide -Description $($Import.Description) -ProfileLocation $($Import.ProfileLocation) -ProfileProgramName $($Import.ProfileProgramName) -AddToClientDesktop $AddToDesktop -AddToClientStartMenu $AddToStartMenu -FolderPath $($Import.FolderPath) -ClientFolder $($Import.ClientFolder) -OfflineAccessAllowed $OfflineAccess -Force -SkipPassThru -Accounts $AccountArray -EncodedIconData $($Import.EncodedIconData)
				}
			}
		}
	}
	
	$ImportSI = $ImportXA | Where-Object {$_.ApplicationType -eq "ServerInstalled"}
	if ($ImportSI.Count -eq $null)
		{$OnlineImp = 0}
	else
		{$OnlineImp = $ImportSI.Count}
		
	$ImportSCSI = $ImportXA | Where-Object {$_.ApplicationType -eq "StreamedToClientOrInstalled"}
	if ($ImportSCSI.Count -eq $null)
		{$STCOnlineImp = 0}
	else
		{$STCOnlineImp = $ImportSCSI.Count}
		
	$ImportSCSS = $ImportXA | Where-Object {$_.ApplicationType -eq "StreamedToClientOrStreamedToServer"}
	if ($ImportSCSS.Count -eq $null)
		{$STCSTSImp = 0}
	else
		{$STCSTSImp = $ImportSCSS.Count}
		
	$ImportSTS = $ImportXA | Where-Object {$_.ApplicationType -eq "StreamedToServer"}
	if ($ImportSTS.Count -eq $null)
		{$STSImp = 0}
	else
		{$STSImp = $ImportSTS.Count}
		
	$ImportSTC = $ImportXA | Where-Object {$_.ApplicationType -eq "StreamedToClient"}
	if ($ImportSTC.Count -eq $null)
		{$STCImp = 0}
	else
		{$STCImp = $ImportSTC.Count}
		
	$ImportContent = $ImportXA | Where-Object {$_.ApplicationType -eq "Content"}
	if ($ImportContent.Count -eq $null)
		{$ContentImp = 0}
	else
		{$ContentImp = $ImportContent.Count}
		
	$ImportServerDesktop = $ImportXA | Where-Object {$_.ApplicationType -eq "ServerDesktop"}
	if ($ImportServerDesktop.Count -eq $null)
		{$SrvrDesktopImp = 0}
	else
		{$SrvrDesktopImp = $ImportServerDesktop.Count}
		
	$ImportAll = $OnlineImp + $ContentImp + $STCOnlineImp + $STCImp + $STSImp + $STCSTSImp + $SrvrDesktopImp
	
	## Write number of imported resources by type to PowerShell console
	Write-Host "XA Server Installed Apps Imported: " $OnlineImp -ForegroundColor Yellow
	Write-Host "XA Published Content Imported: " $ContentImp -ForegroundColor Yellow
	Write-Host "XA Streamed to client Apps Imported: " $STCImp -ForegroundColor Yellow
	Write-Host "XA Streamed to server Apps Imported: " $STSImp -ForegroundColor Yellow
	Write-Host "XA Server Desktops Imported: " $SrvrDesktopImp -ForegroundColor Yellow
	Write-Host "XA Streamed to client Fallback to XA Server Installed Imported: " $STCOnlineImp -ForegroundColor Yellow
	Write-Host "XA Streamed to client Fallback to XA Streamed to Server Apps Imported: " $STCSTSImp -ForegroundColor Yellow
	Write-Host "XA Apps Imported Total: " $ImportAll -ForegroundColor White 

}

###########################################################################################################
##                        Execute Function for Import/Export of XenApp Applications
###########################################################################################################

If (($Import -eq $True) -and ($Export -eq $False))
	{
	XAImport
	}
ElseIf (($Export -eq $True) -and ($Import -eq $False))
	{
	XAExport
	}

###########################################################################################################	
##                      Help Information
###########################################################################################################	


	
function funHelp()
{
$helpText=@"
DESCRIPTION:
NAME: XAApps.ps1 
For importing and exporting XenApp 6.0 applications.

PARAMETERS: 
	-CSV 			Required Name of CSV file to import or export from/to
	-XA5			Required for exporting from XenApp 5 using XenApp Commands CTP v3
	-Import 		Include this parameter to import applications from CSV
	-Export 		Include this parameter to export applications from CSV
	-WorkerGroups 		Include to export workergroups with published resources
	-Servers		Include to export servers with published resources
	-Help			Displays the help topic

Notes:
	For exporting resources from XenApp 5 please install XenApp Commands CTP v3
	from https://www.citrix.com/English/ss/downloads/details.asp?downloadId=1687620&productId=186
	CSV fileds are separated by comma.
	Content redirection is not included when exporting or importing
	
Application Types:
	ServerInstalled
	ServerDesktop
	Content
	StreamedToServer
	StreamedToClient
	StreamedToClientOrInstalled
	StreamedToClientOrStreamedToServer


SYNTAX:
	XAApps.ps1 -CSV "c:\Test.csv" -Export -WorkerGroups -Servers
		Exports published resources from local XenApp server to "C:\Test.csv" with both workergroups 
		and servers for each resource.

	XAApps.ps1 -CSV "c:\Test.csv" -Export -WorkerGroups
		Exports published resources from local XenApp server to "C:\Test.csv" with only workergroups 
		for each resource.
		
	XAApps.ps1 -CSV "c:\Test.csv" -Export -Servers -XA5
		Exports published resources from local XenApp 5 server to "C:\Test.csv" with only servers 
		for each resource.
	
	XAApps.ps1 -CSV "c:\Test.csv" -Import
		Imports published resources from "C:\Test.csv" to local XenApp server.

	XAApps.ps1 -help
		Displays the help topic for the script

"@
$helpText
Exit
}
	
if($help){ "Obtaining help ..." ; funHelp }


