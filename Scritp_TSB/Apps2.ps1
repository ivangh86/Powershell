# Collect XenApp 7.x Published Application properties and export to Excel (Tested on Xenapp 7.8)
# Requires Excel to be installed
# Requires the Citrix XenApp PowerShell SDK 

#Complied by anandkumar.n

add-pssnapin citrix*



$csvfilename = "\CitrixApps.csv"
$csvfile = $csvfilename

Get-BrokerApplication * | select AdminFolderName, ApplicationType, ApplicationName, Description, Enabled, CommandLineExecutable, CommandLineArguments, WorkingDirectory, IconFromClient, PublishedName, WaitForPrinterCreation, ShortcutAddedToDesktop, ShortcutAddedToStartMenu, @{n="AssociatedDesktopGroupUids";e={[string]::join(" ; ", $_.AssociatedDesktopGroupUids)}}, @{n="AssociatedUserFullNames";e={[string]::join(" ; ", $_.AssociatedUserFullNames)}}, @{n="AssociatedUserNames";e={[string]::join(" ; ", $_.AssociatedUserNames)}},@{n="Tags";e={[string]::join(" ; ", $_.Tags)}} | Export-Csv -NoTypeInformation $csvFile

$processes = Import-Csv -Path $csvFile 