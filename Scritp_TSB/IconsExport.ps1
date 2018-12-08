# Adding Citrix Snapins
Add-PSSnapin Citrix*

# Grabs applications
$apps = Get-BrokerApplication -MaxRecordCount 2147483647  |Where-Object {$_.machineConfigurationNames -ne ""} | Sort-Object -Property PublishedName

$Results = @()

foreach($app in $apps)
{
    $Properties = @{
    Name = $app.name
    EncodedIconData = (Get-Brokericon -Uid $app.IconUid).EncodedIconData # Grabs Icon Image
	}
    # Stores each Application setting for export
    $Results += New-Object psobject -Property $properties
}

$Results | export-clixml .\iconos.xml