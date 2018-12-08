$CompArr = (Get-BrokerMachine -MachineName "TSB\WCPFHIDSD*").DNSName


Set-Variable -Name EventAgeDays -Value 1     #we will take events for the latest 7 days
#Set-Variable -Name CompArr -Value @($Machine.DNSName)   # replace it with your server names
#Set-Variable -Name CompArr -Value @($Machines)   # replace it with your server names
Set-Variable -Name LogNames -Value @("Application")  # Checking app and system logs
Set-Variable -Name EventTypes -Value @("Error")  # Loading only Errors and Warnings
Set-Variable -Name ExportFolder -Value "\\serverhqr1\soft_ctx\Team\Ivan\"


#$EventID = "1511"


$el_c = @()   #consolidated error log
$now=get-date
$startdate=$now.adddays(-$EventAgeDays)
$ExportFile=$ExportFolder + "el" + $now.ToString("yyyy-MM-dd---hh-mm-ss") + ".csv"  # we cannot use standard delimiteds like ":"

foreach($comp in $CompArr)
{
  foreach($log in $LogNames)
  {
    Write-Host Processing $comp\$log
    $el = get-eventlog -Newest 500 -ComputerName $comp -log $log <#-Before $startdate#> -EntryType $EventTypes| Where-Object { $_.EventID -eq 1511}

    $el_c += $el  #consolidating
  }
}
$el_sorted = $el_c | Sort-Object TimeGenerated    #sort by time
Write-Host Exporting to $ExportFile
$el_sorted|Select EntryType, TimeGenerated, Source, EventID, MachineName, Message | Export-CSV $ExportFile -NoTypeInfo  #EXPORT
Write-Host Done!