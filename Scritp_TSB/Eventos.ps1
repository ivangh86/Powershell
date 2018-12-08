if ((Get-PSSnapin "Citrix.Broker.Admin.*" -EA silentlycontinue) -eq $null) {
try { Add-PSSnapin Citrix* -ErrorAction Stop }
catch { write-error "Error loading Citrix PowerShell snapin"; Return }
}

$Logs =$("System", "Application")

$Machines = Get-BrokerMachine -MachineName "TSB\WCPFH*"
$contador = 0
$contador
foreach ($Machine in $Machines) {
$Machine.DNSName
#$Events = Get-EventLog Application -Newest 1000 -ComputerName $Machine.DNSName | Where-Object { $_.Source -Like "Citrix Desktop Service*" -and $_.EventID -eq "1511"} #| Format-Table TimeWritten, EventID, Message -auto
$Events = Get-EventLog Application -Newest 500 -ComputerName $Machine.DNSName | Where-Object { $_.Source -Like "*User Profiles Service*" -and $_.EventID -eq "1530"} #| Format-Table TimeWritten,Source,Security,EventID, Message  -auto
#$Events 


if ($Events -ne "") {

       Foreach ($Event in $Events){
        $Machine.DNSName >> "\\serverhqr1\soft_ctx\Team\Ivan\eventlog.txt"
        $Event >> "\\serverhqr1\soft_ctx\Team\Ivan\eventlog.txt"
        $Event | Format-List
        $contador++
        }
   }


}

$contador 

