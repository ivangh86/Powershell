if ((Get-PSSnapin "Citrix.Broker.Admin.*" -EA silentlycontinue) -eq $null) {
try { Add-PSSnapin Citrix* -ErrorAction Stop }
catch { write-error "Error loading Citrix PowerShell snapin"; Return }
}

$Machines = Get-BrokerMachine -MachineName "TSB\WCPNHIDSD*"
$contador = 0
$contador
foreach ($Machine in $Machines) {
$Machine.DNSName
$Events = Get-EventLog Application -Newest 1000 -ComputerName $Machine.DNSName | Where-Object { $_.Source -Like "User*" -and #> $_.EventID -eq "1511"} #| Format-Table TimeWritten, EventID, Message -auto
$Events


if ($Events -ne "") {

       Foreach ($Event in $Events){

        $contador++
        }
   }


}

$contador 

