asnp Citrix*

$servers = Get-XAServer -ServerName XACISV1025P
#$servers = Get-XAServer | Select-Object ServerName
foreach ($xserver in $servers) {
            write-host $xserver.Servername
            $hotfix = Get-XAServerHotfix $Xserver.Servername | select-object hotfixname
            write-host “—————–“
            foreach ($xhotfix in $hotfix) {
                       write-host $xhotfix.hotfixname }
            write-host “—————–“ }