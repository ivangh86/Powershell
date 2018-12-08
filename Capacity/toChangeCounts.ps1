#Script para modificar el .blg y solo dejar x contadores

$servers = Get-Content "\\capelo\soft_citrix\Temporal\Ivan_Gomez\Capacity\servers.txt"
$contadores = Get-Content "\\capelo\soft_citrix\Temporal\Ivan_Gomez\Capacity\LOGCounter.txt"

foreach($server in $servers){
    
    $source = "\\capelo\soft_wintel\Citrix\Template_perfmon\20180320_0646\" + $server + "_20180320-0002\Vital Signs.blg"
    $destination = "\\capelo\soft_wintel\Citrix\Template_perfmon\20180320_0646\" + $server + "_20180320-0002\Vital Signs_2.blg"

    relog $source -cf \\capelo\soft_citrix\Temporal\Ivan_Gomez\Capacity\LOGCounter.txt -o $destination

}