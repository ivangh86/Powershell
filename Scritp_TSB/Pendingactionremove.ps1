
$machines = (get-brokerhostingpoweraction -MaxRecordCount 2000 -State "Pending").MachineName
#$machines = "TSB\wcpnhidsd446"
$horafin = (get-date).AddHours(-1)


If ($machines -ne $null) {

foreach ($machine in $machines){

        $Action = get-brokerhostingpoweraction -MaxRecordCount 2000 -State "Pending"  -MachineName $machine
        $Action

        if ($_.RequestTime -lt $horafin )

        {write-host "Acciones pendiente mas de una hora" 
        write-host "Eliminando la Accion pendiente en $machine " 

        $Action |Remove-BrokerHostingPowerAction

        write-host "--------------------------"
        write-host "--------------------------"
        write-host "ELIMINADA"
        write-host "--------------------------"
        write-host "--------------------------"

        }


        else {write-host "No hay acciones pendientes de mas de una hora"}

        }

Else {write-host "no hay acciones pendientes"}

}