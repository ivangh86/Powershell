$DesktopsGroups =  Get-BrokerDesktopGroup * |Select-Object * -First 1



$apps = Get-BrokerApplication * | Select-Object * -First 1


foreach ($app in $apps){

    $app.AllAssociatedDesktopGroupUids 
    if  $app.AllAssociatedDesktopGroupUids  -contains 

    }


$Objeto = New-Object PSObject
Add-Member -InputObject $Objeto -MemberType NoteProperty -Name "propiedad1" -Value $apps.Name

$Objeto