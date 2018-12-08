 

  
  #Grupos
  $Giconos_Rochade = "GICONOS SSCC TS Rochade"
  $GTS_Rochade = "GTS_Icono_Rochade"  


  #Usuarios
  $Usuarios_Giconos_Rochade = (Get-ADGroupMember  "GICONOS SSCC TS Rochade").SamAccountName
  $Usuarios_GTS_Rochade = (Get-ADGroupMember "GTS_Icono_Rochade").SamAccountName



  
  #Hacemos backup de los grupos
  #$GICONOS_Rochade > "C:\temp\GICONOS SSCC TS Rochade.txt"
  #$GTS_Rochade  > "C:\temp\GTS_Icono_Rochade.txt"

write-host "************REVISAMOS GTS_Icono_Rochade *********************"

foreach ($user in $Usuarios_Giconos_Rochade) {

TRY{

    if ($Usuarios_GTS_Rochade -notcontains $user )

        {
        Write-host "$user no existe, lo añadimos"
        Add-ADGroupMember $GTS_Rochade –Member $user
        write-host "$user añadido al grupo $GTS_Rochade"
        write-host "Comprobación:"
        (GET-ADUSER –Identity $user –Properties MemberOf | Select-Object MemberOf).MemberOf -match  $Giconos_Rochade
        
        }

    else {  Write-Host " $user Ya existe"}

    }

CATCH 

    {


    Write-host "Se ha producido un problema con el usuario $user"

    }

           

    } 

   