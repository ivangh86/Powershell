#Añadiendo Modulos Citrix
Add-PSSnapin Citrix*

$date = Get-Date
$date > "C:\Temp\ReportXenAppImportError.txt"


#Logs

$LogfileErr = "C:\temp\ReportXenAppImportError.txt"

Function LogWrite
{
   Param ([string]$logstring)

   Add-content $LogfileErr -value $logstring
}

# Importando aplicaciones con sus correspondientes iconos 
# Activamos el Grupo de entrega donde queremos que se publiquen las aplicaciones.
$dg = "SDI W2K8R2 OFI DES"
 
# Importando datos de la aplicación
# Activando el Path con el XML apropiado
$apps = Import-Clixml "C:\temp\Apps27072017163947.xml"


foreach ($app in $apps)
{
   #Importando Icono 
   $IconUid = New-BrokerIcon -EncodedIconData $app.EncodedIconData
  
   #Añadimos aplicación
    
    Try
    {
    
    if ($app.CommandLineArguments.Length -lt 2) {$app.CommandLineArguments = " "}
    if ($app.CommandLineExecutable.Length -lt 2) {$app.CommandLineExecutable = " "}
    if ($app.WorkingDirectory -eq $null) {$app.WorkingDirectory = " "}
    #if ($app.Description -lt 2) {$app.Description = " "}
    If($app.ClientFolder -eq $null) {$app.ClientFolder = " "}
        #Adding Application
        write-host "BrowserName" $app.BrowserName
        write-host "ComdLineExe" $app.CommandLineExecutable
        write-host "Description" $app.Description
        write-host "ComdLineArg" $app.CommandLineArguments
        write-host "Enabled    " $app.Enabled
        write-host "Name       " $app.Name
        write-host "UserFiltEna" $app.UserFilterEnabled
        write-host "WorkingDire" $app.WorkingDirectory
        write-host "Published  " $app.PublishedName
        write-host "AdminFolderName"  $app.AdminFolderName
        write-host "Importando" $app.ApplicationName
         
 
        #If($app.ClientFolder -ne $null)
        #{
        
        #  | Out-Null
        # -AdminFolder $app.AdminFolderName
        # -Description $app.Description lo dejamos vacio por lo que no se pone
        New-BrokerApplication -AdminFolder $app.AdminFolderName -BrowserName $app.BrowserName -CommandLineExecutable $app.CommandLineExecutable -WorkingDirectory $app.WorkingDirectory -DesktopGroup $dg  -Name $app.ApplicationName -PublishedName $app.PublishedName -ClientFolder $app.ClientFolder -Enabled $app.Enabled -UserFilterEnabled $app.UserFilterEnabled -CommandLineArguments $app.CommandLineArguments -ShortcutAddedToStartMenu $true | Out-Null
        #}
        #else{
        
        #  | Out-Null
        # -AdminFolder $app.AdminFolderName
        #New-BrokerApplication -AdminFolder $app.AdminFolderName -BrowserName $app.BrowserName -CommandLineExecutable $app.CommandLineExecutable -WorkingDirectory $app.WorkingDirectory -DesktopGroup $dg  -Name $app.ApplicationName -PublishedName $app.PublishedName -ClientFolder "\" -Enabled $app.Enabled -UserFilterEnabled $app.UserFilterEnabled -CommandLineArguments $app.CommandLineArguments -Description $app.Description | Out-Null
        #}


        
        # Es necesario hacerlo con la ruta de la aplicación ($app.Name), si se van a ubicar en raiz habrá que sustituir por $app.ApplicationName

        #Setting applications icon  
        Set-BrokerApplication $app.Name -IconUid $IconUid.Uid
        #Write-Output "Revisando el icono" $IconUid.Uid 
    }
    Catch
    { 

        LogWrite "Aplicación:" 
        LogWrite $app.ApplicationName
        LogWrite "Errores"
        LogWrite $_.Exception.Message  
        LogWrite $_.Exception.ItemName 
        LogWrite "BrowseName:"
        LogWrite $app.BrowserName
        LogWrite "PublishedName:"
        LogWrite $app.PublishedName
        LogWrite "ClientFolder:"
        LogWrite $app.ClientFolder 
        LogWrite "Ejecutable:"
        LogWrite $app.CommandLineExecutable
        LogWrite "Descripcion:" 
        LogWrite $app.Description
        LogWrite "Argumentos:" 
        LogWrite $app.CommandLineArguments
        LogWrite "Habilitada"
        LogWrite $app.Enabled
        LogWrite "Ruta Origen:" 
        LogWrite $app.Name
        LogWrite "Filtro Usuarios:" 
        LogWrite $app.UserFilterEnabled
        LogWrite "Directorio de trabajo"
        LogWrite $app.WorkingDirectory
        LogWrite "Nombre de Publicación" 
        LogWrite $app.PublishedName
        LogWrite "############################################################"
        write-host "error en" $app.ApplicationName "mirar log" -ForegroundColor Red
    }
 
  # Asociacion de Usaurios
 If($app.AssociatedUserNames -ne $null) 
 {
 Try
 {
    $users = $app.AssociatedUserNames
 
    foreach($user in $users)
    {   
        # Es necesario hacerlo con la ruta de la aplicación ($app.Name), si se van a ubicar en raiz habrá que sustituir por $app.ApplicationName
        #"Añadiendo Permisos..." $user $app.ApplicationName 
        Add-BrokerUser -Name "$user" -Application $app.Name


              }
       }
    Catch
    {
       LogWrite "Añadiendo Permisos"
       LogWrite $app.ApplicationName
       LogWrite $user
       LogWrite $_.Exception.Message 
       LogWrite $_.Exception.ItemName       
       LogWrite "############################################################" 
 }
}
}

 