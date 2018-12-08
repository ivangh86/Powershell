# Script para automatizacion de la creacion de accesos directos para XA
# version 2.0
# Modificaciones de la version: No borrar lnk sino se crea, incluir el parametro de la aplicacion, generar fichero de log


# Definicion de Variables
# Elegir la ruta donde se almacenaran los shortcuts y los logs. Importante acabar la ruta con un \. Ej: c:\shortcuts\
$RutaShortCuts = "C:\shortcuts\"
$RutaLog = "C:\shortcuts\"


# Cargamos cmdlets Citrix
asnp citrix*
$apps = Get-BrokerApplication | select BrowserName

#Borramos Shortcuts temporales anteriores
$DeleteTemp = $RutaShortCuts + "*.tmp"
remove-item $DeleteTemp

# Creamos el fichero de log
# Datos almacenados: Nombre de aplicacion,Nombre Icono, Ruta ejecutable, Argumentos, Working Directory, Creacion correcta / incorrecta
$LogTime = Get-Date -Format "yyyy_MM_dd_hh_mm"
$FicheroLog = $RutaLog + "LogShortcut_" + $LogTime + ".csv"

foreach ($app in $apps)
{
$in = Get-BrokerApplication -BrowserName $app.BrowserName
if ($in.Description -Like "*prefer*") {
                
                #Listamos aplicaciones

               Write-Host “-----------------------------------------------------------------------------------"
               Write-Host “Ruta Ejecutable: ” $in.CommandLineExecutable
                Write-Host “Argumentos: ” $in.CommandLineArguments
               Write-Host “Ruta Trabajo: ” $in.WorkingDirectory
               Write-Host “Keywords: ” $in.Description
               $IconName=$in.Description.split("=")[-1]
               Write-Host “Nombre Acceso Directo: ” $IconName
               Write-Host “-----------------------------------------------------------------------------------"
                
                # Creacion del Shortcut
                
                $WScriptShell = New-Object -ComObject WScript.Shell
                $ShortcutFile = $RutaShortCuts + $IconName + ".lnk"
                $ShortcutFile_Temp = $RutaShortCuts + $IconName + ".lnk.tmp"
               # Renombramos el anterior si existe
                If (Test-Path $ShortcutFile){
                                Rename-Item $ShortcutFile $ShortcutFile_Temp
                                }
                $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
               $Shortcut.TargetPath = $in.CommandLineExecutable
                # Añadimos argumentos si fuera necesario
                if ($in.CommandLineArguments) { 
                                $Shortcut.Arguments = $in.CommandLineArguments
                               }
                $Shortcut.Save()
                
                # Comprobamos la correcta creacion sino restauramos el antiguo
                If (Test-Path $ShortcutFile){
                                # Registramos OK
                                $LogApp = $in.ApplicationName + "," + $IconName + "," + $in.CommandLineExecutable + "," + $in.CommandLineArguments + "," + $in.WorkingDirectory + ",CORRECTO"
                                # Borramos el temporal si existe
                               If (Test-Path $ShortcutFile_Temp){
                                                Remove-Item $ShortcutFile_Temp
                                                }
                                }
                                # Restablecemos el Anterior si no se ha creado el nuevo
                                else{
                                                #Registramos Error
                                                $LogApp = $in.ApplicationName + "," + $IconName + "," + $in.CommandLineExecutable + "," + $in.CommandLineArguments + "," + $in.WorkingDirectory + ",ERROR"
                                                Rename-Item $ShortcutFile_Temp $ShortcutFile
                                                }
                                
               }

                # Generamos Log de app
                $LogApp  | Out-File $FicheroLog -Append -Force 
}
