##################### FUNCIONES ########################################

function MostrarMenu
{

    param (
        [string]
        $Titulo = "Menú"
    )
    cls
    Write-Host "================ $Titulo ================"
    Write-Host "1: Presiona 1 para sesiones"
    Write-Host "2: Presiona 2 para servidores"
    Write-Host "3: Presiona Q para salir"
}

function MostrarMenuSesiones
{

    param (
        [string]
        $TituloSesiones = "Menú Sesiones"
    )
    cls
    Write-Host "================ $TituloSesiones ================"
    Write-Host "1: Presiona 1 Para buscar un usuario"
    Write-Host "2: Presiona 2 para mostrar las sesiones desconnectadas"
    Write-Host "3: Presiona Q para salir"
}

function MostrarMenuXenapp
{

    param (
        [string]
        $TituloSesiones = "Menú Xenapp"
    )
    cls
    Write-Host "================ $TituloXenapp ================"
    Write-Host "1: Presiona 1 ara WorkGroup  de un Xenapp"
    Write-Host "2: Presiona 2 para mostrar el evaluador de carga de un Xenapp"
    Write-Host "3: Presiona Q para salir"
}
##################### FIN FUNCIONES ########################################
do
{
    MostrarMenu
    $entrada = Read-Host " Selecciona una opción"

        switch ($entrada)
        {

         "1"
         {
            do
            {
                MostrarMenuSesiones
                $entradaSesiones = Read-Host " Selecciona una opción"
                Switch ($entradaSesiones)
                {
                    "1"
                    {
                       $usuario = Read-Host "introduce el nombre del usuario"
                       cls
                       Get-XASession |where {$_.accountname -like "*$usuario"} | Select-Object AccountName, ServerName, ClientName, State | FL
                    }

                    "2"
                    {
                        Get-XASession |Select ServerName, State, AccountName, BrowserName  |Where {$_.State -eq "Disconnected"} |Sort ServerName | FL   
                    }
            
                    "q"
                    {
                        return
                    }
                }pause
            }until ($entradaSesiones -eq "q")
         }
     
         "2"
         {
            do
            {
                MostrarMenuXenapp
                $entradaXenapp = Read-Host " Selecciona una opción"
                Switch ($entradaXenapp)
                {
                    "1"
                    {
                       $Xenapp = Read-Host "Introduce el nombre del Xenapp"
                        Get-XAWorkerGroup -ServerName "$Xenapp"   
                    }
                    "2"
                    {
                        $Xenapp = Read-Host "Introduce el nombre del Xenapp"
                        Get-XALoadEvaluator -ServerName "$Xenapp" | Select-Object LoadEvaluatorName
                    }
                    "q"
                    {
                        return
                    }
                }pause

            }until ($entradaXenapp -eq "q")
            
     
         }
     
         "q"
         {
            return
         }
        }
        pause
} until ($entrada -eq "q")

# 
