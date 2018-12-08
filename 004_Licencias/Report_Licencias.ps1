<#
.Funcion
Informes sobre licencias de Citrix en uso para productos seleccionados.
 
.DESCRIPCION
Este script consultará a su servidor de licencias Citrix y
Uso y licencias totales para productos individuales.
 
.NOTAS
Requiere Citrix Licensing Snapin (Citrix.Licensing.Admin
Si el servidor de licencias tiene un certificado autofirmado, puede obtener una licencia
Errores al ejecutar esto. He resuelto esto en mis entornos de prueba
Mediante la instalación del cert como un certificado de CA raíz de confianza.
 

Author: Ivan Gomez
Version 1
Fecha   06/07/2017
 
 
#>
############################################################################################
#Definir las siguientes variables para el entorno a consultar
 
#Insertar la URL del Servidor de licencias a consultar
$ServerAddress = "https://WCPIF105:8083"
 
#Insertar los diferentes tipos de licencia a consultar entre comillas simples
#'Citrix XenApp Enterprise','Citrix XenDesktop Enterprise','Citrix EdgeSight for XenApp'
$LicenseTypeOutput = 'Citrix XenApp Enterprise','Citrix XenDesktop Enterprise','Citrix Start-up License','Citrix License Server Diagnostics License','Citrix StorageLink Enterprise','Citrix Provisioning Services','Citrix Provisioning Server for Desktops'
 
############################################################################################
 
#Revisamos si el modulo de licencias para powershell se encuentra disponible
#Dicho modulo se encuentra en el ejecutable: LicensingAdmin_PowerShellSnapIn_x64.msi
if ( (Get-PSSnapin -Name Citrix.Licensing.Admin.V1 -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PsSnapin Citrix.Licensing.Admin.V1
}
 
 
#Crear Tabla
$Cert = (Get-LicCertificate -AdminAddress $ServerAddress).CertHash
$LicAll = @{}
$LicInUse = @{}
$LicAvailable = @{}
 
#Generacion Tabla visualización de licencias. 
$LicAll = Get-LicInventory -AdminAddress $ServerAddress -certhash "$Cert"
foreach ($LicInfo in $LicAll) 
{
    $Prod = $LicInfo.LocalizedLicenseProductName
    $InUse = $LicInfo.licensesinuse
    $Avail = $LicInfo.LicensesAvailable
    if ($LicInUse.ContainsKey($Prod))
        {
                if ($LicInUse.Get_Item($Prod) -le $InUse) 
                {
                    $LicInUse.Set_Item($Prod, $InUse)
                }
        }
    else
        {
            $LicInUse.add($Prod, $InUse)
        }
    if ($LicAvailable.ContainsKey($Prod))
        {
                if ($LicAvailable.Get_Item($Prod) -le $Avail) 
                {
                    $LicAvailable.Set_Item($Prod, $Avail)
                }
        }
    else
        {
            $LicAvailable.add($Prod, $Avail)
        }
}
 
# Uso de licencia de salida para cada tipo solicitado
Foreach ($Type in $LicenseTypeOutput)
{
    $OutPutLicInUse = $LicInUse.Get_Item($Type)
    $OutPutAvail = $LicAvailable.Get_Item($Type)
    Write-Host "En uso" $OutPutLicInUse  $Type  "de" $OutPutAvail "disponibles."
}