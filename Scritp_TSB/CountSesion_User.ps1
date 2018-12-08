###############################################################################
#
#
# Autor: Ivan Gomez
# Fecha: 20/04/2018
# Descripcion: Revisa las sesiones totales y a que usuarios corresponden.
#
###############################################################################



#Production\Others\3270 TSB FirstData PRO

function Count_SessionUser_App {



[CmdletBinding()]
Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$ApplicationName
	
   #[Parameter(Mandatory=$True)]
   #[string]$filePath
)




#Usuarios Conectados
write-host "Usuarios Conectados"

(Get-BrokerSession -ApplicationInUse "$ApplicationName" |Select-Object UserName ).count

#Sesiones Conectadas
write-host "Sesiones Conectadas"


#$Instances = Get-BrokerApplicationInstance -ApplicationName "*3270 TSB FirstData PRO*" | Select-Object Instances | ft -HideTableHeaders -AutoSize
Get-BrokerApplicationInstance -ApplicationName "$ApplicationName" | Select-Object Instances | ft -HideTableHeaders -AutoSize > \\serverhqr1\soft_ctx\Team\Ivan\Instances.txt

Get-Content \\serverhqr1\soft_ctx\Team\Ivan\Instances.txt | ForEach-Object -Begin {[int]$Sum = 0} -Process {$Sum += $_} -End {$Sum}


}

Count_SessionUser_App 