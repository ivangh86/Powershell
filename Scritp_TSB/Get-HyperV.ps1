[CmdletBinding()]
Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$computerName
	
   #[Parameter(Mandatory=$False)]
   #[string]$filePath
)

Write-host -ForegroundColor DarkGreen "Revisando HyperV para la maquina $computerName"
Invoke-Command -Computername $computerName -FilePath "\\serverhqr1\soft_ctx\Team\Ivan\VM\Get-VM.ps1"

pause