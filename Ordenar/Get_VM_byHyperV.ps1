Write-Host "Importamos modulo VMM" -ForegroundColor Green

if ((Get-Module "virtualmachinemanager" -EA silentlycontinue) -eq $null) {
	try { Import-Module "\\VMMBSV02P\C$\Program Files\Microsoft System Center 2012 R2\Virtual Machine Manager\bin\psModules\virtualmachinemanager\virtualmachinemanager.psd1" -ErrorAction Stop }
	catch { write-error "Error loading XenApp Powershell snapin"; Return }
}



$VMMBS= "VMMBSV02P.adgbs.com"
$VMMSF= "VMMISV02P.adgbs.com"
$HyperV = "hypisf331p"

Get-VMMServer $VMMSF | Get-VMHost -ComputerName $HyperV |Get-VM |Select-Object Name