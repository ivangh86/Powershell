################################################################################################
if ((Get-Module "virtualmachinemanager" -EA silentlycontinue) -eq $null) {
	try { Import-Module "\\VMMBSV02P\C$\Program Files\Microsoft System Center 2012 R2\Virtual Machine Manager\bin\psModules\virtualmachinemanager\virtualmachinemanager.psd1" -ErrorAction Stop }
	catch { write-error "Error loading XenApp Powershell snapin"; Return }
}



#Write-Host "Importamos modulo VMM" -ForegroundColor Green


$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$VMMBS= "VMMBSV02P.adgbs.com"
$VMMSF= "VMMISV02P.adgbs.com"
#$computerName = Read-Host "introduce el nombre del Servidor"
$computers = get-content "$scriptDir\serverlist.txt"

#Write-Host "Ejecutamos la consulta para $computerName" -ForegroundColor Green
#write-host "Server,State,Host"




$ServidoresContenido = @()
foreach ($computerName in $computers) {
    

If (($computerName -like "*BSV*") -or ( $computerName -like "*SCV*"))  {

$result = Get-VMMServer $VMMBS | get-vm $computerName 
        #write-host "Maquina de BS"
        If ($result -eq $null) {

        write-host  "$computerName no encontrada en $VMMBS"
        }
        
        else {    
        
        $VM = "" |select-object  name,state,host
        $VM.name  = $result.name
        $VM.state = $result.VirtualMachineState
        $VM.host  =$result.hostname

        #write-host "$VMname,$VMstate,$VMhost"

        #$result |Select-Object vmhost,VirtualMachineState,ComputerName | format-table -AutoSize -HideTableHeaders }
        }

       }

else {

    $VMM = $VMMSF
    $result = Get-VMMServer $VMMSF | get-vm $computerName 
        #write-host "Maquina de SF"
        If ($result -eq $null) {

        write-host  "$computerName no encontrada en $VMMSF"
        }

        else { 
        
        $VM = "" |select-object  name,state,host
        $VM.name  = $result.name
        $VM.state = $result.VirtualMachineState
        $VM.host  =$result.hostname

        #write-host "$VMname,$VMstate,$VMhost"
        #$result |Select-Object vmhost,VirtualMachineState,ComputerName | format-table -AutoSize -HideTableHeaders}
        }

    }

$ServidoresContenido += $VM |Select-Object name,state,host

}



Write-Host "Ejecutamos la consulta para $computerName" -ForegroundColor Green



$ServidoresContenido |Format-Table -AutoSize

Pause