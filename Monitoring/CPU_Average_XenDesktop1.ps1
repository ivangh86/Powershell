###############################################################################################
## XenAppServerMonitor
## Ivan Gomez Hernandez
## DXC Citrix
## v1.0
################################################################################################
Param
(
[string]$serversilo,
[switch]$sendmail,
[switch]$serverlist
)


#Revisamos si existen los modulos de Powershel

 if ((Get-PSSnapin "Citrix.Configuration.Admin.V2" -EA silentlycontinue) -eq $null) {
	try { Add-PSSnapin Citrix* -ErrorAction Stop }
	catch { write-error "Error loading XenApp Powershell snapin"; Return }
}


# Funcion revision procesos CPU mas consumo
Function Get-iLPHighCPUProcess {
<#

.SYNOPSIS
    Retrieve processes that are utilizing the CPU on local or remote systems.
 
.DESCRIPTION
    Uses WMI to retrieve process information from remote or local machines. You can specify to return X number of the top CPU consuming processes
    or to return all processes using more than a certain percentage of the CPU.
 
.EXAMPLE
    Get-HighCPUProcess
 
    Returns the 3 highest CPU consuming processes on the local system.
 
.EXAMPLE
    Get-HighCPUProcess -Count 1 -Computername AppServer01
 
    Returns the 1 highest CPU consuming processes on the remote system AppServer01.
 
 
.EXAMPLE
    Get-HighCPUProcess -MinPercent 15 -Computername "WebServer15","WebServer16"
 
    Returns all processes that are consuming more that 15% of the CPU process on the hosts webserver15 and webserver160
#>
 
[Cmdletbinding(DefaultParameterSetName="ByCount")]
Param(
    [Parameter(ValueFromPipelineByPropertyName)]
    [Alias("PSComputername")]
    [string[]]$Computername = "localhost",
     
    [Parameter(ParameterSetName="ByPercent")]
    [ValidateRange(1,100)]
    [int]$MinPercent,
 
    [Parameter(ParameterSetName="ByCount")]
    [int]$Count = 1
)
 
 
Process {
    Foreach ($Computer in $Computername){
    
        Write-Verbose "Retrieving processes from $Computer"
        $wmiProcs = Get-WmiObject Win32_PerfFormattedData_PerfProc_Process -Filter "idProcess != 0" -ComputerName $Computer
     
        if ($PSCmdlet.ParameterSetName -eq "ByCount") {
            $wmiObjects = $wmiProcs | Sort PercentProcessorTime -Descending | Select -First $Count 
        } elseif ($psCmdlet.ParameterSetName -eq "ByPercent") {
            $wmiObjects = $wmiProcs | Where {$_.PercentProcessorTime -ge $MinPercent} 
        } #end IF
 
        $wmiObjects | Foreach {
            $outObject = [PSCustomObject]@{
                Computername = $Computer
                ProcessName = $_.name
                Percent = $_.PercentProcessorTime
                ID = $_.idProcess
            }
            $outObject
            
        } #End ForeachObject
    } #End ForeachComputer
}
 
}



function Ping ([string]$hostname, [int]$timeout = 1000, [int]$retries = 1) {
	$result = $true
	$ping = new-object System.Net.NetworkInformation.Ping #creates a ping object
	$i = 0
	do {
		$i++
		# write-host "Count: $i - Retries:$retries"
		
		try {
			# write-host "ping"
			$result = $ping.send($hostname, $timeout).Status.ToString()
		} catch {
			# Write-Host "error"
			continue
		}
		if ($result -eq "success") { return $true }
		
	} until ($i -eq $retries)
	return $false
}


# Formato HTML

$headformat = "<style type='text/css'>"
$headformat = $headformat + "BODY{background-color:white;}"
$headformat = $headformat + "TABLE{border-width: 5px;border-style: solid;border-color: black;border-collapse: collapse;}"
$headformat = $headformat + "TH{border-width: 3px;padding: 5px;border-style: solid;border-color: black;text-align: center;background-color:grey;}"
$headformat = $headformat + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;text-align: center;background-color}"
$headformat = $headformat + "</style>"


#$SecureString = ConvertTo-SecureString "ZBzR948NxKOQVkMm" -AsPlainText -Force
#$Credentials = New-Object System.Management.Automation.PSCredential ADGBS\Svc_s_taskpshell, $SecureString

$currentDir = Split-Path $MyInvocation.MyCommand.Path
$ServerListFile = "$currentDir\serverlist.txt"  
$filehtml = "$currentDir\Performance\Report_XenApps.html"
$filetxt = "$currentDir\Performance\Report_XenApps.txt"
$dia = Get-Date -Format dd/MM/yy
$hora = Get-Date -Format HH:mm
#$serversilo = "SDI W2K8R2 OFI IE11 DES"

if ($serverlist) {
Write-Host "Se detecta parametro serverlist, cargamos las maquinas del fichero $ServerListFile" -ForegroundColor DarkYellow
$servers = Get-Content $ServerListFile -ErrorAction SilentlyContinue 

}



elseif ($serversilo) {
Write-Host "Se detecta parametro serversilo, cargamos las maquinas del delivery Group $serversilo"
$servers = Get-BrokerMachine -MaxRecordCount 9999 | Where-Object {($_.DesktopGroupName -like $serversilo) -and ({$_.MachineName -like "XA*" }) } | select MachineName
 
}

else {
Write-Host "No se detecta parametro revisamos todas las maquinas"
$servers = Get-BrokerMachine  -MaxRecordCount 9999 | Where-Object {$_.MachineName -like "XA*" } | select MachineName
}


$date = Get-Date
write-host $date -ForegroundColor Green
#$job = @()

$Serverko = @()
$Serverok = "$currentDir\serverok.txt"
$result = ""
$Computer = ""

$ExisteFichero = Test-Path $Serverok
if ($ExisteFichero -eq $True){del $Serverok}



ForEach($computername in $servers){
 if ($serverlist) {  $Computer =  $computername }

 else { $Computer = $computername.MachineName.Substring(6)}
    
    $tests = @{}
    #Ping server
    
	$result = Ping $Computer 100
    write-host $Computer
    if ($result -ne "SUCCESS") { $tests.Ping = "NOK"  
    write-host "$result if"
    
    if ($tests.Ping = "NOK") {$tests.Ping = "RED"+"$($tests.Ping)"}
    
    
    $objVM = New-Object -TypeName psobject
    $objVM | Add-Member -Name "Ping" -Value $tests.Ping -MemberType NoteProperty
    $objVM | Add-Member -Name "Server" -Value $Computer -MemberType NoteProperty
    
    $Serverko += $objVM
    
    }
    
    else { $tests.Ping = "OK" ;write-host "$result  else"

    Add-Content -value $Computer -Path $Serverok

        }
}

   $date = Get-Date
    write-host $date -ForegroundColor Gray
    $job = @()
    
   $Serveroki = Get-Content  "$Serverok"    
   ForEach($computername in $Serveroki){
          write-host $computername -ForegroundColor DarkYellow
          #parte para HYperV
          #$cpuUtil = invoke-command $computername -Credential $Credentials {Get-Counter -Counter "\Hyper-V Hypervisor Logical Processor(_Total)\% Total Run Time" -MaxSamples 5 | select CounterSamples} -HideComputerName -AsJob
          #$valuecpuUtil = $cpuUtil.CounterSamples.cookedValue | Measure-Object -Average 
          #parte para XenApp 
          $job += invoke-command $computername <#-Credential $Credentials#> {Get-Counter -Counter "\Processor(_total)\% Processor Time" -SampleInterval 1 -MaxSamples 60 | select CounterSamples} -HideComputerName -AsJob    
         }
    #}

        Get-Job | Wait-Job

        $date = Get-Date
        write-host $date -ForegroundColor Magenta

        $resultado = @()

        #$Computer = 

        for ($i=0;$i -lt ($job).Count;$i++){    
            $Computer = $job[$i].location
            $jobresult = receive-Job $job[$i]

            #$tests = @{}
            #$result = Ping $Computer 100
	        #if ($result -ne "SUCCESS") { $tests.Ping = "NOK"|write-host "$result if 2" }    

           
            #write-host "$result  else 2 "
            #$resultado += Receive-Job -Job $job -Location $jobid.Location


        #If ($computer -ne $env:COMPUTERNAME){
        #    $mem = gwmi -Class win32_operatingsystem -computername $computer <#-Credential $Credentials#> | Select-Object @{Name = "MemoryUsage"; Expression = {“{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize) }}
        #    $lastboottime = (Get-WmiObject -Class Win32_OperatingSystem -computername $computer <#-Credential $Credentials #>).LastBootUpTime
        #    $physicaldisks = Get-WmiObject -class Win32_logicaldisk -ComputerName $computer -ErrorAction SilentlyContinue <#-Credential $Credentials #>| Select * | Where-Object {$_.Mediatype -eq "12"} # te saca todos los discos
        #}
        #else{
        #$mem = gwmi -Class win32_operatingsystem -computername $computer | Select-Object @{Name = "MemoryUsage"; Expression = {“{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize) }}
        $valuecpuUtil = $cpuUtil.CounterSamples.cookedValue | Measure-Object -Average
        $lastboottime = (Get-WmiObject -Class Win32_OperatingSystem -computername $Computer).LastBootUpTime
        $physicaldisks = Get-WmiObject -class Win32_logicaldisk -ComputerName $Computer -ErrorAction SilentlyContinue | Select * | Where-Object {$_.Mediatype -eq "12"} # te saca todos los discos



        #Memory Usage
        $mem = ""                 
        $MEMavail = Get-Counter -Computername $Computer -counter "\Memory\Available Mbytes" | Foreach {[math]::round($_.CounterSamples[0].CookedValue, 2)}
        $MEMtotal = ((get-wmiobject -Computername $Computer -class "win32_physicalmemory" -namespace "root\CIMV2").Capacity)
        $sum = $MEMtotal -join '+'
        $MEMsum = Invoke-Expression $sum
        $MEMtotalGB = ($MEMsum / 1048576)
        $MEMload = ($MEMavail * 100 / $MEMtotalGB) 
        $Memloadcounter = (100 - [math]::round($MEMload, 2))
        $mem = "$($Memloadcounter)%"

        #Disk Queue
        $diskcounterq =""
	    $diskcounterq = Get-Counter -Counter "\LogicalDisk(C:)\Avg. disk queue length" -ComputerName $Computer
        $diskq = $diskcounterq.CounterSamples | Select-Object CookedValue
        $diskQcounter = $([math]::round($diskq.Cookedvalue, 2))
            
        #if ($diskqcounter -eq 0) {$diskqcounter = "GREEN"+"$($diskqcounter)"}
        if ($diskqcounter -gt 1) {$diskqcounter = "YELLOW"+"$($diskqcounter)"}
        elseif ($disqcounter -gt 2) {$diskqcounter = "RED"+"$($diskqcounter)"}

        $Proceso = ""
        $Proceso = @() 

        $Proceso += Get-iLPHighCPUProcess -Computername $Computer -Count 1 

       
        
        $drives = @()
        for($j=0; $j -lt ([Array]$physicaldisks).count ; $j++){
        $espacioDiscoTotal = $physicaldisks[$j].Size
        $espacioDiscoLibre = $physicaldisks[$j].FreeSpace
        $porcen = ($espacioDiscoLibre*100) / $espacioDiscoTotal
        $porcen = [math]::Round($porcen,2)
        $unitDC = $physicaldisks[$j].Name
        $porcen = 100-$porcen
        $porcen = [math]::Round($porcen,2)
        $Drives += "$($unitDC)"+"$($porcen)%"
    }
         

        $sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($lastboottime)

        $CPUdata = $jobresult.CounterSamples.cookedValue | Measure-Object -Maximum -Minimum -Average
    
        $CPUAvg = $CPUdata.Average
        $CPUMax = $CPUdata.Maximum
        $CPUMin = $CPUdata.Minimum    
        $CPUMax = "$("{0:N2}" -f ($CPUMax))%"
        $CPUMin = "$("{0:N2}" -f ($CPUMin))%"    

        If ($CPUAvg -gt 75){$CPUAvg = "YELLOW"+"$("{0:N2}" -f ($CPUAvg))%"}
        elseif ($CPUAvg -gt 90){$CPUAvg = "RED"+"$("{0:N2}" -f ($CPUAvg))%"}
        else{$CPUAvg = "$("{0:N2}" -f ($CPUAvg))%"}

        If ($mem -gt 95){$mem = "RED"+"$($mem)"}
        else{$mem = "$($mem)"}
        <#
    If ($CPUAvg.Average -gt 75){$headformat = $headformat + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;text-align: center;background-color:red;}"}
    else{$headformat = $headformat + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;text-align: center;background-color:green;}"}
    #>

        
    $objVM = New-Object -TypeName psobject
    $objVM | Add-Member -Name "Ping" -Value "GREENOK" -MemberType NoteProperty
    $objVM | Add-Member -Name "Server" -Value $Computer -MemberType NoteProperty
    $objVM | Add-Member -Name "CPU Average / 60 sec" -Value $CPUAvg -MemberType NoteProperty
    $objVM | Add-Member -Name "CPU Max" -Value $CPUMax -MemberType NoteProperty
    $objVM | Add-Member -Name "CPU Min" -Value $CPUMin -MemberType NoteProperty
    $objVM | Add-Member -Name "Memory" -Value $mem -MemberType NoteProperty
    $objVM | Add-Member -Name "Unidad C:" -Value $drives[0].Substring(2) -MemberType NoteProperty
    $objVM | Add-Member -Name "Unidad D:" -Value $drives[1].Substring(2) -MemberType NoteProperty
    $objVM | Add-Member -Name "Unidad E:" -Value $drives[2].Substring(2) -MemberType NoteProperty
    $objVM | Add-Member -Name "UpTime" -Value $sysuptime.days -MemberType NoteProperty
    $objVM | Add-Member -Name "DiskQueue" -Value $diskQcounter -MemberType NoteProperty
    $objVM | Add-Member -Name "TopProcess" -Value  $Proceso.ProcessName -MemberType NoteProperty
    $objVM | Add-Member -Name "TopProcessPercent" -Value  $Proceso.Percent -MemberType NoteProperty

 
    #$result

    Write-Host $Computer $CPUAvg

    $resultado += $objVM
    
}



$resultado += $Serverko

$resultado = $resultado | ConvertTo-Html -Head $headformat -Body "<H3>Server Health Report System Status XenApp 7.15 BS - $dia $hora</H3>" |  foreach {
 $PSItem -replace ("<td>GREEN", "<td style='background-color:#5ac031'>")  -replace ("<td>RED", "<td style='background-color:#fc2f48 '>") -replace ("<td>YELLOW", "<td style='background-color:#fc2f48 '>")
} | out-file $filehtml
& $filehtml



$resultado -replace ("GREEN","")
$resultado -replace ("RED","")
$resultado -replace ("YELLOW","")


$resultado | Export-Csv -Delimiter ";" $filetxt
#>

#Send email functionality from below line, use it if you want 


if ($sendmail){
$smtpServer = "relay.adgbs.com"
$smtpFrom = "0901STServers@bancsabadell.com"
$smtpTo = "gomez-hernandez@dxc.com"
#$smtpTo = "giorgio.alf.di-lorenzo@hpe.com"
$messageSubject = "Servers Health report XenApp 7.15 BS - $dia $hora"
$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true
$message.Body = "<head><pre>$style</pre></head>"
$message.Body += Get-Content $filehtml

$att = new-object Net.Mail.Attachment($filetxt)
$message.Attachments.Add($att)
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)
$message.Dispose()
}