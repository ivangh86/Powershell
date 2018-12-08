###############################################################################################
## XenAppServerMonitor
## Ivan Gomez Hernandez
## DXC Citrix
## v1.0
################################################################################################
Param
(
[string]$serversilo,
[switch]$sendmail = $false,
[switch]$serverlist = $false
)


 if ((Get-PSSnapin "Citrix.Configuration.Admin.V2" -EA silentlycontinue) -eq $null) {
	try { Add-PSSnapin Citrix* -ErrorAction Stop }
	catch { write-error "Error loading XenApp Powershell snapin"; Return }
}

function Ping([string]$hostname, [int]$timeout = 1000, [int]$retries = 1) {
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

$serverlist = Get-Content $ServerListFile -ErrorAction SilentlyContinue 

}

elseif ($serversilo) {

$Serverlist = Get-BrokerMachine | Where-Object {$_.DesktopGroupName -like $serversilo } | select MachineName
}

else {

$Serverlist = Get-BrokerMachine | select MachineName
}



$date = Get-Date
write-host $date -ForegroundColor Green
#$job = @()

$Serverko = @()
$Serverok = "$currentDir\serverok.txt"
del $Serverok

ForEach($computername in $Serverlist){

    $tests = @{}
    #Ping server 
	$result = Ping $computername.MachineName.Substring(6) 100
    write-host $computername.MachineName.Substring(6)
    if ($result -ne "SUCCESS") { $tests.Ping = "NOK"  
    write-host "$result if"
    
    if ($tests.Ping = "NOK") {$tests.Ping = "RED"+"$($tests.Ping)"}
    
    
    $objVM = New-Object -TypeName psobject
    $objVM | Add-Member -Name "Ping" -Value $tests.Ping -MemberType NoteProperty
    $objVM | Add-Member -Name "Server" -Value $computername.MachineName.Substring(6) -MemberType NoteProperty
    
    $Serverko += $objVM
    
    }
    
    else { $tests.Ping = "OK" ;write-host "$result  else"

    Add-Content -value $computername.MachineName.Substring(6) -Path $Serverok

        }
}

   $date = Get-Date
    write-host $date -ForegroundColor Gray
    $job = @()
    
   $Serverok = Get-Content  $Serverok    
   ForEach($computername in $Serverok){

          write-host $computername -ForegroundColor DarkYellow
          #parte para HYperV
          #$cpuUtil = invoke-command $computername -Credential $Credentials {Get-Counter -Counter "\Hyper-V Hypervisor Logical Processor(_Total)\% Total Run Time" -MaxSamples 5 | select CounterSamples} -HideComputerName -AsJob
          #$valuecpuUtil = $cpuUtil.CounterSamples.cookedValue | Measure-Object -Average 
          #parte para XenApp 
          $job += invoke-command $computername <#-Credential $Credentials#> {Get-Counter -Counter "\Processor(_total)\% Processor Time" -SampleInterval 1 -MaxSamples 5 | select CounterSamples} -HideComputerName -AsJob    
         }
    #}


    
        Get-Job | Wait-Job

        $date = Get-Date
        write-host $date -ForegroundColor Magenta

        $resultado = @()

        #$computer = 

        for ($i=0;$i -lt ($job).Count;$i++){    
            $Computer = $job[$i].location
            $jobresult = receive-Job $job[$i]

            #$tests = @{}
            #$result = Ping $Computer 100
	        #if ($result -ne "SUCCESS") { $tests.Ping = "NOK"|write-host "$result if 2" }    
            $tests.Ping = "OK"
            if ($tests.Ping = "OK") {$tests.Ping = "GREEN"+"$($tests.Ping)"}
           
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
        $lastboottime = (Get-WmiObject -Class Win32_OperatingSystem -computername $computer).LastBootUpTime
        $physicaldisks = Get-WmiObject -class Win32_logicaldisk -ComputerName $computer -ErrorAction SilentlyContinue | Select * | Where-Object {$_.Mediatype -eq "12"} # te saca todos los discos

        #Memory Usage
        $mem = ""                 
        $MEMavail = Get-Counter -Computername $computer -counter "\Memory\Available Mbytes" | Foreach {[math]::round($_.CounterSamples[0].CookedValue, 2)}
        $MEMtotal = ((get-wmiobject -Computername $computer -class "win32_physicalmemory" -namespace "root\CIMV2").Capacity)
        $sum = $MEMtotal -join '+'
        $MEMsum = Invoke-Expression $sum
        $MEMtotalGB = ($MEMsum / 1048576)
        $MEMload = ($MEMavail * 100 / $MEMtotalGB) 
        $Memloadcounter = (100 - [math]::round($MEMload, 2))
        $mem = "$($Memloadcounter)%"

        #Disk Queue
        $diskcounterq =""
	    $diskcounterq = Get-Counter -Counter "\LogicalDisk(C:)\Avg. disk queue length" -ComputerName $computer
        $diskq = $diskcounterq.CounterSamples | Select-Object CookedValue
        $diskQcounter = $([math]::round($diskq.Cookedvalue, 2))
            
        # ejemplo If ($mem.MemoryUsage -gt 97){$mem = "RED"+"$($mem.MemoryUsage)%"}
        #if ($diskqcounter -eq 0) {$diskqcounter = "GREEN"+"$($diskqcounter)"}
        #elseif ($diskqcounter -gt 1) {$diskqcounter = "RED"+"$($diskqcounter)"}
        #elseif ($disqcounter -gt 2) {$diskqcounter = "RED"+"$($diskqcounter)"}


        }
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

        If ($CPUAvg -gt 75){$CPUAvg = "RED"+"$("{0:N2}" -f ($CPUAvg))%"}
        else{$CPUAvg = "$("{0:N2}" -f ($CPUAvg))%"}
        If ($mem -gt 97){$mem = "RED"+"$($mem)"}
        else{$mem = "$($mem)"}
        <#
    If ($CPUAvg.Average -gt 75){$headformat = $headformat + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;text-align: center;background-color:red;}"}
    else{$headformat = $headformat + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;text-align: center;background-color:green;}"}
    #>

        
    $objVM = New-Object -TypeName psobject
    $objVM | Add-Member -Name "Ping" -Value "OK" -MemberType NoteProperty
    $objVM | Add-Member -Name "Server" -Value $computer -MemberType NoteProperty
    $objVM | Add-Member -Name "CPU Average / 60 sec" -Value $CPUAvg -MemberType NoteProperty
    $objVM | Add-Member -Name "CPU Max" -Value $CPUMax -MemberType NoteProperty
    $objVM | Add-Member -Name "CPU Min" -Value $CPUMin -MemberType NoteProperty
    $objVM | Add-Member -Name "Memory" -Value $mem -MemberType NoteProperty
    $objVM | Add-Member -Name "Drives" -Value $drives -MemberType NoteProperty
    $objVM | Add-Member -Name "UpTime" -Value $sysuptime.days -MemberType NoteProperty
    $objVM | Add-Member -Name "DiskQueue" -Value $diskQcounter -MemberType NoteProperty

 
    #$result

    Write-Host $computer $CPUAvg

    $resultado += $objVM
    

}



$resultado += $Serverko

$resultado = $resultado | ConvertTo-Html -Head $headformat -Body "<H3>Server Health Report System Status XenApp 7.15 BS</H3>" |  foreach {
 $PSItem -replace ("<td>GREEN", "<td style='background-color:#5ac031'>")  -replace ("<td>RED", "<td style='background-color:#fc2f48 '>")
} | out-file $filehtml
& $filehtml



$resultado -replace ("GREEN","")
$resultado -replace ("RED","")


$resultado | Export-Csv -Delimiter ";" $filetxt
#>

##Send email functionality from below line, use it if you want 


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