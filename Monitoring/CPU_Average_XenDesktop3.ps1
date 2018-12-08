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

$dia = Get-Date -Format dd/MM/yy
$hora = Get-Date -Format HH:mm
#$serversilo = "SDI W2K8R2 OFI IE11 DES"

if ($serverlist) {
Write-Host "Se detecta parametro serverlist, cargamos las maquinas del fichero $ServerListFile" -ForegroundColor DarkYellow
$servers = Get-Content $ServerListFile -ErrorAction SilentlyContinue 

}

elseif ($serversilo) {
Write-Host "Se detecta parametro serversilo, cargamos las maquinas del delivery Group $serversilo" -ForegroundColor DarkYellow
$servers = Get-BrokerMachine -MaxRecordCount 9999 | Where-Object {($_.DesktopGroupName -like $serversilo) -and ({$_.MachineName -like "XA*" }) } | select MachineName
 
}

else {
Write-Host "No se detecta parametro revisamos todas las maquinas" -ForegroundColor DarkYellow
$servers = Get-BrokerMachine  -MaxRecordCount 9999 | Where-Object {$_.MachineName -like "ADGBS\XA*" } | select MachineName
}


$date = Get-Date
write-host $date -ForegroundColor Green
#$job = @()

$Serverko = @()
$Serverok = "$currentDir\serverok.txt"
$ServerPerfect = "$currentDir\serverperfect.txt"
$result = ""
$Computer = ""

$ExisteFichero = Test-Path $Serverok
if ($ExisteFichero -eq $True){del $Serverok}

$ExisteFichero = Test-Path $ServerPerfect
if ($ExisteFichero -eq $True){del $ServerPerfect}


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
    
    else { 
    
         $tests.Ping = "OK" ;write-host "$result  else"

         $queryinvoke = invoke-command $Computer -ScriptBlock  {Test-Path -Path "C:\Windows"}
         if ($queryinvoke -eq $True){
         Add-Content -value $Computer -Path $Serverok
         }

         else {
            $objVM = New-Object -TypeName psobject
            $objVM | Add-Member -Name "Ping" -Value "REDErrorInvoke" -MemberType NoteProperty
            $objVM | Add-Member -Name "Server" -Value $Computer -MemberType NoteProperty
            write-host "No funciona el InvokeCommand en $Computer"
            $Serverko += $objVM
         
         }


        }
}

   $date = Get-Date
   write-host $date -ForegroundColor Gray
   $job = @()   
   $Serveroki = Get-Content  $Serverok
   

 ForEach($Computer in $Serveroki){

     try {
        
        # Here we specify the WMI query
        $Searcher = [WmiSearcher]'SELECT * FROM Win32_ComputerSystem'
        
        # This is where we set the timeout to 5 seconds.
        $Searcher.Options.TimeOut = "0:0:5"
        
        # Create necessary objects, specify a scope and namespace for the remote computer (or localhost or '.'),
        # then set the scope on the WMISearcher object, which is usually "\\computername\root\cimv2".
        $ConnectionOptions = New-Object Management.ConnectionOptions
        $ManagementScope = New-Object Management.ManagementScope("\\${Computer}\root\cimv2", $ConnectionOptions)
        $Searcher.Scope = $ManagementScope
        
        # Perform the query and retrieve the desired property, and expand it to a
        # string with -ExpandProperty.
        $MemSize = $Searcher.Get() | Select-Object -ExpandProperty TotalPhysicalMemory 
        
        # Use this check for if there's a non-terminating error.
        if ($?) {
            
            $MemSizeGB = ($MemSize/1GB).ToString('N')
            
            Write-Output "${computer}: $MemSizeGB" "WMI funciona correctamente" 
            
        }
        
        # If a non-terminating error occurs, we get here, but I think all errors are terminating with $searcher.Get() ...
        else {
            
            Write-Output "${Computer}: Error: $($Error[0])"
            
        }
        
    }
    
    # If WMI has a terminating error, we get here.
    catch {
        
        Write-Output "${Computer}: Error (terminating): $($Error[0])"
    
    }

    if ($MemSizeGB  -eq $null){

    $objVM = New-Object -TypeName psobject
    $objVM | Add-Member -Name "Estado" -Value "ProblemaWMI" -MemberType NoteProperty
    $objVM | Add-Member -Name "Server" -Value $Computer -MemberType NoteProperty
    
    $Serverwmi += $objVM
    }

    else {Add-Content -value $Computer -Path $ServerPerfect}


}


    $ServerPerfect = Get-Content $ServerPerfect
   
   ForEach($computername in $ServerPerfect){
          write-host $computername -ForegroundColor DarkYellow
          #parte para HYperV
          #$cpuUtil = invoke-command $computername -Credential $Credentials {Get-Counter -Counter "\Hyper-V Hypervisor Logical Processor(_Total)\% Total Run Time" -MaxSamples 5 | select CounterSamples} -HideComputerName -AsJob
          #$valuecpuUtil = $cpuUtil.CounterSamples.cookedValue | Measure-Object -Average 
          #parte para XenApp 
          $job += invoke-command $computername  <#-Credential $Credentials#> { 

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
        $wmiProcs = Get-WmiObject Win32_PerfFormattedData_PerfProc_Process -Filter "idProcess != 0"  -ComputerName $Computer | Where-Object {(($_.name -notlike "WmiPrvSE*") -and ($_.name -notlike "svchost*")) }
     
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

            $TopProcess = Get-iLPHighCPUProcess -Count 3  #| Select-Object ProcessName |Format-Table  -HideTableHeaders
            #JOB 0
            $Procesos = $TopProcess.ProcessName[0]+"--"+$TopProcess.ProcessName[1]+"--"+$TopProcess.ProcessName[2]
                     
            $Procesos
            #JOB 1
            [string]$0 = $TopProcess.Percent[0]
            [string]$1 = $TopProcess.Percent[1]
            [string]$2 = $TopProcess.Percent[2]
            [string]$Porcentage = $0+"--"+$1+"--"+$2
            $Porcentage


            $MaxSamples = 10   
            $res = Get-Counter -Counter "\Processor(_total)\% Processor Time" -SampleInterval 1 -MaxSamples $MaxSamples  | select CounterSamples
                    
            $CPUdata = $res.CounterSamples.cookedValue | Measure-Object -Maximum -Minimum -Average 
             
            $CPUAvg = $CPUdata.Average
          
            $CPUMax = $CPUdata.Maximum
          
            $CPUMin = $CPUdata.Minimum
          
            $CPUAvg = "$("{0:N2}" -f ($CPUAvg))%"    
            $CPUMax = "$("{0:N2}" -f ($CPUMax))%"
            $CPUMin = "$("{0:N2}" -f ($CPUMin))%"
            #JOB 2  
            $CPUAvg
            #JOB 3
            $CPUMax
            #JOB 4
            $CPUMin 

  
              $MEMavail = Get-Counter -counter "\Memory\Available Mbytes" -SampleInterval 1 -MaxSamples  $MaxSamples | select CounterSamples
              $MEMdata = $MEMavail.CounterSamples.cookedValue |  Measure-Object -Maximum -Minimum -Average    
              $MEMAvg = $MEMdata.Average       
              $MEMMax = $MEMdata.Maximum
              $MEMMin = $MEMdata.Minimum
              
 
              #$mem = $mem | Foreach {[math]::round($_.CounterSamples[0].CookedValue, 2)}  
              $MEMtotal = ((get-wmiobject -class "win32_physicalmemory" -namespace "root\CIMV2").Capacity)                     
              #$MEMavail = Get-Counter -counter "\Memory\Available Mbytes" | Foreach {[math]::round($_.CounterSamples[0].CookedValue, 2)}       
              $sum = $MEMtotal -join '+'
              $MEMsum = Invoke-Expression $sum
              $MEMtotalGB = ($MEMsum / 1048576)

              #JOB 5
              $MEMload = ""
              $Memloadcounter = ""
              $MEMload = ($MEMAvg * 100 / $MEMtotalGB) 
              $Memloadcounter = (100 - [math]::round($MEMload, 2))
              $memav = "$($Memloadcounter)%"
              #$MEMAvg
              $memav


              #JOB 6
              $MEMload = ""
              $Memloadcounter = ""
              $MEMload = ($MEMMax * 100 / $MEMtotalGB) 
              $Memloadcounter = (100 - [math]::round($MEMload, 2))
              $memmx = "$($Memloadcounter)%"
              #$MEMMax
              $memmx
              
              #JOB 7
              $MEMload = ""
              $Memloadcounter = ""
              $MEMload = ($MEMMin * 100 / $MEMtotalGB) 
              $Memloadcounter = (100 - [math]::round($MEMload, 2))
              $memmi = "$($Memloadcounter)%"
              #$MEMMin
              $memmi

              
              $physicaldisks = Get-WmiObject -class Win32_logicaldisk  -ErrorAction SilentlyContinue | Select * | Where-Object {$_.Mediatype -eq "12"}

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
              #JOB 8
              if ($drives[0].Substring(2) -gt 97.00) {$drives[0] = "RED"+"$($drives[0])"}
              elseif ($drives[0].Substring(2) -gt 90.00) {$drives[0] = "YELLOW"+"$($drives[0])"}
              $drives[0]

              #JOB 9
              if ($drives[1].Substring(2) -gt 97.00) {$drives[1] = "RED"+"$($drives[1])"}
              elseif ($drives[1].Substring(2) -gt 90.00) {$drives[1] = "YELLOW"+"$($drives[1])"}
              $drives[1]

              #JOB 10
              if ($drives[2].Substring(2) -gt 97.00) {$drives[2] = "RED"+"$($drives[2])"}
              elseif ($drives[2].Substring(2) -gt 90.00) {$drives[2] = "YELLOW"+"$($drives[2])"}
              $drives[2]


              ##$diskcounterq =""
	          ##$diskcounterq = Get-Counter -Counter "\LogicalDisk(C:)\Avg. disk queue length" -SampleInterval 1 -MaxSamples $MaxSamples  | select CounterSamples
              ##$diskq = $diskcounterq.CounterSamples.CookedValue |  Measure-Object -Maximum -Minimum -Average 
              ##$diskQcounter = #$([math]::round($diskq.Cookedvalue, 2))  
              ##$diskAvg = $diskq.Average           
              ##$diskAvg = $([math]::round($diskAvg, 2))
              ##$diskMax = $diskq.Maximum
              ##$diskMax = $([math]::round($diskAvg, 2))
              ##$diskMin = $diskq.Minimum
              ##$diskMin = $([math]::round($diskAvg, 2))

              ##if ($diskAvg -gt 1) {$diskAvg = "YELLOW"+"$($diskAvg)"}
              ##elseif ($diskAvg -gt 2) {$diskAvg = "RED"+"$($diskAvg)"}
              #JOB 11
              ##$diskAvg
              #$diskMax
              #$diskMin 
         
              #JOB 12
              ##$lastboottime = (Get-WmiObject -Class Win32_OperatingSystem  <#-Credential $Credentials #>).LastBootUpTime
              ##$sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($lastboottime)
              ##$sysuptime.days
              } -HideComputerName -AsJob         
         
         }
      
      #$MaxSamples = 10   
      $TimeOut = 1000 
        Get-Job
        Get-Job | Wait-Job -Timeout $TimeOut

        $date = Get-Date
        write-host $date -ForegroundColor Magenta

        $resultado = @()

        #$Computer = 

        for ($i=0;$i -lt ($job).Count;$i++){    
            $Computer = $job[$i].location
            $Computer
            $jobresult = receive-Job $job[$i]
            write-host "este es el jobresult" $jobresult
            write-host "TopProcess:"$jobresult[0]
            write-host "TopProcessPercent:"$jobresult[1]
            write-host "CPUAvg:"$jobresult[2]
            write-host "CPUMax:"$jobresult[3]
            write-host "CPUMin:"$jobresult[4]
            write-host "MEMAvg:"$jobresult[5]
            write-host "Porcentage Ocupado C:"$jobresult[6]
            write-host "Porcentage Ocupado D:"$jobresult[7]
            write-host "Porcentage Ocupado E:"$jobresult[8]
            write-host "DiskQueue Average:"$jobresult[9]
            write-host "Uptime:"$jobresult[10]
        
            $objVM = New-Object -TypeName psobject
            $objVM | Add-Member -Name "Ping" -Value "GREENOK" -MemberType NoteProperty
            $objVM | Add-Member -Name "Server" -Value $Computer -MemberType NoteProperty
            $objVM | Add-Member -Name "CPU Average / 10 sec" -Value $jobresult[2] -MemberType NoteProperty
            $objVM | Add-Member -Name "CPU Max" -Value $jobresult[3] -MemberType NoteProperty
            $objVM | Add-Member -Name "CPU Min" -Value $jobresult[4] -MemberType NoteProperty
            $objVM | Add-Member -Name "Memory Average / 10 sec" -Value $jobresult[5] -MemberType NoteProperty
            $objVM | Add-Member -Name "Memory Max" -Value $jobresult[6] -MemberType NoteProperty
            $objVM | Add-Member -Name "Memory Min" -Value $jobresult[7] -MemberType NoteProperty
            $objVM | Add-Member -Name "%Ocupado C:" -Value $jobresult[8].Substring(0) -MemberType NoteProperty
            $objVM | Add-Member -Name "%Ocupado D:" -Value $jobresult[9].Substring(0) -MemberType NoteProperty
            $objVM | Add-Member -Name "%Ocupado E:" -Value $jobresult[10].Substring(0) -MemberType NoteProperty
            ##$objVM | Add-Member -Name "DiskQueue Average / 10 sec" -Value $jobresult[11] -MemberType NoteProperty
            $objVM | Add-Member -Name "3 TopProcess" -Value  $jobresult[0] -MemberType NoteProperty
            $objVM | Add-Member -Name "3 TopProcessPercent" -Value  $jobresult[1] -MemberType NoteProperty
            ##$objVM | Add-Member -Name "UpTime" -Value $jobresult[12] -MemberType NoteProperty

        $resultadojob += $jobresult
        $resultado += $objVM
    
  
}


$resultado += $Serverko
$resultado += $Serverwmi


write-host "justo antes de hacer el export resultado"
$resultado 
$currentDir = Split-Path $MyInvocation.MyCommand.Path

write-host $filetxt
#$resultado -replace ("GREEN","")
#$resultado -replace ("RED","")
#$resultado -replace ("YELLOW","")
$dateformat = Get-Date -format "HH_mm_ss_dd_MM_yyyy"


$filetxt = New-Item $currentDir\Performance\$dateformat"_Report_XenApp.csv" -ItemType file
$filetxt 
$filecsv = ""
$filecsv = $filetxt


write-host "Ruta:" $filecsv

$resultado | Export-Csv -NoTypeInformation -Delimiter "," $filecsv


$Elapsed=(Get-Date)-$date
$resultado = $resultado | ConvertTo-Html -Head $headformat -Body "<H3>Server Health Report System Status XenApp 7.15 BS - $dia $hora</H3> <H4>Tiempo Trasncurrido Script $Elapsed</H2>"  |  foreach {
$PSItem -replace ("<td>GREEN", "<td style='background-color:#5ac031'>")  -replace ("<td>RED", "<td style='background-color:#fc2f48 '>") -replace ("<td>YELLOW", "<td style='background-color: #e9f618'>") } | out-file $filehtml
& $filehtml




#write-host "justo antes del Export-csv"
#write-host "Segundo resultado"
#$resultado
#$filetxt = "$currentDir\Performance\Report_XenApps.txt"
#$filetxt
#$resultado | Export-Csv -Delimiter ";" $filetxt





get-job | Remove-Job -Force

write-host "Fichero HTML = $filehtml"
write-host "Fichero CSV =" $filecsv

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
#$filetxt = "$currentDir\Performance\Report_XenApp.csv"
$att =  new-object System.Net.Mail.Attachment($filecsv)
$att1 = new-object System.Net.Mail.Attachment($filehtml)
$message.Attachments.Add($att)
$message.Attachments.Add($att1)

$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)
$message.Dispose()
}