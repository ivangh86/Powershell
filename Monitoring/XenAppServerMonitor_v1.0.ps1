Param($silo,$offset,$country)
################################################################################################
## XenAppServerMonitor
## Bart Jacobs (bart@bj-it.be)
## v1.0
################################################################################################
 if ((Get-PSSnapin "Citrix.Configuration.Admin.V2" -EA silentlycontinue) -eq $null) {
	try { Add-PSSnapin Citrix* -ErrorAction Stop }
	catch { write-error "Error loading XenApp Powershell snapin"; Return }
}

# Change the below variables to suit your environment
#==============================================================================================
# Default load evaluator assigned to servers. Can have multiple values in format "LE1", "LE2",
# if a match is made to ANY of the listed LEs SUCCESS is returned for the LE check.
$defaultLE       = "LE1"

# Silo Parameter
# The script is started with an identifier of the Silo of servers you are monitoring. This section parses that parameter and populates some variables accordingly
# More silo's -> More 
if ($silo -eq "SDI W2K8R2 SOLVIA DES")
	{
	$serversilo = "SDI W2K8R2 SOLVIA DES*"
	$silotitle = "XenApp Producion Silo Citrix Dashboard"
	$siloerrortitle = "XenApp Production Silo Error Report"
	
	}

#==============================================================================================
 
$currentDir = Split-Path $MyInvocation.MyCommand.Path
$outputdir = "c:\temp\otros"
$logfile    = Join-Path $outputdir $logfilename
$resultsHTM = Join-Path $outputdir $resultfilename
$errorsHTM  = Join-Path $outputdir $errorfilename

$headerNames  = "Ping", "CPU", "Memory", "Disk", "DiskQueue"
$headerWidths = "4", "5", "5", "5", "5"

#==============================================================================================
function LogMe() {
	Param(
		[parameter(Mandatory = $true, ValueFromPipeline = $true)] $logEntry,
		[switch]$display,
		[switch]$error,
		[switch]$warning,
		[switch]$progress
	)


	if ($error) {
		$logEntry = "[ERROR] $logEntry" ; Write-Host "$logEntry" -Foregroundcolor Red}
	elseif ($warning) {
		Write-Warning "$logEntry" ; $logEntry = "[WARNING] $logEntry"}
	elseif ($progress) {
		Write-Host "$logEntry" -Foregroundcolor Green}
	elseif ($display) {
		Write-Host "$logEntry" }
	 
	#$logEntry = ((Get-Date -uformat "%D %T") + " - " + $logEntry)
	$logEntry | Out-File $logFile -Append
}

#==============================================================================================
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


#==============================================================================================
Function writeHtmlHeader
{
param($title, $fileName)
$date = ( Get-Date -format R)
$head = @"
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
<title>$title</title>
<STYLE TYPE="text/css">
<!--
td {
font-family: Tahoma;
font-size: 11px;
border-top: 1px solid #999999;
border-right: 1px solid #999999;
border-bottom: 1px solid #999999;
border-left: 1px solid #999999;
padding-top: 0px;
padding-right: 0px;
padding-bottom: 0px;
padding-left: 0px;
overflow: hidden;
}
body {
margin-left: 5px;
margin-top: 5px;
margin-right: 0px;
margin-bottom: 10px;
table {
table-layout:fixed; 
border: thin solid #000000;
}
-->
</style>
</head>
<body>
<table width='1200'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='48' align='center' valign="middle">
<font face='tahoma' color='#003399' size='4'>
<!--<img src="http://servername/administration/icons/xenapp.png" height='42'/>-->
<! <strong>$title - $date</strong></font>
</td>
</tr>
</table>

<table width='1200'>
<tr bgcolor='#CCCCCC'>
<td width=100% height='48' align='center' valign="middle">
<font face='tahoma' color='#003399' size='4'>
</td>
</tr>
</table>
"@
$head | Out-File $fileName
}

# ==============================================================================================
Function writeTableHeader
{
param($fileName)
$tableHeader = @"
<table width='1200'><tbody>
<tr bgcolor=#CCCCCC>
<td width='6%' align='center'><strong>ServerName</strong></td>
"@

$i = 0
while ($i -lt $headerNames.count) {
	$headerName = $headerNames[$i]
	$headerWidth = $headerWidths[$i]
	$tableHeader += "<td width='" + $headerWidth + "%' align='center'><strong>$headername</strong></td>"
	$i++
}

$tableHeader += "</tr>"

$tableHeader | Out-File $fileName -append
}

# ==============================================================================================
Function writeData
{
	param($data, $fileName)
	
	$data.Keys | sort | foreach {
		$tableEntry += "<tr>"
		$computerName = $_
		$tableEntry += ("<td bgcolor='#CCCCCC' align=center><font color='#003399'>$computerName</font></td>")
		#$data.$_.Keys | foreach {
		$headerNames | foreach {
			#"$computerName : $_" | LogMe -display
			try {
				if ($data.$computerName.$_[0] -eq "SUCCESS") { $bgcolor = "#387C44"; $fontColor = "#FFFFFF" }
				elseif ($data.$computerName.$_[0] -eq "WARNING") { $bgcolor = "#FF7700"; $fontColor = "#FFFFFF" }
				elseif ($data.$computerName.$_[0] -eq "ERROR") { $bgcolor = "#FF0000"; $fontColor = "#FFFFFF" }
				else { $bgcolor = "#CCCCCC"; $fontColor = "#003399" }
				$testResult = $data.$computerName.$_[1]
			}
			catch {
				$bgcolor = "#CCCCCC"; $fontColor = "#003399"
				$testResult = ""
			}
			
			$tableEntry += ("<td bgcolor='" + $bgcolor + "' align=center><font color='" + $fontColor + "'>$testResult</font></td>")
		}
		
		$tableEntry += "</tr>"
	}
	
	$tableEntry | Out-File $fileName -append
}

 
# ==============================================================================================
Function writeHtmlFooter
{
param($fileName)
@"
</table>
<table width='1200'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='25' align='left'>
<font face='courier' color='#003399' size='2'><strong>Default Load Evaluator  = $DefaultLE</strong></font>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='25' align='left'>
<font face='courier' color='#003399' size='2'><strong>Default VDISK Image         = $DefaultVDISK</strong></font>
</td>
</tr>
</table>
</body>
</html>
"@ | Out-File $FileName -append
}

# ==============================================================================================
# ==                                       MAIN SCRIPT                                        ==
# ==============================================================================================
# Company Parameter
# The script is started with an identifier of the country/datacenter of servers you are monitoring. This section parses that parameter and populates some variables accordingly
# More silo's -> More 

#if ($country -eq "BE")
#	{
#	$XAZDC = "XAZDP001"       
#	}
#
#if ($country -eq "NL")
#	{
#   $XAZDC = "XAZDP002"
#    }
	
#   Set-XADefaultComputerName -Scope CurrentUser -ComputerName $XAZDC

# Script loop

while ($true) 
{
"Sleeping " + $offset + "s" | LogMe -display -progress
# To run the script multiple times, an offset parameter was introduced. Depending on that parameter, the script loop waits for a number of seconds defined in this parameter.
start-sleep -s $offset 

# Get Start Time
$startDTM = (Get-Date)
"Checking server health..." | LogMe -display
"Remove logfile..." | LogMe -display
rm $logfile -force -EA SilentlyContinue

# Data structure overview:
# Individual tests added to the tests hash table with the test name as the key and a two item array as the value.
# The array is called a testResult array where the first array item is the Status and the second array
# item is the Result. Valid values for the Status are: SUCCESS, WARNING, ERROR and $NULL.
# Each server that is tested is added to the allResults hash table with the computer name as the key and
# the tests hash table as the value.
# The following example retrieves the Logons status for server NZCTX01:
# $allResults.NZCTX01.Logons[0]

$allResults = @{}

"Get server list..." | LogMe -display
 
Get-brokermachine | select HostedMachineName | Where{$_.DesktopGroupName -like $serversilo} | % {

	$tests = @{}	
	$server = $_.HostedMachineName
		
	$server | LogMe -display -progress
	        			
	# Ping server 
	$result = Ping $server 100
	if ($result -ne "SUCCESS") { $tests.Ping = "ERROR", $result }
	else { $tests.Ping = "SUCCESS", $result 
									
            #CPU Usage
            $CPUused = Get-Counter -Computername $server -counter "\Processor(_total)\% Processor Time" | Foreach {[math]::round($_.CounterSamples[0].CookedValue, 2)}
            $CPULoad = "$($CPUused)%"
			
            $tests.CPU = "SUCCESS", $CPULoad
            if ($CPUused -gt 80) {
			$tests.CPU = "WARNING", $CPULoad}
            if ($CPUused -gt 95) {
			$tests.CPU = "ERROR", $CPULoad}
            
            #Memory Usage                 
            $MEMavail = Get-Counter -Computername $server -counter "\Memory\Available Mbytes" | Foreach {[math]::round($_.CounterSamples[0].CookedValue, 2)}
            $MEMtotal = ((get-wmiobject -Computername $server -class "win32_physicalmemory" -namespace "root\CIMV2").Capacity)
            $sum = $MEMtotal -join '+'
            $MEMsum = Invoke-Expression $sum
            $MEMtotalGB = ($MEMsum / 1048576)
            $MEMload = ($MEMavail * 100 / $MEMtotalGB) 
            $Memloadcounter = (100 - [math]::round($MEMload, 2))
            $MemUsed = "$($Memloadcounter)%"
            
            $tests.Memory = "SUCCESS", $MemUsed
            if ($Memloadcounter -gt 80) {
			$tests.Memory = "WARNING", $MemUsed}
            if ($Memloadcounter -gt 95) {
			$tests.Memory = "ERROR", $MemUsed}
           

            #Disk Usage
            $diskcounter = Get-Counter -Counter "\LogicalDisk(C:)\% Free Space" -ComputerName $server
            $diskfree = $diskcounter.CounterSamples | Select-Object CookedValue
            $diskUsedcounter = (100 - [math]::round($diskfree.Cookedvalue, 2))     
            $CDrive = "$($diskUsedcounter)%"       
            
            $tests.Disk = "SUCCESS", $CDrive
            if (($diskUsedcounter) -gt 85) {
			$tests.Disk = "WARNING", $CDrive}
            if (($diskUsedcounter) -gt 95) {
			$tests.Disk = "ERROR", $CDrive}

            #Disk Queue
	        $diskcounterq = Get-Counter -Counter "\LogicalDisk(C:)\Avg. disk queue length" -ComputerName $server
            $diskq = $diskcounterq.CounterSamples | Select-Object CookedValue
            $diskQcounter = $([math]::round($diskq.Cookedvalue, 2))
            
            $tests.DiskQueue = "SUCCESS", $diskqcounter
            if (($diskqcounter) -gt 1) {
			$tests.DiskQueue = "WARNING", $diskqcounter}
            if (($disqcounter) -gt 2) {
			$tests.DiskQueue = "ERROR", $diskqcounter}
	}

	$allResults.$server = $tests
}

# Write all results to an html file
Write-Host ("Saving results to html report: " + $resultsHTM)
writeHtmlHeader $silotitle $resultsHTM
writeTableHeader $resultsHTM
$allResults | sort-object -property FolderPath | % { writeData $allResults $resultsHTM }

# Get End Time
$endDTM = (Get-Date)

# Echo Time elapsed
"Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"

}