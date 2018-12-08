## VARIABLES


Param(
  [Parameter(Mandatory=$True,Position=1,HelpMessage="It uses LIKE FILTER Criteria - You can filter using * and ? as wildcard")]
   [string]$Filter,  
   [Parameter(Mandatory=$True,Position=2)]
   [string]$FilePath_CSV
)

Set-ExecutionPolicy Bypass
if ((Get-PSSnapin "Citrix.XenApp.Commands" -EA silentlycontinue) -eq $null) {
	try { Add-PSSnapin Citrix.XenApp.Commands -ErrorAction Stop }
	catch { write-error "Error loading XenApp Powershell snapin"; Return }
}

#$Help1 = write-host "It uses MATCH FILTER Criteria - You can filter using ? as wildcard e.g. Na?e => would match name,nate,nane,..." -foregroundcolor "Yellow" + write-host "In order to use special characters, you will need to escape them using backslash e.g. look for + use \+" -foregroundcolor "Yellow"

## ERROR CONTROL
# -foregroundcolor "color_name" -backgroundcolor "color_name" •Black•Blue•Cyan•DarkBlue•DarkCyan•DarkGray•DarkGreen•DarkMagenta•DarkRed•DarkYellow•Gray•Green•Magenta•Red•White•Yellow

function Error_Control {
	if ($error_filter) {
		Write-Host "EXIT SCRIPT as no matching Apps using filter criteria: ""$Filter"" have not been found" -foregroundcolor "Red"
		exit 1
	}
	if ($error_path) {
		Write-Host "EXIT SCRIPT with error as it is not able to export the info to specified path: ""$FilePath_CSV"" seems not valid" -foregroundcolor "Red"
		exit 1
	}
	if ($error_append) {
	write-host "Error with APPEND function to CSV to: $FilePath_CSV , TRY TO CREATE NEW FILE" -foregroundcolor "green"
		$Report_app | Export-csv -Delimiter ";" -Path $FilePath_CSV -Erroraction SilentlyContinue -ErrorVariable error_path -notypeinformation -Force
		if ($error_path) {Error_Control}
			else {
			write-host "Successfuly exported CSV to: $FilePath_CSV" -foregroundcolor "green"
		}
	invoke-item $FilePath_CSV
	}
}

## LIGHT QUERY

$Apps = @()
$Apps = Get-XAApplication -folderpath "$Filter" -Erroraction SilentlyContinue -ErrorVariable error_filter | select-object -ExpandProperty BrowserName
if ($error_filter) {Error_Control}
write-host "CHECK if the desired Apps match your filter criteria: ""$Filter"" " -foregroundcolor "Yellow"
write-host "Filtered Matching Apps list:"
foreach ($app_name in $Apps) {
	write-host $app_name -foregroundcolor "Green"
}

## HEAVY QUERY

write-host "Press any key to continue or Ctrl+C to cancel"
$HOST.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown,AllowCtrlC") | OUT-NULL
$HOST.UI.RawUI.Flushinputbuffer()

$Report_App = @()
$i=0

foreach ($app_name in $Apps) {
	$Report_App += ,(Get-XAApplicationReport -BrowserName $app_name | select-object BrowserName,DisplayName,Description,CommandLineExecutable,Enabled,AddToClientStartMenu,@{Name="WorkerGroupNames";Expression={$_.WorkerGroupNames -join " | "}},@{Name="Accounts";Expression={$_.Accounts -join " | "}})
	write-host " "
	write-host " //////////// App Info - $app_name ////////////"
	write-host " "
	$Report_App[$i]
	write-host "//////////////////////////////////////////////////////////////////////////////////////////////"
	$i++
}

## EXPORT RESULTS 

write-host ""
write-host "Trying to export info to specified FILE path ""$FilePath_CSV""" -foregroundcolor "Yellow"
write-host ""
write-host "APPEND info to destination FILE path? (Only compatible above PS 3.0) type yes(Y) or no(N)" -foregroundcolor "Yellow"
$answer = Read-Host "(Y) or (N)"
while("Y","N" -notcontains $answer)
{
	$answer = Read-Host "(Y) or (N)"
}
if ($answer -eq "Y") {
$Report_app | Export-csv -Delimiter ";" -Path $FilePath_CSV -Erroraction SilentlyContinue -ErrorVariable error_append -notypeinformation -Append
if ($error_append) {Error_Control}
else {
write-host "Successfuly exported and APPENDED info to CSV to: $FilePath_CSV" -foregroundcolor "green"
}
invoke-item $FilePath_CSV

}
if ($answer -eq "N") {
$Report_app | Export-csv -Delimiter ";" -Path $FilePath_CSV -Erroraction SilentlyContinue -ErrorVariable error_path -notypeinformation -Force
if ($error_path) {Error_Control}
else {
write-host "Successfuly exported CSV to: $FilePath_CSV" -foregroundcolor "green"
}
invoke-item $FilePath_CSV
}
exit 0