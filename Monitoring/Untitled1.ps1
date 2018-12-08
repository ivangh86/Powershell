$computers = "XAVBSV3042P","XAPBSV3027P"
foreach ($computer in $computers){

if (Test-WSMan -ComputerName $computer -ErrorAction Ignore) {
Write-host “connect to $computer”
}
else {
Write-Warning -Message “Couldn’t connect to $computer”
}

}