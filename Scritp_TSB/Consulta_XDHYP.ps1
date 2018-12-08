
asnp Citrix*
$ErrorActionPreference=Continue
Get-ChildItem -Path "XDHyp:\Connections\VMM_NH_wcpnhidsc004\SDI_NH.hostgroup\wfpnhidhy062.host" -force -recurse | ?{ $_.IsMachine } 