Function Get-VMHost
{
    (get-item "HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters").GetValue("HostName")
}

Get-VMHost


