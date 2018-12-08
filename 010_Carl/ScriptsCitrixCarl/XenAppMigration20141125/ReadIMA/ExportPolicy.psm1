# Copyright Citrix Systems, Inc.

$ErrorActionPreference = "Stop"

[System.Xml.XmlElement]
function New-PrinterNode
{
    param
    (
        [System.Object]$printer
    )

    $node = New-XmlNode "SessionPrinter"
    if ($printer.Path -ne $null) { [void]$node.AppendChild((New-XmlNode "Path" $printer.Path)) }
    if ($printer.Model -ne $null) { [void]$node.AppendChild((New-XmlNode "Model" $printer.Model)) }
    if ($printer.Location -ne $null) { [void]$node.AppendChild((New-XmlNode "Location" $printer.Location)) }
    if ($printer.Settings -ne $null)
    {
        $s = New-XmlNode "Settings"
        foreach ($setting in ($printer.Settings | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
        {
            if ($printer.Settings.$setting -ne $null)
            {
                [void]$s.AppendChild((New-XmlNode $setting $printer.Settings.$setting))
            }
        }
        [void]$node.AppendChild($s)
    }

    return $node
}

[System.Xml.XmlElement]
function New-AssignmentNode
{
    param
    (
        [System.Object]$item
    )

    $node = New-XmlNode "Assignment"

    if ($item.DefaultPrinterOption -ne $null)
    {
        [void]$node.AppendChild((New-XmlNode "DefaultPrinterOption" $item.DefaultPrinterOption))
    }
    if ($item.SpecificDefaultPrinter -ne $null)
    {
        [void]$node.AppendChild((New-XmlNode "SpecificDefaultPrinter" $item.SpecificDefaultPrinter))
    }

    $t = New-XmlNode "SessionPrinters"
    foreach ($printer in $item.SessionPrinters)
    {
        [void]$t.AppendChild((New-PrinterNode $printer))
    }
    [void]$node.AppendChild($t)

    $t = New-XmlNode "Filters"
    $item.Filters | % { [void]$t.AppendChild((New-XmlNode "Filter" $_)) }
    [void]$node.AppendChild($t)

    return $node
}

[System.Xml.XmlElement]
function New-SettingNode
{
    param
    (
        [string]$policy,
        [System.Object]$setting
    )

    Write-LogFile ([string]::Format('Exporting setting "{0}"', $setting.PSChildName)) 3

    $matches = [Regex]::Match($setting.PSParentPath, "[^:]*::Farm:\\[^\\]*\\[^\\]*\\Settings\\(.*)")
    $path = ""
    if ($matches.Success -and $matches.Groups.Count -ge 2)
    {
        $path = $matches.Groups[1].Value
    }

    $node = New-XmlNode "Setting"
    [void]$node.AppendChild((New-XmlNode "Name" $setting.PSChildName))
    [void]$node.AppendChild((New-XmlNode "Path" $path))
    [void]$node.AppendChild((New-XmlNode "State" $setting.State))
    if ($setting.Value -ne $null) { [void]$node.AppendChild((New-XmlNode "Value" $setting.Value)) }

    if ($setting.Values -ne $null)
    {
        $v = New-XmlNode "Values"
        $setting.Values | % { [void]$v.AppendChild((New-XmlNode "Value" $_)) }
        [void]$node.AppendChild($v)
    }

    # Special case for Multi-Port Policy.
    if ($setting.CgpPort1 -ne $null) { [void]$node.AppendChild((New-XmlNode "CgpPort1" $setting.CgpPort1)) }
    if ($setting.CgpPort2 -ne $null) { [void]$node.AppendChild((New-XmlNode "CgpPort2" $setting.CgpPort2)) }
    if ($setting.CgpPort3 -ne $null) { [void]$node.AppendChild((New-XmlNode "CgpPort3" $setting.CgpPort3)) }
    if ($setting.CgpPort1Priority -ne $null) { [void]$node.AppendChild((New-XmlNode "CgpPort1Priority" $setting.CgpPort1Priority)) }
    if ($setting.CgpPort2Priority -ne $null) { [void]$node.AppendChild((New-XmlNode "CgpPort2Priority" $setting.CgpPort2Priority)) }
    if ($setting.CgpPort3Priority -ne $null) { [void]$node.AppendChild((New-XmlNode "CgpPort3Priority" $setting.CgpPort3Priority)) }

    # Special case for Printer Assignments.
    if ($setting.PSChildName -eq "PrinterAssignments")
    {
        if ($setting.Assignments.Count -gt 0)
        {
            $x = New-XmlNode "Assignments"
            foreach ($a in $setting.Assignments)
            {
                [void]$x.AppendChild((New-AssignmentNode $a))
            }
            [void]$node.AppendChild($x)
        }
        else
        {
            $log = [string]::Format('INFO: Empty PrinterAssignments setting for policy "{0}", value ignored and not exported', $policy)
            Write-LogFile $log 4
        }
    }

    # Special case for Default Printer
    if ($setting.DefaultPrinterOption -ne $null)
    {
        [void]$node.AppendChild((New-XmlNode "DefaultPrinterOption" $setting.DefaultPrinterOption))
    }
    if ($setting.SpecificDefaultPrinter -ne $null)
    {
        [void]$node.AppendChild((New-XmlNode "SpecificDefaultPrinter" $setting.SpecificDefaultPrinter))
    }

    # Special case for Session Printers
    if ($setting.PSChildName -eq "SessionPrinters")
    {
        if ($setting.Printers.Count -gt 0)
        {
            $x = New-XmlNode "SessionPrinters"
            foreach ($p in $setting.Printers)
            {
                [void]$x.AppendChild((New-PrinterNode $p))
            }
            [void]$node.AppendChild($x)
        }
        else
        {
            $log = [string]::Format('INFO: Empty SessionPrinters setting for policy "{0}", value ignored and not exported', $policy)
            Write-LogFile $log 4
        }
    }

    # Special case for Health Monitoring Tests
    if ($setting.PSChildName -eq "HealthMonitoringTests")
    {
        if ($setting.HmrTests.Tests.Count -gt 0)
        {
            $h = New-XmlNode "HmrTests"
            $s = New-XmlNode "Tests"
            foreach ($test in $setting.HmrTests.Tests)
            {
                $t = New-XmlNode "Test"
                foreach ($p in ($test | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
                {
                    if ($test.$p -ne $null)
                    {
                        [void]$t.AppendChild((New-XmlNode $p $test.$p))
                    }
                }
                [void]$s.AppendChild($t)
            }
            [void]$h.AppendChild($s)
            [void]$node.AppendChild($h)
        }
        else
        {
            $log = [string]::Format('INFO: Empty HmrTests setting for policy "{0}", value ignored and not exported', $policy)
            Write-LogFile $log 4
        }
    }

    return $node
}

[System.Xml.XmlElement]
function New-FilterNode
{
    param
    (
        [System.Object]$filter
    )

    Write-LogFile ([string]::Format('Exporting assignment "{0}"', $filter.Name)) 3

    $matches = [Regex]::Match($filter.PSParentPath, "[^\\]*\\[^\\]*\\[^\\]*\\[^\\]*\\[^\\]*\\(.*)")
    $path = ""
    if ($matches.Success -and $matches.Groups.Count -ge 2)
    {
        $path = $matches.Groups[1].Value
    }

    $node = New-XmlNode "Filter"
    [void]$node.AppendChild((New-XmlNode "Name" $filter.Name))
    [void]$node.AppendChild((New-XmlNode "Mode" $filter.Mode))
    [void]$node.AppendChild((New-XmlNode "Enabled" $filter.Enabled))
    [void]$node.AppendChild((New-XmlNode "FilterType" $filter.FilterType))
    [void]$node.AppendChild((New-XmlNode "FilterValue" $filter.FilterValue))
    [void]$node.AppendChild((New-XmlNode "Synopsis" $filter.Synopsis))
    [void]$node.AppendChild((New-XmlNode "Comment" $filter.Comment))
    [void]$node.AppendChild((New-XmlNode "Path" $path))

    if ($filter.FilterType -eq "AccessControl")
    {
        [void]$node.AppendChild((New-XmlNode "ConnectionType" $filter.ConnectionType))
        [void]$node.AppendChild((New-XmlNode "AccessGatewayFarm" $filter.AccessGatewayFarm))
        [void]$node.AppendChild((New-XmlNode "AccessCondition" $filter.AccessCondition))
    }

    if ($filter.FilterType -eq "OU")
    {
        [void]$node.AppendChild((New-XmlNode "DN" $filter.DN))
    }

    return $node
}

[System.Xml.XmlElement]
function New-PolicyNode
{
    param
    (
        [System.Object]$policy
    )

    $name = $policy.Name
    Write-LogFile ([string]::Format('Exporting policy "{0}"', $name)) 1 $true

    $node = New-XmlNode "Policy" $null $name
    [void]$node.AppendChild((New-XmlNode "PolicyName" $name))
    [void]$node.AppendChild((New-XmlNode "Description" $policy.Description))
    [void]$node.AppendChild((New-XmlNode "Enabled" $policy.Enabled))
    [void]$node.AppendChild((New-XmlNode "Priority" $policy.Priority))

    Write-LogFile ([string]::Format('Exporting settings')) 2
    $settings = New-XmlNode "Settings"
    $path = ($policy.PSPath + "\Settings")
    dir $path -Recurse | ? { ($_.State -ne $null) -and ($_.State -ne "NotConfigured") } | % {
        [void]$settings.AppendChild((New-SettingNode $name $_))
    }
    [void]$node.AppendChild($settings)

    if ($name -ne "Unfiltered")
    {
        Write-LogFile ([string]::Format('Exporting object assignments')) 2
        $filters = New-XmlNode "Filters"
        $path = ($policy.PSPath + "\Filters")
        dir $path -Recurse | ? { $_.FilterType -ne $null } | % { [void]$filters.AppendChild((New-FilterNode $_)) }
        [void]$node.AppendChild($filters)
    }

    return $node
}

<#
    .Synopsis
        Export XenApp farm policies to a XML file.
    .Parameter XmlOutputFile
        The name of the XML file that stores the output. The file name must be given
        with a .XML extension. The file must not exist. If a path is given, the parent
        path of the file must exist.
    .Parameter NoLog
        Do not generate log output. If this switch is specified, the LogFile parameter
        is ignored.
    .Parameter LogFile
        File for storing the logs. If the file exists and NoClobber is not specified,
        the contents of the file are overwritten. If the file exists and NoClobber is
        specified, an error message is displayed and the script quits. If the log file
        is not specified and NoLog is also not specified, the log is still generated
        and the log file is located under the user's home directory, as specified by
        the $HOME environment variable. The name of the log file is generated using the
        current time stamp: XFarmYYYYMMDDHHmmss-RRRRRR.Log, here YYYY is the year, MM
        is month, DD is day, HH is hour, mm is minute, ss is second, and RRRRRR is a
        six digit random hexdecimal number.
    .Parameter NoClobber
        Do not overwrite an existing log file. This switch has no effect if the given
        log file does not exist.
    .Parameter NoDetails
        If specified, detailed reports about the progress of the script execution is
        not sent to the console.
    .Parameter SuppressLogo
        Suppress the logo.
    .Description
        Use this cmdlet to export the policy data in XenApp farm GPO to a XML file.
        This cmdlet must be run on a XenApp controller and must have the Citrix Group
        Policy PowerShell Provider snap-in installed on the local server. The user who
        executes this command must have read access to the policy data in the XenApp
        farm.

        The XML file references the PolicyData.XSD file, which specifies the schema
        for the data stored in the file.
    .Inputs
        None
    .Outputs
        None
    .Example
        Export-Policy -XmlOutputFile .\MyPolicies.XML
        Export policies and store them in the 'MyPolicies.XML' file in the current
        directory. The log is generated and can be found under the $HOME directory.
    .Example
        Export-Policy -XmlOutputFile .\MyPolicies.XML -LogFile .\PolicyExport.log
        Export policies and store them in the 'MyPolicies.XML' file in the current
        directory. Store the log in the file PolicyExport.log file in the current
        directory.
#>
function Export-Policy
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="Explicit")]
        [string] $XmlOutputFile,
        [Parameter(Mandatory=$false)]
        [switch] $NoLog,
        [Parameter(Mandatory=$false)]
        [switch] $NoClobber,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [string] $LogFile,
        [Parameter(Mandatory=$false)]
        [switch] $NoDetails,
        [Parameter(Mandatory=$false)]
        [switch] $SuppressLogo
    )

    Print-Logo $SuppressLogo

    if ((Get-Module LogUtilities) -eq $null)
    {
        Write-Error "Module LogUtilities.psm1 is not imported, see ReadMe.txt for usage of this script"
        return
    }
    if ((Get-Module XmlUtilities) -eq $null)
    {
        Write-Error "Module XmlUtilities.psm1 is not imported, see ReadMe.txt for usage of this script"
        return
    }

    $ShowProgress = !$NoDetails
    try
    {
        Start-Logging -NoLog:$Nolog -NoClobber:$NoClobber $LogFile -ShowProgress:$ShowProgress
    }
    catch
    {
        Write-Error $_.Exception.Message
        return
    }

    Write-LogFile ([string]::Format('Export-Policy Command Line:'))
    Write-LogFile ([string]::Format('    -XmlOutputFile {0}', $XmlOutputFile))
    Write-LogFile ([string]::Format('    -LogFile {0}', $LogFile))
    Write-LogFile ([string]::Format('    -NoClobber = {0}', $NoClobber))
    Write-LogFile ([string]::Format('    -NoDetails = {0}', $NoDetails))

    [void](Assert-XmlOutput $XmlOutputFile)

    $isXenApp = $true
    Write-LogFile ('Loading Citrix Group Policy Provider Snap-in')
    $s = (Get-PSSnapin Citrix.Common.GroupPolicy -Registered -ErrorAction SilentlyContinue)
    if ($s -ne $null)
    {
        if (Test-Path $s.ModuleName)
        {
            Import-Module $s.ModuleName
        }
        else
        {
            Import-Module ([Reflection.Assembly]::LoadWithPartialName("Citrix.GroupPolicy.PowerShellProvider"))
            $isXenApp = $false
        }
    }
    else
    {
        Write-Error ([string]::Format("{0}`n{1}", "Citrix Group Policy Provider Snapin is not installed",
            "You must have Citrix Group Policy Provider Snapin installed to use this script."))
        return
    }

    Write-LogFile ('Mount Group Policy GPO')
    if ($isXenApp)
    {
        [void](New-PSDrive -Name Farm -Root \ -PSProvider CitrixGroupPolicy -FarmGpo localhost)
    }
    else
    {
        [void](New-PSDrive -Name Farm -Root \ -PSProvider CitrixGroupPolicy -Controller localhost)
    }

    Initialize-Xml

    $root = New-XmlNode "Policies"
    [void]$root.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
    [void]$root.SetAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
    [void]$root.SetAttribute("xmlns", "PolicyData.xsd")

    Write-LogFile ('Exporting user policies') 0 $true
    $u = New-XmlNode "User"
    $count = 0
    try
    {
        Get-ChildItem Farm:\User\* | % { [void]$u.AppendChild((New-PolicyNode $_)); $count++ }
    }
    catch
    {
        Stop-Logging "User policy export aborted" $_.Exception.Message
        return
    }
    [void]$root.AppendChild($u)
    Write-LogFile ([string]::Format("{0} user policies exported", $count)) 0 $true

    Write-LogFile ('Exporting computer policies') 0 $true
    $c = New-XmlNode "Computer"
    $count = 0
    try
    {
        Get-ChildItem Farm:\Computer\* | % { [void]$c.AppendChild((New-PolicyNode $_)); $count++ }
    }
    catch
    {
        Stop-Logging "Computer policy export aborted" $_.Exception.Message
        return
    }
    [void]$root.AppendChild($c)
    Write-LogFile ([string]::Format("{0} computer policies exported", $count)) 0 $true

    Save-XmlData $root $XmlOutputFile

    Stop-Logging "Policy export completed"
}

# SIG # Begin signature block
# MIIZOwYJKoZIhvcNAQcCoIIZLDCCGSgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZga4FjkKjKfbalsKJpb6y4bL
# ncGgghQGMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggVbMIIEQ6ADAgECAhBZKu9UZuRZKU4BOkjfD0g6MA0GCSqGSIb3DQEBBQUAMIG0
# MQswCQYDVQQGEwJVUzEXMBUGA1UEChMOVmVyaVNpZ24sIEluYy4xHzAdBgNVBAsT
# FlZlcmlTaWduIFRydXN0IE5ldHdvcmsxOzA5BgNVBAsTMlRlcm1zIG9mIHVzZSBh
# dCBodHRwczovL3d3dy52ZXJpc2lnbi5jb20vcnBhIChjKTEwMS4wLAYDVQQDEyVW
# ZXJpU2lnbiBDbGFzcyAzIENvZGUgU2lnbmluZyAyMDEwIENBMB4XDTE0MDcyNDAw
# MDAwMFoXDTE1MDkxNTIzNTk1OVowgYwxCzAJBgNVBAYTAlVTMRAwDgYDVQQIEwdG
# bG9yaWRhMRgwFgYDVQQHEw9Gb3J0IExhdWRlcmRhbGUxHTAbBgNVBAoUFENpdHJp
# eCBTeXN0ZW1zLCBJbmMuMRMwEQYDVQQLFApQb3dlclNoZWxsMR0wGwYDVQQDFBRD
# aXRyaXggU3lzdGVtcywgSW5jLjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAJdE2UnYZRRIwMKq8jivnq5ZSHHTN4bbuXURARHgEYeCnqqGXLwEiu2EqbG5
# GcWhDkPdoyuoRR1TOQDDAxn9Es0dmcSpySCgqr2EjlwdIiTEB91x0upUDuLrAgsY
# 9KcsqxFGB8hZiK3i4VzloaiJfxgL59y4GaVgnjZ8ZM7N52A5eHX5vK7YQtNsxVN/
# s2Yn308BtQe1bDmSo7NaxLtrPrva9933cYTDYmi4KKJ8qXvqZ6mjUBcDNcjpSJ7N
# BvUXIK7Q5g7tUNlLg4PRuvgGFjXfx/41bufwqaneF7I7r9aPKBY/9GGEzz4+lEVh
# Dw206sF/nbbv80X+JrEydXDVYi8CAwEAAaOCAY0wggGJMAkGA1UdEwQCMAAwDgYD
# VR0PAQH/BAQDAgeAMCsGA1UdHwQkMCIwIKAeoByGGmh0dHA6Ly9zZi5zeW1jYi5j
# b20vc2YuY3JsMGYGA1UdIARfMF0wWwYLYIZIAYb4RQEHFwMwTDAjBggrBgEFBQcC
# ARYXaHR0cHM6Ly9kLnN5bWNiLmNvbS9jcHMwJQYIKwYBBQUHAgIwGRYXaHR0cHM6
# Ly9kLnN5bWNiLmNvbS9ycGEwEwYDVR0lBAwwCgYIKwYBBQUHAwMwVwYIKwYBBQUH
# AQEESzBJMB8GCCsGAQUFBzABhhNodHRwOi8vc2Yuc3ltY2QuY29tMCYGCCsGAQUF
# BzAChhpodHRwOi8vc2Yuc3ltY2IuY29tL3NmLmNydDAfBgNVHSMEGDAWgBTPmanq
# eyb0S8mOj9fwBSbv49KnnTAdBgNVHQ4EFgQUY2hVDhNISSu2X+tH8+1CUYnxnMEw
# EQYJYIZIAYb4QgEBBAQDAgQQMBYGCisGAQQBgjcCARsECDAGAQEAAQH/MA0GCSqG
# SIb3DQEBBQUAA4IBAQAkEEcP/T+8Lo4uMgEtZlGbgzwynjip4miCHfhlImbna4lR
# T+h3vLzJPh7cRGhgtUAKPU4WQ+VnaPEfLAe6gYxFm8bbp9ws2LtDeuN/wQPFeMpW
# dbMyLH7fGdZIqPz/KE9A4+Gk+EmJhsAyfpfy/dSCrBkjTNRfDZyBK9IST8e3MXH4
# q8JVRQUWy7D0+hudTG++7CIyu5fYSgsWqo8WzlKF2pKn25PFxWkzuTnI1+2X7Kxs
# 4+EaaJxwXUMVPZa4mCxz0OKZWLczqNOhEmELFhKHMb5R59WFbW9bt4Sl3e2sqHWS
# q/5L0rv7KSx2eksR7dXxYsvY81GMVVIQPGWA49HwMIIGCjCCBPKgAwIBAgIQUgDl
# qiVW/BqG7ZbJ1EszxzANBgkqhkiG9w0BAQUFADCByjELMAkGA1UEBhMCVVMxFzAV
# BgNVBAoTDlZlcmlTaWduLCBJbmMuMR8wHQYDVQQLExZWZXJpU2lnbiBUcnVzdCBO
# ZXR3b3JrMTowOAYDVQQLEzEoYykgMjAwNiBWZXJpU2lnbiwgSW5jLiAtIEZvciBh
# dXRob3JpemVkIHVzZSBvbmx5MUUwQwYDVQQDEzxWZXJpU2lnbiBDbGFzcyAzIFB1
# YmxpYyBQcmltYXJ5IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IC0gRzUwHhcNMTAw
# MjA4MDAwMDAwWhcNMjAwMjA3MjM1OTU5WjCBtDELMAkGA1UEBhMCVVMxFzAVBgNV
# BAoTDlZlcmlTaWduLCBJbmMuMR8wHQYDVQQLExZWZXJpU2lnbiBUcnVzdCBOZXR3
# b3JrMTswOQYDVQQLEzJUZXJtcyBvZiB1c2UgYXQgaHR0cHM6Ly93d3cudmVyaXNp
# Z24uY29tL3JwYSAoYykxMDEuMCwGA1UEAxMlVmVyaVNpZ24gQ2xhc3MgMyBDb2Rl
# IFNpZ25pbmcgMjAxMCBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# APUjS16l14q7MunUV/fv5Mcmfq0ZmP6onX2U9jZrENd1gTB/BGh/yyt1Hs0dCIzf
# aZSnN6Oce4DgmeHuN01fzjsU7obU0PUnNbwlCzinjGOdF6MIpauw+81qYoJM1SHa
# G9nx44Q7iipPhVuQAU/Jp3YQfycDfL6ufn3B3fkFvBtInGnnwKQ8PEEAPt+W5cXk
# lHHWVQHHACZKQDy1oSapDKdtgI6QJXvPvz8c6y+W+uWHd8a1VrJ6O1QwUxvfYjT/
# HtH0WpMoheVMF05+W/2kk5l/383vpHXv7xX2R+f4GXLYLjQaprSnTH69u08MPVfx
# MNamNo7WgHbXGS6lzX40LYkCAwEAAaOCAf4wggH6MBIGA1UdEwEB/wQIMAYBAf8C
# AQAwcAYDVR0gBGkwZzBlBgtghkgBhvhFAQcXAzBWMCgGCCsGAQUFBwIBFhxodHRw
# czovL3d3dy52ZXJpc2lnbi5jb20vY3BzMCoGCCsGAQUFBwICMB4aHGh0dHBzOi8v
# d3d3LnZlcmlzaWduLmNvbS9ycGEwDgYDVR0PAQH/BAQDAgEGMG0GCCsGAQUFBwEM
# BGEwX6FdoFswWTBXMFUWCWltYWdlL2dpZjAhMB8wBwYFKw4DAhoEFI/l0xqGrI2O
# a8PPgGrUSBgsexkuMCUWI2h0dHA6Ly9sb2dvLnZlcmlzaWduLmNvbS92c2xvZ28u
# Z2lmMDQGA1UdHwQtMCswKaAnoCWGI2h0dHA6Ly9jcmwudmVyaXNpZ24uY29tL3Bj
# YTMtZzUuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AudmVyaXNpZ24uY29tMB0GA1UdJQQWMBQGCCsGAQUFBwMCBggrBgEFBQcDAzAo
# BgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVmVyaVNpZ25NUEtJLTItODAdBgNVHQ4E
# FgQUz5mp6nsm9EvJjo/X8AUm7+PSp50wHwYDVR0jBBgwFoAUf9Nlp8Ld7LvwMAnz
# Qzn6Aq8zMTMwDQYJKoZIhvcNAQEFBQADggEBAFYi5jSkxGHLSLkBrVaoZA/ZjJHE
# u8wM5a16oCJ/30c4Si1s0X9xGnzscKmx8E/kDwxT+hVe/nSYSSSFgSYckRRHsExj
# jLuhNNTGRegNhSZzA9CpjGRt3HGS5kUFYBVZUTn8WBRr/tSk7XlrCAxBcuc3IgYJ
# viPpP0SaHulhncyxkFz8PdKNrEI9ZTbUtD1AKI+bEM8jJsxLIMuQH12MTDTKPNjl
# N9ZvpSC9NOsm2a4N58Wa96G0IZEzb4boWLslfHQOWP51G2M/zjF8m48blp7FU3aE
# W5ytkfqs7ZO6XcghU8KCU2OvEg1QhxEbPVRSloosnD2SGgiaBS7Hk6VIkdMxggSf
# MIIEmwIBATCByTCBtDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJ
# bmMuMR8wHQYDVQQLExZWZXJpU2lnbiBUcnVzdCBOZXR3b3JrMTswOQYDVQQLEzJU
# ZXJtcyBvZiB1c2UgYXQgaHR0cHM6Ly93d3cudmVyaXNpZ24uY29tL3JwYSAoYykx
# MDEuMCwGA1UEAxMlVmVyaVNpZ24gQ2xhc3MgMyBDb2RlIFNpZ25pbmcgMjAxMCBD
# QQIQWSrvVGbkWSlOATpI3w9IOjAJBgUrDgMCGgUAoIGcMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBTKBvbpmryM2LtSGBSp152qHHverTA8BgorBgEEAYI3AgEMMS4w
# LKAYgBYAQwBpAHQAcgBpAHgAIABGAGkAbABloRCADnd3dy5jaXRyaXguY29tMA0G
# CSqGSIb3DQEBAQUABIIBAIeEobqg7WqRW2fawPyEkaKsU9h34wzXKRrH0WEYjuGv
# rriR8NF79Z15iazRv0u8ZtKaqchVhbV1HKIz3MxtJXRSyciAc0/JWtzy8GplJ9LW
# iWrUdRLegxr1u0k3SK1kf0LPa+lLQ0LwOqvNGIV0lfyj7CAhaCBylwJGISMjB9Xd
# UWr7ETrhWEqoUlKBUlqoL0Q73my2irS3mzVMgV6Npn+C7H0js1lBkPa5e8GrBeSY
# BYeFmONYqBhClSEI1/s0UA/BFh3WA1k/UEtDvCtjhoqQ4f93RT+KHz/3qgSLLz/J
# rZ4bWE1Fm+wpJGXfZ5NgfSvxlkm7azj6QBZydo/4L2ChggILMIICBwYJKoZIhvcN
# AQkGMYIB+DCCAfQCAQEwcjBeMQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50
# ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcg
# U2VydmljZXMgQ0EgLSBHMgIQDs/0OMj+vzVuBNhqmBsaUDAJBgUrDgMCGgUAoF0w
# GAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTQxMTI1
# MjAxNTE5WjAjBgkqhkiG9w0BCQQxFgQUQKqHvCEtjVNnzj7Lv5BpJx+fHrkwDQYJ
# KoZIhvcNAQEBBQAEggEAH9ephHV0VIpS2V2shoSmEgkDHRss+gEtTlSt6dzXDGlF
# Oyu/qZH3+4p96Y1Y6OIi9Nnpwu+uCDgVEFaKBPCd7waft6RQiEv3B0VZi2CU3jKs
# aWaMM/LL/7iYbzruUnzfjsM3sf1cBbTQX8RkuvPI3rpB0fgDMdNLtMgMxLhyjTmk
# Mjnwrkv06Ar3So3UQLgvh1/2+tLFpdkbBUj2gKSK/PfusHn1W+xzCge3BNtNyDDq
# r0HyzP/wRQ/Z7xHAp+y/KRTlcaL95yvCY+fqimlvz6dCLgSjF7kavkL88LkQ7vNg
# aKZReD9meurzSozEf87z54/Ent5CGHEBVCVekQ1w3Q==
# SIG # End signature block
