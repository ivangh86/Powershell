# Copyright Citrix Systems, Inc.

#Requires -Version 2.0

. .\Version.ps1

$ErrorActionPreference = "Stop"

Set-Variable -Name NoPreview -Scope Script -Value $true
Set-Variable -Name IgnoredPolicies -Scope Script -Value 0
Set-Variable -Name ImportedPolicies -Scope Script -Value 0

function Set-PrinterSettings
{
    param
    (
        [string]$path,
        [object]$printer
    )

    $item = (Get-Item ($path + "\SessionPrinters\*")) | ? { $_.Path -eq $printer.Path }
    if ($item -eq $null)
    {
        Write-LogFile ([string]::Format('WARNING: Printer "{0}", not found, skipped', $printer.Path)) 5
        return
    }

    $item.Model = $printer.Model
    $item.Location = $printer.Location
    $printer.Settings.get_ChildNodes() | % {
        Write-LogFile ([string]::Format('{0} = "{1}"', [string]$n, $_.Name)) 5
        $n = $_.Name;
        Write-LogFile ([string]::Format('{0} = "{1}"', [string]($item.Settings.$n), [string]($printer.Settings.$n))) 5
        if ($Script:NoPreview)
        {
            $item.Settings.$n = $printer.Settings.$n
        }
    } 
}

function Add-PrinterAssignment
{
    param
    (
        [string]$path,
        [object]$setting
    )

    Write-LogFile ([string]::Format('Enabling printer assignments for "{0}"', $path)) 4

    [int]$i = 1
    foreach ($a in $setting.Assignments.Assignment)
    {
        $printers = @()
        $nodes = ($a.get_ChildNodes() | ? { $_.Name -eq "SessionPrinters" }).get_ChildNodes()
        $nodes | % { $printers += $_.Path }
        $filters = @()
        ($a.get_childNodes() | ? { $_.Name -eq "Filters" }).get_ChildNodes() | % { $filters += $_.InnerText }
        $dpo = $a.DefaultPrinterOption
        $sdp = $a.SpecificDefaultPrinter
        $log = [string]::Format('New-Item -Path ("{0}") -Name {1} -Filter "{2}" -DefaultPrinterOption {3} -SessionPrinter "{4}"',
            ($path + "\Assignments"), $i, [string]$filters, $dpo, [string]$printers)
        Write-LogFile ($log) 5
        if ($Script:NoPreview)
        {
            $item = New-Item -Path ($path + "\Assignments") -Name $i -Filter $filters -DefaultPrinterOption $dpo -SessionPrinter $printers
        }
        if (!([string]::IsNullOrEmpty($sdp)))
        {
            if ($Script:NoPreview)
            {
                Write-LogFile ([string]::Format('Set-ItemProperty -Path "{0}" -Name SpecificDefaultPrinter -Value {1}', $item, $sdp)) 5
                Set-ItemProperty -Path $item -Name SpecificDefaultPrinter -Value $sdp
            }
            else
            {
                Write-LogFile ([string]::Format('Set-ItemProperty -Path "<item path>" -Name SpecificDefaultPrinter -Value {0}', $sdp)) 5
            }
        }
        if ($Script:NoPreview)
        {
            $sp = $item.PSPath.SubString("CitrixGroupPolicy::".Length)
            foreach ($p in $nodes)
            {
                Set-PrinterSettings $sp $p
            }
        }
    }

}

Set-Variable -Option Constant -Name IgnoredCategories -Value `
@(`
    "Licensing",`
    "PowerAndCapacityManagement",`
    "ServerSettings",`
    "XMLService",`
    "ServerSessionSettings",`
    "Shadowing",`
    "Ports"
)

Set-Variable -Option Constant -Name IgnoredSettings -Value `
@(`
    "ConcurrentLogOnLimit",`
    "LingerDisconnectTimerInterval",`
    "LingerTerminateTimerInterval",`
    "PrelaunchDisconnectTimerInterval",`
    "PrelaunchTerminateTimerInterval",`
    "PromptForPassword",`
    "PvsIntegrationEnabled",`
    "PvsImageUpdateDeadlinePeriod",`
    "RAVEBufferSize",`
    "UseDefaultBufferSize",`
    "EnableEnhancedCompatibility",`
    "EnhancedCompatibility",`
    "EnhancedCompatibilityPrograms",`
    "FilterAdapterAddresses",`
    "FilterAdapterAddressesPrograms",`
    "AllowSpeedFlash",`
    "HDXFlashEnable",`
    "HDXFlashBackwardsCompatibility",`
    "HDXFlashEnableLogging",`
    "HDXFlashLatencyThreshold",`
    "LimitComBw",`
    "LimitComBWPercent",`
    "LimitLptBw",`
    "LimitLptBwPercent",`
    "ClientPrinterNames",`
    "AllowRetainedRestoredClientPrinters",`
    "FlashClientContentURLRewritingRules",`
    "OemChannelBandwidthLimit",`
    "OemChannelBandwidthPercent",`
    "OemChannels"
)

function Enable-Setting
{
    param
    (
        [string]$root,
        [object]$setting
    )

    if (($setting -eq $null) -or ([string]::IsNullOrEmpty($setting.Name)))
    {
        return
    }

    Write-LogFile ([string]::Format('Importing setting "{0}"', $setting.Name)) 3
    foreach ($cat in ([string]$setting.Path).Split('\'))
    {
        if ($IgnoredCategories -contains $cat)
        {
            Write-LogFile ([string]::Format('INFO: Unsupported category "{0}" for "{1}", setting ignored', $setting.Path, $setting.Name)) 4
            return
        }
    }
    if ($IgnoredSettings -contains $setting.Name)
    {
        Write-LogFile ([string]::Format('INFO: Unsupported setting "{0}", ignored', $setting.Name)) 4
        return
    }

    $path = $root + "\Settings\" + $setting.Path + "\" + $setting.Name

    Write-LogFile ([string]::Format('Enabling setting "{0}"', $path)) 3
    $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name State -Value {1}', $path, $setting.State)
    Write-LogFile ($log) 4
    if ($Script:NoPreview)
    {
        Set-ItemProperty -Path $path -Name "State" -Value $setting.State
    }

    if ($setting.Value -ne $null)
    {
        $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name Value -Value "{1}"', $path, $setting.Value)
        Write-LogFile ($log) 4
        if ($Script:NoPreview)
        {
            Set-ItemProperty -Path $path -Name "Value" -Value $setting.Value
        }
    }

    if ($setting.DefaultPrinterOption -ne $null)
    {
        $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name DefaultPrinterOption -Value "{1}"',
            $path, $setting.DefaultPrinterOption)
        Write-LogFile ($log) 4
        if ($Script:NoPreview)
        {
            Set-ItemProperty -Path $path -Name "DefaultPrinterOption" -Value $setting.DefaultPrinterOption
        }
    }

    1..3 | % {
        $t = "CgpPort" + $_
        if ($setting.$t -ne $null)
        {
            $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name "{1}" -Value "{2}"', $path, $t, $setting.$t)
            Write-LogFile ($log) 4
            if ($Script:NoPreview)
            {
                Set-ItemProperty -Path $path -Name $t -Value $setting.$t
            }
        }
    }

    1..3 | % {
        $t = "CgpPort" + $_ + "Priority"
        if ($setting.$t -ne $null)
        {
            $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name "{1}" -Value "{2}"', $path, $t, $setting.$t)
            Write-LogFile ($log) 4
            if ($Script:NoPreview)
            {
                Set-ItemProperty -Path $path -Name $t -Value $setting.$t
            }
        }
    }

    if ($setting.HmrTests -ne $null)
    {
        Write-LogFile ([string]::Format('INFO: HMR test settings for "{0}" ignored', $path)) 4
    }

    if ($setting.Values -ne $null)
    {
        $values = @()
        $setting.Values.Value | % { $values += $_ }
        $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name Values -Value @({1})', $path, [string]$values)
        Write-LogFile ($log) 4
        if ($Script:NoPreview)
        {
            Set-ItemProperty -Path $path -Name "Values" -Value $values
        }
    }

    if ($setting.Name -eq "PrinterAssignments")
    {
        Add-PrinterAssignment $path $setting
    }
}

Set-Variable -Scope Script -Name FilterAdded -Value $false

function Add-Filter
{
    param
    (
        [string]$root,
        [object]$filter
    )

    $type = $filter.FilterType
    $name = $filter.Name

    if ([string]::IsNullOrEmpty($type) -or [string]::IsNullOrEmpty($name))
    {
        return
    }

    Write-LogFile ([string]::Format('Importing {0} assignment "{1}"', $type, $name)) 4

    if ($type -eq "WorkerGroup")
    {
        Write-LogFile ([string]::Format('INFO: WorkerGroup filter "{0}" ignored', $name)) 5
        return
    }

    $path = Join-Path (Join-Path $root "Filters") $type
    $item = Join-Path $path $name
    if (Test-Path $item)
    {
        Write-LogFile ([string]::Format('Filter "{0}" already exists, skipped', $name)) 5 -isWarning
        return
    }

    if (($type -eq "BranchRepeater") -or ($type -eq "AccessControl"))
    {
        $log = [string]::Format('New-Item -Path "{0}" -Name "{1}" -ErrorAction SilentlyContinue', $path, $name)
        Write-LogFile ($log) 5
        if ($Script:NoPreview)
        {
            [void]($r = New-Item -Path $path -Name $name -ErrorAction SilentlyContinue)
        }
    }
    elseif ($type -eq "OU")
    {
        $log = [string]::Format('New-Item -Path "{0}" -Name "{1}" -DN "{2}" -ErrorAction SilentlyContinue', $path, $name, $filter.DN)
        Write-LogFile ($log) 5
        try
        {
            if ($Script:NoPreview)
            {
                [void]($r = New-Item -Path $path -Name $name -DN $filter.DN -ErrorAction SilentlyContinue)
            }
        }
        catch
        {
        }
    }
    else
    {
        $log = [string]::Format('New-Item -Path "{0}" -Name "{1}" -FilterValue "{2}" -ErrorAction SilentlyContinue',
            $path, $name, $filter.FilterValue)
        Write-LogFile ($log) 5
        try
        {
            if ($Script:NoPreview)
            {
                [void]($r = New-Item -Path $path -Name $name -FilterValue $filter.FilterValue -ErrorAction SilentlyContinue)
            }
        }
        catch
        {
        }
    }
    if ($r -eq $null)
    {
        Write-LogFile ([string]::Format('Invalid data in filter "{0}", ignored', $name)) 5 -isWarning
        return
    }

    $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name Enabled -Value {1}', $item, $filter.Enabled)
    Write-LogFile ($log) 5
    if ($Script:NoPreview)
    {
        Set-ItemProperty -Path $item -Name Enabled -Value $filter.Enabled
    }

    $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name Mode -Value {1}', $item, $filter.Mode)
    Write-LogFile ($log) 5
    if ($Script:NoPreview)
    {
        Set-ItemProperty -Path $item -Name Mode -Value $filter.Mode
    }

    if (![string]::IsNullOrEmpty($filter.Comment))
    {
        $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name Comment -Value "{1}"', $item, $filter.Comment)
        Write-LogFile ($log) 5
        if ($Script:NoPreview)
        {
            Set-ItemProperty -Path $item -Name Comment -Value $filter.Comment
        }
    }

    if ($type -eq "AccessControl")
    {
        $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name ConnectionType -Value {1}', $item, $filter.ConnectionType)
        Write-LogFile ($log) 5
        if ($Script:NoPreview)
        {
            Set-ItemProperty -Path $item -Name ConnectionType -Value $filter.ConnectionType
        }
        if (![string]::IsNullOrEmpty($filter.AccessGatewayFarm))
        {
            $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name AccessGatewayFarm -Value "{1}"',
                $item, $filter.AccessGatewayFarm)
            Write-LogFile ($log) 5
            if ($Script:NoPreview)
            {
                Set-ItemProperty -Path $item -Name AccessGatewayFarm -Value $filter.AccessGatewayFarm
            }
        }
        if (![string]::IsNullOrEmpty($filter.AccessCondition))
        {
            $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name AccessCondition -Value "{1}"',
                $item, $filter.AccessCondition)
            Write-LogFile ($log) 5
            if ($Script:NoPreview)
            {
                Set-ItemProperty -Path $item -Name AccessCondition -Value $filter.AccessCondition
            }
        }
    }

    $Script:FilterAdded = $true
}

<#
    .Synopsis
        Create new policy with the provided XML data.
#>

function Add-Policy
{
    param
    (
        [string]$root,
        [object]$policy
    )

    $scope = $root.SubString(5).Trim('\')
    $name = $policy.PolicyName
    $path = $root + $name

    $count = 0
    foreach ($s in $policy.Settings.Setting)
    {
        if ($IgnoredSettings -contains $s.Name)
        {
            continue
        }
        $ignore = $false
        foreach ($cat in ([string]$s.Path).Split('\'))
        {
            if ($IgnoredCategories -contains $cat)
            {
                $ignore = $true
                break
            }
        }
        if (!$ignore)
        {
            $count++
        }
    }

    if ($count -eq 0)
    {
        Write-LogFile ([string]::Format('Not importing policy "{0}" because it has no valid settings', $name)) 1 -isWarning
        $Script:IgnoredPolicies++
        return
    }

    Write-LogFile ([string]::Format('Importing {0} policy "{1}"', $scope, $name)) 1 $true
    if (Test-Path -Path $path)
    {
        Write-LogFile ([string]::Format("{0} policy {1} already exists, skipped", $scope, $name)) 2 $true -isWarning
        $Script:IgnoredPolicies++
        return
    }

    Write-LogFile ([string]::Format('Creating new {0} policy "{1}"', $scope, $name)) 2 $true
    Write-LogFile ([string]::Format('New-Item "{0}"', $path)) 3
    if ($Script:NoPreview)
    {
        [void](New-Item $path)
    }

    $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name Description -Value "{1}"', $path, $policy.Description)
    Write-LogFile ($log) 3
    if ($Script:NoPreview)
    {
        Set-ItemProperty -Path $path -Name "Description" -Value $policy.Description
    }

    $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name Enabled -Value {1}', $path, $policy.Enabled)
    Write-LogFile ($log) 3
    if ($Script:NoPreview)
    {
        Set-ItemProperty -Path $path -Name "Enabled" -Value $policy.Enabled
    }

    $log = [string]::Format('Set-ItemProperty -Path "{0}" -Name Priority -Value {1}', $path, $policy.Priority)
    Write-LogFile ($log) 3
    if ($Script:NoPreview)
    {
        Set-ItemProperty -Path $path -Name "Priority" -Value $policy.Priority
    }

    Write-LogFile ([string]::Format('Configure other settings for "{0}"', $name)) 3
    foreach ($s in $policy.Settings.Setting)
    {
        Enable-Setting $path $s
    }

    $Script:FilterAdded = $false
    Write-LogFile ([string]::Format('Configure object assignments for "{0}"', $name)) 3
    $policy.Filters.Filter | % { Add-Filter $path $_ }
    if ($Script:FilterAdded)
    {
        $log = [string]::Format('Object assignments for "{0}" imported. {1}', $name,
            "Please carefully review the object assignments to ensure the policy is applied properly")
        Write-LogFile $log 3 -isWarning
    }
    $Script:ImportedPolicies++
}

<#
    .Synopsis
        Import XenApp farm policies from a XML file.
    .Parameter XmlInputFile
        The name of the XML file that stores the output from a previous export action.
        The XML schema must conform to the schema as defined in the XSD file.
    .Parameter XsdFile
        The name of the XSD file for the XML file. If this parameter is not specified,
        the default XSD file is PolicyData.XSD. The XSD file is used to validate the
        syntax of the input XML file.
    .Parameter NoLog
        If this switch is specified, no logs are generated.
    .Parameter NoClobber
        If this switch is specified, do not overwrite existing log file. This switch is
        ignored if NoLog is specified or if LogFile is not specified.
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
    .Parameter Preview
        If this switch is specified, the policy data is read from the XML file but no
        policies are imported to the target site. The log file contains logging
        information about the commands to be used if actual import were executed. This
        is useful for administrators to see what will actually happen if real policies
        are imported.
    .Parameter NoDetails
        If this switch is specified, detailed messages about the progress of the script
        execution will not be sent to the console.
    .Parameter SuppressLogo
        Suppress the logo.
    .Description
        Use this cmdlet to import the policies that have been previously exported to a
        XML file by the Export-Policy cmdlet. This cmdlet must be run on a XenDesktop
        controller and must have the Citrix Group Policy PowerShell Provider snap-in
        installed on the local server. The administrator must have permissions to
        create new policies in the XenDesktop site.

        Although most of the XenApp 6.x policies are imported, many of them are no
        longer supported by XenApp/XenDesktop 7.5. Administrators should review all
        the policies being imported and all the ones not imported to ensure that the
        policy configuration for the target site is well defined.

        It's not necessary to explicitly load the Citrix PowerShell Snapins into the
        session, the script automatically loads the snapins.
    .Inputs
        A XML data file and optionally other parameters.
    .Outputs
        None
    .Link
        https://www.citrix.com/downloads/xenapp/sdks/powershell-sdk.html
    .Example
        Import-Policy -XmlInputFile .\MyPolicies.xml
        Import policies stored in the file 'MyPolicies.xml' located under the current
        directory. The PolicyData.Xsd file must be in the same directory. The log file
        is automatically generated and placed under the $HOME directory.
    .Example
        Import-Policy -XmlInputFile .\MyPolicies.xml -XsdFile ..\PolicyData.xsd -NoDetails
        Import policies stored in the file 'MyPolicies.xml' located in the current
        directory and use the XML schema file 'PolicyData.xsd' in the parent
        directory. Do not show the detailed information to the console, but everything
        will still be logged in the log file, which is under user's $HOME directory.
    .Example
        Import-Policy -XmlInputFile .\MyPolicies.xml -Preview -LogFile .\PolicyImport.log
        Import policies stored in the file 'MyPolicies.xml' located in the current
        directory but do not actually import anything by specifying the -Preview
        switch. Put the logs in the 'PolicyImport.log' file located in the current
        directory. The log file stores all the information about the commands used
        to create the policies.
#>

function Import-Policy
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="Explicit")]
        [string]$XmlInputFile,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [string]$XsdFile,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [switch]$NoLog,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [switch]$NoClobber,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [string]$LogFile,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [switch]$Preview,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [switch]$NoDetails,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
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

    $Script:NoPreview = !$Preview

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

    Write-LogFile ([string]::Format('Import-Policy Command Line:'))
    Write-LogFile ([string]::Format('    -XmlInputFile {0}', $XmlInputFile))
    Write-LogFile ([string]::Format('    -XsdFile {0}', $XsdFile))
    Write-LogFile ([string]::Format('    -LogFile {0}', $LogFile))
    Write-LogFile ([string]::Format('    -Preview = {0}', $Preview))
    Write-LogFile ([string]::Format('    -NoClobber = {0}', $NoClobber))
    Write-LogFile ([string]::Format('    -NoDetails = {0}', $NoDetails))

    if ([string]::IsNullOrEmpty($XsdFile))
    {
        $XsdFile = ".\PolicyData.xsd"
    }
    $homedir = Resolve-Path .
    $xmlPath = Resolve-Path $XmlInputFile
    $xsdPath = Resolve-Path $XsdFile

    Assert-XmlInput "PolicyData.xsd" $xmlPath $xsdPath

    Write-LogFile ("Loading Citrix Group Policy Provider Snap-in") 0 $true
    $p = [Reflection.Assembly]::LoadWithPartialName("Citrix.GroupPolicy.PowerShellProvider")
    if ($p -eq $null)
    {
        Write-Error ([string]::Format("{0}`n{1}",
                                      "Citrix Group Policy Provider Snapin is not installed",
                                      "You must have Citrix Group Policy Provider Snapin installed to use this script."))
        return
    }

    Write-LogFile ('Import-Module -Assembly [Reflection.Assembly]::LoadWithPartialName("Citrix.GroupPolicy.PowerShellProvider")') 1
    [void](Import-Module -Assembly ([Reflection.Assembly]::LoadWithPartialName("Citrix.GroupPolicy.PowerShellProvider")))

    Write-LogFile ("Mount Group Policy GPO") 0 $true
    Write-LogFile ("New-PSDrive -Name Site -Root \ -PSProvider CitrixGroupPolicy -Controller localhost") 1
    [void](New-PSDrive -Name Site -Root \ -PSProvider CitrixGroupPolicy -Controller localhost)

    Write-LogFile ("Turn off auto write back") 0 $true
    Write-LogFile ("(Get-PSDrive Site).AutoWriteBack = $false") 1
    if ($Script:NoPreview)
    {
        (Get-PSDrive Site).AutoWriteBack = $false
    }

    Write-LogFile ("Read XML file")
    Write-LogFile ([string]::Format('[xml]$policies = Get-Content "{0}"', $xmlPath)) 1
    [xml]$policies = Get-Content $xmlPath

    $TotalUserPolicies = 0
    $TotalComputerPolicies = 0
    $Script:IgnoredPolicies = 0
    $Script:ImportedPolicies = 0

    Write-LogFile ("Importing user policies") 0 $true
    try
    {
        foreach ($p in $policies.Policies.User.Policy)
        {
            Add-Policy "Site:\User\" $p
            $TotalUserPolicies++
        }
    }
    catch
    {
        Stop-Logging "User policy import aborted" $_.Exception.Message
        return
    }

    Write-LogFile ("Importing computer policies") 0 $true
    try
    {
        foreach ($p in $policies.Policies.Computer.Policy)
        {
            Add-Policy "Site:\Computer\" $p
            $TotalComputerPolicies++
        }
    }
    catch
    {
        Stop-Logging "Computer policy import aborted" $_.Exception.Message
        return
    }

    Write-LogFile ("Save changes")
    Write-LogFile ("(Get-PSDrive Site).Save()") 1
    if ($Script:NoPreview)
    {
        (Get-PSDrive Site).Save()
    }
    Write-LogFile ("Remove-PSDrive Site") 1
    Remove-PSDrive Site

    $broker = (Get-PSSnapin -Registered Citrix.Broker.Admin.V2 -ErrorAction SilentlyContinue)
    if ($broker -ne $null)
    {
        Import-Module (Get-PSSnapin -Registered Citrix.Broker.Admin.V2).ModuleName -Force
        Set-BrokerSiteMetadata -Name "PolicyImportToolVersion-f4ff73e5-8f40-4fa0-92db-ecdb1a878ac9" -Value $XAMigrationToolVersionNumber
    }

    # Summarize the activities, count number of policies imported, etc.
    Write-LogFile ([string]::Format('Total number of user policies : {0}', $TotalUserPolicies)) 0 $true
    Write-LogFile ([string]::Format('Total number of computer policies : {0}', $TotalComputerPolicies)) 0 $true
    Write-LogFile ([string]::Format('Number of policies ignored: {0}', $Script:IgnoredPolicies)) 0 $true
    Write-LogFile ([string]::Format('Number of policies imported: {0}', $Script:ImportedPolicies)) 0 $true

    Stop-Logging "Policy import completed"
}

# SIG # Begin signature block
# MIIZOwYJKoZIhvcNAQcCoIIZLDCCGSgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQURtORWKaFTuzxlB9DBOzsosvc
# i8ygghQGMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
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
# SIb3DQEJBDEWBBRZlT5ZcwzwhWVUyMjPawhKSE4r2DA8BgorBgEEAYI3AgEMMS4w
# LKAYgBYAQwBpAHQAcgBpAHgAIABGAGkAbABloRCADnd3dy5jaXRyaXguY29tMA0G
# CSqGSIb3DQEBAQUABIIBAEi91AWYNjElOGoO8vhh+orQmucd5Ox9jwK6jXFIGmL3
# Ugsf1pZcK8KJ1srjkXGvVI0QGl1p+08vtRjVbbQhwpulknHSjyXrsR+avb3/oid3
# qNAzFZMTFwTbpHGkWYhyURqqZsZ+o5Sg4PUCIFkBlj5q7260SSh9pwMj6phI3c5q
# QobDRyKhgxmYYSdfb9SbzNAk7ICuhyITYaz9y2hfFTthWKwOhWDMOVbuStvhf6kv
# if+HkOijCN0uUb8Tj+rNiVBy4JrLf+I6upp3gpGgU/GgtHKg1xn7xWRx/F2OW9bI
# 0zl0hgVtKAEKA0b0u6tTZkdzknbCNAoI+VTV5pxlCM6hggILMIICBwYJKoZIhvcN
# AQkGMYIB+DCCAfQCAQEwcjBeMQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50
# ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcg
# U2VydmljZXMgQ0EgLSBHMgIQDs/0OMj+vzVuBNhqmBsaUDAJBgUrDgMCGgUAoF0w
# GAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTQxMTI1
# MjAxNTI1WjAjBgkqhkiG9w0BCQQxFgQUfXVNxP6RuiCTctWuIx9eIhFCubswDQYJ
# KoZIhvcNAQEBBQAEggEAV7zzLwc2T2B30NgsF0INJ3zmZjqBOQ9sPf/dIVd4kYbM
# lOeqTDGXRn5eJCyCm1/9NKIlbpxDpwdHroQVM9rj5T14VPILYC5LZkMWW/9AdcBQ
# KaZmNq+Snj4ip9RGZuw7O8gJgkq1nRtowXZMqqYn10i9JZk6TnL8uIdiW8VizTKZ
# rerPktN3CscpK1/gKwGKtzORKLN2n2cDzvo2LmAE2x3puxMmQLbl+R21be7/bu+7
# 8/eVsBlKiKhFwpvLAm99j00vM5KfZqaLXDZPw49SzcCu7LvPehWaUVFfFZz2SMV0
# MMDce5AZZx8rZdDsXVzoia+c6tyR9+hTepRaPcKLfA==
# SIG # End signature block
