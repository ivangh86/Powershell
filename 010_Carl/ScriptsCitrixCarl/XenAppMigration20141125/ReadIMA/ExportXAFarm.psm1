# Copyright Citrix Systems, Inc.

$ErrorActionPreference = "Stop"

Set-Variable -Name TotalAppsExported -Scope Script -Value 0
Set-Variable -Name TotalAppsSkipped -Scope Script -Value 0

<#
    .Synopsis
        Create a single administrator node.
    .Parameter AdminObj
        A Citrix.XenApp.Commands.XAAdministrator object.
#>

[System.Xml.XmlElement]
function New-AdministratorNode
{
    param
    (
        [Citrix.XenApp.Commands.XAAdministrator] $adminObj
    )

    $node = New-XmlNode "Administrator" $null $adminObj.AdministratorName
    foreach ($p in ($adminObj | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
    {
        if (($p -ne "MachineName") -and ($adminObj.$p -ne $null) -and ($adminObj.$p -isnot [System.Array]))
        {
            [void]$node.AppendChild((New-XmlNode $p $adminObj.$p))
        }
    }

    if ($adminObj.AdministratorType -eq "Custom")
    {
        $privs = New-XmlNode "FarmPrivileges"
        $adminObj.FarmPrivileges | % { [void]$privs.AppendChild((New-XmlNode "FarmPrivilege" $_)) }
        [void]$node.AppendChild($privs)

        $folders = New-XmlNode "FolderPrivileges"
        foreach ($f in $adminObj.FolderPrivileges)
        {
            $x = New-XmlNode "FolderPrivilege"
            [void]$x.AppendChild((New-XmlNode "FolderPath" $f.FolderPath))
            $privs = New-XmlNode "FolderPrivileges"
            $f.FolderPrivileges | % { [void]$privs.AppendChild((New-XmlNode "FolderPrivilege" $_)) }
            [void]$x.AppendChild($privs)
            [void]$folders.AppendChild($x)
        }
        [void]$node.AppendChild($folders)
    }

    return $node
}

[System.Xml.XmlElement]
function New-ConfigLoggingNode
{
    Write-LogFile "Exporting farm configuration logging settings" 0 $true
    $configObj = Get-XAConfigurationLog
    $node = New-XmlNode "ConfigurationLogging"
    foreach ($p in ($configObj | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
    {
        if (($p -ne "MachineName") -and ($configObj.$p -ne $null) -and  ($configObj.$p -isnot [System.Array]))
        {
            [void]$node.AppendChild((New-XmlNode $p $configObj.$p))
        }
    }

    return $node
}

[System.Xml.XmlElement]
function New-ServerNode
{
    param
    (
        [object]$server
    )

    $node = New-XmlNode "Server" $null $server.ServerName
    foreach ($p in ($server | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
    {
        if (($p -ne "MachineName") -and ($server.$p -ne $null) -and ($server.$p -isnot [System.Array]))
        {
            [void]$node.AppendChild((New-XmlNode $p $server.$p))
        }
    }

    if (($server.IPAddresses -ne $null) -and ($server.IPAddresses.Count -gt 0))
    {
        $ips = New-XmlNode "IPAddresses"
        $server.IPAddresses | % { [void]$ips.AppendChild((New-XmlNode "IPAddress" $_)) }
        [void]$node.AppendChild($ips)
    }

    return $node
}

[System.Xml.XmlElement]
function New-LoadEvaluatorNode
{
    param
    (
        [object]$leObj
    )

    Write-LogFile ([string]::Format('Exporting load evaluator "{0}"', $leObj.LoadEvaluatorName)) 1 $true

    $node = New-XmlNode "LoadEvaluator" $null $leObj.LoadEvaluatorName
    foreach ($p in ($leObj | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
    {
        if (($p -ne "MachineName") -and ($leObj.$p -ne $null) -and ($leObj.$p -isnot [System.Array]))
        {
            [void]$node.AppendChild((New-XmlNode $p $leObj.$p))
        }
        if (($leObj.$p -ne $null) -and ($leObj.$p -is [System.Array]) -and ($leObj.$p.Count -gt 0))
        {
            $items = New-XmlNode $p
            if ($p -like "*Schedule")
            {
                $leObj.$p | % { [void]$items.AppendChild((New-XmlNode "TimeOfDay" $_)) }
            }
            elseif ($p -eq "IPRanges")
            {
                $leObj.$p | % { [void]$items.AppendChild((New-XmlNode "IPRange" $_)) }
            }
            elseif ($leObj.$p -is [System.Int32[]])
            {
                [void]$items.AppendChild((New-XmlNode "NoLoad" $leObj.$p[0]))
                [void]$items.AppendChild((New-XmlNode "FullLoad" $leObj.$p[1]))
            }
            if ($items.HasChildNodes)
            {
                [void]$node.AppendChild($items)
            }
        }
    }

    return $node
}

[System.Xml.XmlElement]
function New-LBPolicyNode
{
    param
    (
        [Citrix.XenApp.Commands.XAPolicy]$lbpObj
    )

    Write-LogFile ([string]::Format('Exporting load balancing policy "{0}"', $lbpObj.PolicyName)) 1 $true

    $node = New-XmlNode "LoadBalancingPolicy" $null $lbpObj.PolicyName
    $config = Get-XALoadBalancingPolicyConfiguration $lbpObj.PolicyName
    $filter = Get-XALoadBalancingPolicyFilter $lbpObj.PolicyName

    foreach ($p in ($lbpObj | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
    {
        if (($p -ne "MachineName") -and ($lbpObj.$p -ne $null) -and ($lbpObj.$p -isnot [System.Array]))
        {
            [void]$node.AppendChild((New-XmlNode $p $lbpObj.$p))
        }
    }

    foreach ($p in ($config | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
    {
        if (($p -ne "MachineName") -and ($p -ne "PolicyName") -and ($config.$p -ne $null) -and ($config.$p -isnot [System.Array]))
        {
            [void]$node.AppendChild((New-XmlNode $p $config.$p))
        }
    }

    if (($config.WorkerGroupPreferences -ne $null) -and ($config.WorkerGroupPreferences.Count -gt 0))
    {
        $wgps = New-XmlNode "WorkerGroupPreferences"
        $config.WorkerGroupPreferences | % { [void]$wgps.AppendChild((New-XmlNode "WorkerGroupPreference" $_)) }
        [void]$node.AppendChild($wgps)
    }

    foreach ($p in ($filter | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
    {
        if (($p -ne "MachineName") -and ($p -ne "PolicyName") -and ($filter.$p -ne $null) -and ($filter.$p -isnot [System.Array]))
        {
            [void]$node.AppendChild((New-XmlNode $p $filter.$p))
        }
        if (($filter.$p -ne $null) -and ($filter.$p -is [string[]]) -and (-not $p.EndsWith("Accounts")))
        {
            $s = New-XmlNode $p
            $w = $p.TrimEnd('s')
            if ($p.EndsWith("Addresses"))
            {
                $w = $w.TrimEnd('e')
            }
            $filter.$p | % { [void]$s.AppendChild((New-XmlNode $w $_)) }
            [void]$node.AppendChild($s)
        }
        if (($filter.$p -ne $null) -and ($filter.$p -is [System.Array]) -and ($p.EndsWith("Accounts")))
        {
            $s = New-XmlNode $p
            foreach ($q in $filter.$p)
            {
                $a = New-XmlNode $p.TrimEnd('s')
                foreach ($r in ($q | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
                {
                    if ($r -ne "MachineName")
                    {
                        [void]$a.AppendChild((New-XmlNode $r $q.$r))
                    }
                }
                [void]$s.AppendChild($a)
            }
            [void]$node.AppendChild($s)
        }
    }

    return $node
}

<#
    .Synopsis
        Create a node for a given application.
    .Parameter AppObj
        A Citrix.XenApp.Commands.XAApplication or Citrix.XenApp.Commands.XAApplicationReport object.
        This can also be a string, which should be the browser name of the application.
#>

[System.Xml.XmlElement]
function New-ApplicationNode
{
    param
    (
        [object] $appObj,
        [string] $iconDir,
        [bool] $embedIcon
    )

    if ($appObj -is [string])
    {
        if (IsNullOrWhiteSpace $appObj)
        {
            return $null
        }
        $appObj = Get-XAApplication $appObj
    }
    else
    {
        if (($appObj -isnot [Citrix.XenApp.Commands.XAApplication]) -and
            ($appObj -isnot [Citrix.XenApp.Commands.XAApplicationReport]))
        {
            return $null
        } 
    }

    $appName = $appObj.BrowserName
    $appNode = New-XmlNode "Application" $null $appName
    $appType = $appObj.ApplicationType
    $appData = Get-XAApplicationReport $appName

    # Common properties of primitivie types.
    foreach ($p in Get-XAApplicationParameter $appType)
    {
        if (($appObj.$p -ne $null) -and ($appObj.$p -isnot [System.Array]))
        {
            $propNode = New-XmlNode $p $appObj.$p
            [void]$appNode.AppendChild($propNode)
        }
    }

    # Accounts
    $accounts = New-XmlNode "Accounts"
    foreach ($acct in $appData.Accounts)
    {
        $node = New-XmlNode "Account"
        foreach ($p in ($acct | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
        {
            if (($p -ne "MachineName") -and ($acct.$p -ne $null) -and ($acct.$p -isnot [System.Array]))
            {
                [void]$node.AppendChild((New-XmlNode $p $acct.$p))
            }
        }
        [void]$accounts.AppendChild($node)
    }
    [void]$appNode.AppendChild($accounts)

    # Server list, which is not valid for Content and StreamedToClient applications.
    if (($appType -ne "Content") -and ($appType -ne "StreamedToClient"))
    {
        if (($appData.ServerNames -ne $null) -and ($appData.ServerNames.Count -gt 0))
        {
            $servers = New-XmlNode "Servers"
            $appData.ServerNames | % { [void]$servers.AppendChild((New-XmlNode "Server" $_)) }
            [void]$appNode.AppendChild($servers)
        }
        if (($appData.WorkerGroupNames -ne $null) -and ($appData.WorkerGroupNames.Count -gt 0))
        {
            $wgs = New-XmlNode "WorkerGroups"
            $appData.WorkerGroupNames | % { [void] $wgs.AppendChild(($w = New-XmlNode "WorkerGroup" $_)) }
            [void]$appNode.AppendChild($wgs)
        }
    }

    if (($appData.FileTypes -ne $null) -and ($appData.FileTypes.Count -gt 0))
    {
        $fts = New-XmlNode "FileTypes"
        foreach ($entry in $appData.FileTypes)
        {
            $x = New-XmlNode "FileType" $null $entry.FileTypeName
            foreach ($p in ($entry | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
            {
                if (($p -ne "MachineName") -and ($entry.$p -ne $null) -and ($entry.$p -isnot [System.Array]))
                {
                    [void]$x.AppendChild((New-XmlNode $p $entry.$p))
                }
                elseif (($entry.$p -ne $null) -and ($entry.$p -is [string[]]))
                {
                    $s = New-XmlNode $p
                    $entry.$p | % { [void]$s.AppendChild((New-XmlNode $p.TrimEnd('s') $_)) }
                    [void]$x.AppendChild($s)
                }
            }
            [void]$fts.AppendChild($x)
        }
        [void]$appNode.AppendChild($fts)
    }

    if (($appData.AccessSessionConditions -ne $null) -and ($appData.AccessSessionConditions.Count -gt 0))
    {
        $n = New-XmlNode "AccessSessionConditions"
        $appData.AccessSessionConditions | % { [void]$n.AppendChild((New-XmlNode "AccessSessionCondition" $_)) }
        [void]$appNode.AppendChild($n)
    }

    if (($appData.AlternateProfiles -ne $null) -and ($appData.AlternateProfiles.Count -gt 0))
    {
        $n = New-XmlNode "AlternateProfiles"
        foreach ($p in $appData.AlternateProfiles)
        {
            $t = New-XmlNode "AlternateProfile"
            [void]$t.AppendChild((New-XmlNode "ProfileLocation" $t.ProfileLocation))
            [void]$t.AppendChild((New-XmlNode "IPRange" $t.IPRange))
        }
        [void]$appNode.AppendChild($n)
    }

    if ($embedIcon)
    {
        $node = New-XmlNode "IconData" (Get-XAApplicationIcon $appName).EncodedIconData
        [void]$appNode.AppendChild($node)
    }
    else
    {
        $iconFile = (Join-Path $iconDir $appName) + ".txt"
        Write-LogFile ([string]::Format('Saving {0} icon data to file "{1}"', $appName, $iconFile)) 2
        (Get-XAApplicationIcon $appName).EncodedIconData | Out-File -Force $iconFile
        $n = New-XmlNode "IconFileName" (Join-Path (Split-Path -Leaf $iconDir) (Split-Path -Leaf $iconFile))
        [void]$appNode.AppendChild($n)
    }

    return $appNode
}

[System.Xml.XmlElement]
function New-ApplicationsNode
{
    param
    (
        [string] $xmlFile,
        [bool] $embedIcon,
        [int] $appLimit = 65536,
        [int] $skipApps = 0
    )

    $parent = Split-Path $xmlFile
    if ([string]::IsNullOrEmpty($parent))
    {
        $parent = ".\"
    }
    $parent = (Resolve-Path $parent).Path

    if (!$embedIcon)
    {
        $iconDir = Join-Path $parent (([IO.FileInfo]$xmlFile).BaseName + "-icons")
        Write-LogFile ([string]::Format('Creating folder for application icons: "{0}"', $iconDir)) 1
        [void](mkdir $iconDir -Force)
    }

    $appsNode = New-XmlNode "Applications"

    foreach ($appObj in Get-XAApplication)
    {
        if ($Script:TotalAppsExported -ge $appLimit)
        {
            break
        }
        if ($Script:TotalAppsSkipped -lt $skipApps)
        {
            Write-LogFile ([string]::Format('INFO: Skipping application "{0}"', $appObj.DisplayName)) 1 $true
            $Script:TotalAppsSkipped++
            continue
        }
        Write-LogFile ([string]::Format('Exporting application "{0}"', $appObj.DisplayName)) 1 $true
        $appNode = New-ApplicationNode $appObj $iconDir $embedIcon 
        if ($appNode -ne $null)
        {
            [void]$appsNode.AppendChild($appNode)
            $Script:TotalAppsExported++
        }
    }

    return $appsNode
}

[System.Xml.XmlElement]
function New-WorkerGroupNode
{
    param
    (
        [object]$worker
    )

    Write-LogFile ([string]::Format('Exporting worker group "{0}"', $worker.WorkerGroupName)) 1 $true
    $node = New-XmlNode "WorkerGroup" $null $worker.WorkerGroupName
    foreach ($p in ($worker | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
    {
        if (($p -ne "MachineName") -and ($worker.$p -ne $null) -and ($worker.$p -isnot [System.Array]))
        {
            [void]$node.AppendChild((New-XmlNode $p $worker.$p))
        }
    }

    if (($worker.ServerNames -ne $null) -and ($worker.ServerNames.Count -gt 0))
    {
        $servers = New-XmlNode "ServerNames"
        $worker.ServerNames | % { [void]$servers.AppendChild((New-XmlNode "ServerName" $_)) }
        [void]$node.AppendChild($servers)
    }

    if (($worker.ServerGroups -ne $null) -and ($worker.ServerGroups.Count -gt 0))
    {
        $groups = New-XmlNode "ServerGroups"
        $worker.ServerGroups | % { [void]$groups.AppendChild((New-XmlNode "ServerGroup" $_)) }
        [void]$node.AppendChild($groups)
    }

    if (($worker.OUs -ne $null) -and ($worker.OUs.Count -gt 0))
    {
        $ous = New-XmlNode "OUs"
        $worker.ServerNames | % { [void]$ous.AppendChild((New-XmlNode "OU" $_)) }
        [void]$node.AppendChild($ous)
    }

    return $node
}

[System.Xml.XmlElement]
function New-ZoneNode
{
    param
    (
        [object]$zone
    )

    Write-LogFile ([string]::Format('Exporting zone "{0}"', $zone.ZoneName)) 1 $true
    $node = New-XmlNode "Zone" $null $zone.ZoneName
    [void]$node.AppendChild((New-XmlNode "ZoneName" $zone.ZoneName))
    [void]$node.AppendChild((New-XmlNode "DataCollector" $zone.DataCollector))

    $servers = New-XmlNode "Servers"
    Get-XAServer -ZoneName $zone.ZoneName | % { [void]$servers.AppendChild((New-ServerNode $_)) }
    [void]$node.AppendChild($servers)

    return $node
}

<#
    .Synopsis
        Export XenApp farm configuration data to a XML file.
    .Parameter XmlOutputFile
        The name of the XML file that stores the output. The file name must have a .xml
        extension. It must not exist but if a path is given, the parent path must exist.
    .Parameter NoLog
        Do not generate log output.
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
    .Parameter EmbedIconData
        Include the icon data for applications in the XML file. If this switch is not
        specified, icon data is stored separately in files and the files are named using
        the browser name of the application. See the description for more detailed
        information about the icon data files.
    .Parameter NoClobber
        Do not overwrite an existing log file. This switch is applicable only if the
        LogFile parameter is specified.
    .Parameter NoDetails
        If this switch is specified, detailed messages about the progress of the script
        execution will not be sent to the console.
    .Parameter IgnoreAdmins
        Do not export administrators.
    .Parameter IgnoreApps
        Do not export applications.
    .Parameter IgnoreServers
        Do not export servers.
    .Parameter IgnoreZones
        Do not export zones.
    .Parameter IgnoreOthers
        Do not export configuration logging, load evaluators, worker groups, printer drivers,
        and load balancing policies.
    .Parameter AppLimit
        Export only the specified number of applications. If this value is 0, no applications
        are exported. The actual number of applications exported may be smaller than this
        limit.
    .Parameter SkipApps
        Skip the first specified number of applications.
    .Parameter SuppressLogo
        Suppress the logo.
    .Description
        Use this cmdlet to export the configuration data in a XenApp farm to a XML file.
        This cmdlet must be run on a XenApp controller and must have the Citrix XenApp
        Commands PowerShell snap-in installed on the local server.

        The data stored in the XML is organized in the same manner as what is displayed
        in the Citrix AppCenter.

        All data is retrieved using the Citrix XenApp commands.

        The user must have at least read access to all the objects in the farm.

        Application icon data is stored under a folder named by appending the string
        -icons to the base name of the XML file. For example, if the value of the XmlOutputFile
        parameter is FarmData.xml, then the folder FarmData-icons will be created to store
        the application icons. The icon data files under this folder are .txt files named
        using the browser name of the published application. Although the files are .txt
        files, the data stored is encoded binary icon data, which can be read by the
        import script to re-create the application icon.

        You can selectively export some of the farm objects and ignore other objects. For
        example, use the IgnoreZones switch to avoid exporting zone data. In case some of
        the farm objects cause some XenApp commands to fail, these switches can be used to
        work around those calls.

        Use the AppLimit and SkipApps parameters to fine tune your export. If you have large
        number of applications and the applications are to be imported to different delivery
        groups, by exporting selected applications to sepaparate XML files, you can avoid
        editing and specifying specific delivery group for each application in the XML file.
        See help for the Import-XAFarm command for more information about importing applications.
    .Inputs
        None
    .Outputs
        None
    .Link
        https://www.citrix.com/downloads/xenapp/sdks/powershell-sdk.html
    .Example
        Export-XAFarm -XmlOutputFile '.\MyFarmObjects.XML'
        Export all the farm objects and store the data in the file 'MyFarmObjects.XML' in the
        current directory. The log file is generated and located in the $HOME directory. During
        the export, progress of the script is displayed in the console.
    .Example
        Export-XAFarm -XmlOutputFile .\MyFarmObjects.XML -LogFile .\FarmExport.log
        Export all the farm objects and store the data in the file 'MyFarmObjects.XML' in the
        current directory. The log file is FarmExport.log and located in the current directory.
        During the export, progress of the script is displayed in the console.
    .Example
        Export-XAFarm -XmlOutputFile .\MyFarmObjects.XML -LogFile .\FarmExport.log -NoClobber
        Export all the farm objects and store the data in the file 'MyFarmObjects.XML' in the
        current directory. Store the log data in file FarmExport.log in the current directory
        but do not overwrite the file if it exists. During the export, progress of the script
        is displayed in the console.
    .Example
        Export-XAFarm -XmlOutputFile .\MajorObjects.XML -NoDetails -IgnoreZones -IgnoreOthers
        Export application, administrator, server objects and do not export other objects. The
        log file is generated and located in the $HOME directory. Do not display the progress
        of the script execution during the export.
    .Example
        Export-XAFarm -XmlOutputFile .\SomeAppObjects.XML -SkipApps 20 -AppLimit 100
        Export all farm objects but just some of the applications. The first 20 applications are
        not exported and limit the number of applications exported to 100. The applications are
        listed in random order, so there is no way to know which applications are exported. The
        applications are listed using the Get-XAApplication command.
#>
function Export-XAFarm
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="Explicit")]
        [ValidateScript({ Assert-XmlOutput $_ })]
        [string] $XmlOutputFile,
        [Parameter(Mandatory=$false)]
        [switch] $NoLog,
        [Parameter(Mandatory=$false)]
        [switch] $NoClobber,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [string] $LogFile,
        [Parameter(Mandatory=$false)]
        [switch] $EmbedIconData,
        [Parameter(Mandatory=$false)]
        [switch] $NoDetails,
        [Parameter(Mandatory=$false)]
        [switch] $IgnoreAdmins,
        [Parameter(Mandatory=$false)]
        [switch] $IgnoreApps,
        [Parameter(Mandatory=$false)]
        [switch] $IgnoreServers,
        [Parameter(Mandatory=$false)]
        [switch] $IgnoreZones,
        [Parameter(Mandatory=$false)]
        [switch] $IgnoreOthers,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [ValidateRange(0,65536)]
        [int] $AppLimit=65536,
        [Parameter(Mandatory=$false, ParameterSetName="Explicit")]
        [int] $SkipApps,
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

    Write-LogFile ([string]::Format('Export-XAFarm Command Line:'))
    Write-LogFile ([string]::Format('    -XmlOutputFile {0}', $XmlOutputFile))
    Write-LogFile ([string]::Format('    -LogFile {0}', $LogFile))
    Write-LogFile ([string]::Format('    -EmbedIconData = {0}', $EmbedIconData))
    Write-LogFile ([string]::Format('    -AppLimit {0}', $AppLimit))
    Write-LogFile ([string]::Format('    -SkipApps {0}', $SkipApps))
    Write-LogFile ([string]::Format('    -NoClobber = {0}', $NoClobber))
    Write-LogFile ([string]::Format('    -NoDetails = {0}', $NoDetails))
    Write-LogFile ([string]::Format('    -IgnoreAdmins = {0}', $IgnoreAdmins))
    Write-LogFile ([string]::Format('    -IgnoreApps = {0}', $IgnoreApps))
    Write-LogFile ([string]::Format('    -IgnoreServers = {0}', $IgnoreServers))
    Write-LogFile ([string]::Format('    -IgnoreZones = {0}', $IgnoreZones))
    Write-LogFile ([string]::Format('    -IgnoreOthers = {0}', $IgnoreOthers))

    Initialize-Xml

    $s = (Get-PSSnapin -Registered Citrix.XenApp.Commands -ErrorAction SilentlyContinue)
    if ($s -eq $null)
    {
        Write-Error ([string]::Format("{0}`n{1}",
            "The Citrix XenApp Commands PowerShell Snapin is not installed",
            "You must have the Citrix XenApp Commands PowerShell Snapin installed to use this script"))
        return
    }

    $s = (Get-PSSnapin Citrix.XenApp.Commands -ErrorAction SilentlyContinue)
    if ($s -eq $null)
    {
        $m = (Get-PSSnapin -Registered Citrix.XenApp.Commands).ModuleName
        Import-Module (Join-Path (Split-Path -Parent -Path $m) Citrix.XenApp.Commands.dll)
    }

    Write-LogFile "Exporting farm object"
    $farm = Get-XAFarm
    $root = New-XmlNode "Farm"
    [void]$root.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
    [void]$root.SetAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
    [void]$root.SetAttribute("xmlns", "XAFarmData.xsd")
    [void]$root.AppendChild((New-XmlNode "FarmName" $farm.FarmName))

    $TotalAdmins = 0
    $TotalServers = 0
    $TotalWorkerGroups = 0
    $TotalLoadEvaluators = 0
    $TotalLoadBalancingPolicies = 0
    $TotalPrinterDrivers = 0
    $TotalZones = 0
    $Script:TotalAppsSkipped = 0
    $Script:TotalAppsExported = 0

    if (!$IgnoreOthers)
    {
        [void]$root.AppendChild((New-ConfigLoggingNode))
    }
    else
    {
        Write-LogFile 'INFO: Configuration logging data not exported' 1 $true
    }

    if (!$IgnoreAdmins)
    {
        Write-LogFile "Exporting administrators" 0 $true
        $admins = New-XmlNode "Administrators"
        try
        {
            foreach ($a in (Get-XAAdministrator))
            {
                Write-LogFile ([string]::Format('Exporting administrator "{0}"', $a.AdministratorName)) 1 $true
                [void]$admins.AppendChild((New-AdministratorNode $a))
                $TotalAdmins++
            }
        }
        catch
        {
            Stop-Logging "Exporting administrators aborted, you may try to use the -IgnoreAdmins switch to work around this problem" $_.Exception.Message
            return
        }
        [void]$root.AppendChild($admins)
        Write-LogFile ([string]::Format('{0} administrators exported', $TotalAdmins)) 1 $true
    }
    else
    {
        Write-LogFile 'INFO: Administrators not exported' 1 $true
    }

    if (!$IgnoreApps)
    {
        Write-LogFile "Exporting applications" 0 $true
        try
        {
            [void]$root.AppendChild((New-ApplicationsNode $XmlOutputFile $EmbedIconData $AppLimit $SkipApps))
        }
        catch
        {
            Stop-Logging "Exporting applications aborted" $_.Exception.Message
            return
        }
    }
    else
    {
        Write-LogFile 'INFO: Applications not exported' 1 $true
    }

    if (!$IgnoreServers)
    {
        Write-LogFile "Exporting servers" 0 $true
        $servers = New-XmlNode "Servers"
        try
        {
            foreach ($s in (Get-XAServer))
            {
                Write-LogFile ([string]::Format('Exporting server {0}', $s.ServerName)) 1 $true
                [void]$servers.AppendChild((New-ServerNode $s))
                $TotalServers++
            }
        }
        catch
        {
            Stop-Logging "Exporting servers aborted, you may try to use the -IgnoreServers switch to work around this problem" $_.Exception.Message
            return
        }
        [void]$root.AppendChild($servers)
        Write-LogFile ([string]::Format('{0} servers exported', $TotalServers)) 1 $true
    }
    else
    {
        Write-LogFile 'INFO: Servers not exported' 1 $true
    }

    if (!$IgnoreOthers)
    {
        Write-LogFile "Exporting load evaluators" 0 $true
        $les = New-XmlNode "LoadEvaluators"
        try
        {
            foreach ($e in (Get-XALoadEvaluator))
            {
                [void]$les.AppendChild((New-LoadEvaluatorNode $e))
                $TotalLoadEvaluators++
            }
        }
        catch
        {
            Stop-Logging "Exporting load evaluators aborted, you may try to use the -IgnoreOthers switch to work around this problem" $_.Exception.Message
            return
        }
        [void]$root.AppendChild($les)
        Write-LogFile ([string]::Format('{0} load evaluators exported', $TotalLoadEvaluators)) 1 $true
    }
    else
    {
        Write-LogFile 'INFO: Load evaluators not exported' 1 $true
    }

    if (!$IgnoreOthers)
    {
        Write-LogFile "Exporting load balancing policies" 0 $true
        $policies = New-XmlNode "LoadBalancingPolicies"
        try
        {
            foreach ($p in (Get-XALoadBalancingPolicy))
            {
                [void]$policies.AppendChild((New-LBPolicyNode $p))
                $TotalLoadBalancingPolicies++
            }
        }
        catch
        {
            Stop-Logging "Exporting load balancing policies aborted, you may try to use the -IgnoreOthers switch to work around this problem" $_.Exception.Message
            return
        }
        [void]$root.AppendChild($policies)
        Write-LogFile ([string]::Format('{0} load balancing policies exported', $TotalLoadBalancingPolicies)) 1 $true
    }
    else
    {
        Write-LogFile 'INFO: Load balancing policies not exported' 1 $true
    }

    if (!$IgnoreOthers)
    {
        Write-LogFile "Exporting printer drivers" 0 $true
        $drivers = New-XmlNode "PrinterDrivers"
        try
        {
            foreach ($d in (Get-XAPrinterDriver))
            {
                Write-LogFile ([string]::Format('Exporting printer driver "{0}"', $d.DriverName)) 1 $true
                $node = New-XmlNode "PrinterDriver"
                foreach ($p in ($d | Get-Member -MemberType Property | Select-Object -ExpandProperty Name))
                {
                    if ($p -ne "MachineName")
                    {
                        [void]$node.AppendChild((New-XmlNode $p $d.$p))
                    }
                    [void]$drivers.AppendChild($node)
                }
                $TotalPrinterDrivers++
            }
        }
        catch
        {
            Stop-Logging "Exporting printer drivers aborted, you may try to use the -IgnoreOthers switch to work around this problem" $_.Exception.Message
            return
        }
        [void]$root.AppendChild($drivers)
        Write-LogFile ([string]::Format('{0} printer drivers exported', $TotalPrinterDrivers)) 1 $true
    }
    else
    {
        Write-LogFile 'INFO: Printer drivers not exported' 1 $true
    }

    if (!$IgnoreOthers)
    {
        Write-LogFile "Exporting worker groups" 0 $true
        $workers = New-XmlNode "WorkerGroups"
        try
        {
            foreach ($w in (Get-XAWorkerGroup))
            {
                [void]$workers.AppendChild((New-WorkerGroupNode $w))
                $TotalWorkerGroups++
            }
        }
        catch
        {
            Stop-Logging "Exporting worker groups aborted, you may try to use the -IgnoreOthers switch to work around this problem" $_.Exception.Message
            return
        }
        [void]$root.AppendChild($workers)
        Write-LogFile ([string]::Format('{0} worker groups exported', $TotalWorkerGroups)) 1 $true
    }
    else
    {
        Write-LogFile 'INFO: Worker groups not exported' 1 $true
    }

    if (!$IgnoreZones)
    {
        Write-LogFile "Exporting zones" 0 $true
        $zones = New-XmlNode "Zones"
        try
        {
            foreach ($z in (Get-XAZone))
            {
                [void]$zones.AppendChild((New-ZoneNode $z))
                $TotalZones++
            }
        }
        catch
        {
            Stop-Logging "Exporting zones aborted, you may consider to use the -IgnoreZones switch to work around this problem" $_.Exception.Message
            return
        }
        [void]$root.AppendChild($zones)
        Write-LogFile ([string]::Format('{0} zones exported', $TotalZones)) 1 $true
    }
    else
    {
        Write-LogFile 'INFO: Zones not exported' 1 $true
    }

    Write-LogFile "Saving XML file"
    Save-XmlData $root $XmlOutputFile

    Write-LogFile "" 0 $true
    Write-LogFile "Exporting completed successfully" 0 $true
    Write-LogFile ([string]::Format('{0} administrators exported', $TotalAdmins)) 0 $true
    Write-LogFile ([string]::Format('{0} servers exported', $TotalServers)) 0 $true
    Write-LogFile ([string]::Format('{0} applications exported', $Script:TotalAppsExported)) 0 $true
    if ($Script:TotalAppsSkipped -gt 0)
    {
        Write-LogFile ([string]::Format('First {0} applications skipped', $Script:TotalAppsSkipped)) 0 $true
    }
    Write-LogFile ([string]::Format('{0} worker groups exported', $TotalWorkerGroups)) 0 $true
    Write-LogFile ([string]::Format('{0} load evaluators exported', $TotalLoadEvaluators)) 0 $true
    Write-LogFile ([string]::Format('{0} load balancing policies exported', $TotalLoadBalancingPolicies)) 0 $true
    Write-LogFile ([string]::Format('{0} printer drivers exported', $TotalPrinterDrivers)) 0 $true
    Write-LogFile ([string]::Format('{0} zones exported', $TotalZones)) 0 $true

    Stop-Logging "XenApp 6.x farm export completed"
}

# SIG # Begin signature block
# MIIZOwYJKoZIhvcNAQcCoIIZLDCCGSgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU4FKh3KxiKlb+L3Y1FRZDq0k5
# yKGgghQGMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
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
# SIb3DQEJBDEWBBSQTvjGImzBVSp9YMrLN2+8ggTvrTA8BgorBgEEAYI3AgEMMS4w
# LKAYgBYAQwBpAHQAcgBpAHgAIABGAGkAbABloRCADnd3dy5jaXRyaXguY29tMA0G
# CSqGSIb3DQEBAQUABIIBAE8e0Rj4aF5VFCCEUf58J6XuBFzHo4jCP7/2IPN1YzYh
# x2j7uzgjslW/3bQ+lbDHa2j1XVyp5JQM3Jgbc8KI4AWBuBODa/iTT/gTgwKqqlHG
# D5qHDnevpg9Sd24hcxOdOvD/mYzP//obkl20uXLSKGba9nrlLMXxtJh6Qe2hvNQ2
# UrhrzyWEDwlO/HNGa+5tIv28XtG1F9AEq07oqWWIwuXcdD4/ssUH//rZVI78QXl4
# 2vmOCSD5KUOdTx2DvKqzdXGRTAV6cTzAbACaf69Zy+pgPonslskNtSMlNFpM754c
# gR6JoIeqLMj44DbJRP6Vzy51N/6aINdZ9oxqVRvpHTehggILMIICBwYJKoZIhvcN
# AQkGMYIB+DCCAfQCAQEwcjBeMQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50
# ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcg
# U2VydmljZXMgQ0EgLSBHMgIQDs/0OMj+vzVuBNhqmBsaUDAJBgUrDgMCGgUAoF0w
# GAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTQxMTI1
# MjAxNTI4WjAjBgkqhkiG9w0BCQQxFgQUrEqmFaPDuxQ3OBsaFy6al18+9eswDQYJ
# KoZIhvcNAQEBBQAEggEAIEmaJA7Cjy2GdSz3hYTcNQX0uY08w94l+r3l7sTqKuma
# n1QAp3np3kZwHWq36ivEOnsHL9gyCsv8u12OOCUARAfBbLLYM1uEzzWGNHPn8T6v
# ZwpBLruNup66dixiUEwRx4av/PSj0/79Xh7GFP8AKhvc0rK3N8DiR9K+vWvqdcVh
# 5oaj2lzT4arFvab8O1CqhliFw7UlELDeJw4g7eELp0sLvcAW4LBpbQmGuqVYHBt9
# 4bzkdalcNRnQyze0G6i4HNDSmodN33/nWpMymjqCIY3JJyC8EEGbC2i8Q2UlXp8x
# x5aBaF9Koc1ZmWXsCsJ4IIvqhC2tFkbj+4OKSJdjaA==
# SIG # End signature block
