# Copyright Citrix Systems, Inc.

Set-Variable -Name TimeStampFormat -Scope Script -Option Constant -Value "{0:d4}-{1:d2}-{2:d2} {3:d2}:{4:d2}:{5:d2}:{6:d3}"
Set-Variable -Name LogFileNameFormat -Scope Script -Option Constant -Value "XFarm{0:d4}{1:d2}{2:d2}{3:d2}{4:d2}{5:d2}-{6:X}.Log"
Set-Variable -Name LogFileName -Scope Script -Value $null
Set-Variable -Name EnableLogging -Scope Script -Value $false
Set-Variable -Name ShowProgress -Scope Script -Value $false

. .\Version.ps1

function Format-Message
{
    param
    (
        [string]$message,
        [int]$indent
    )

    $d = Get-Date
    $space = ""
    for ([int]$i = 0; $i -lt $indent; $i++)
    {
        $space += "  "
    }

    $t = [string]::Format($Script:TimeStampFormat, $d.Year, $d.Month, $d.Day, $d.Hour, $d.Minute, $d.Second, $d.Millisecond)
    return [string]::Format("{0} {1}{2}", $t, $space, $message)
}

[void]
function Start-Logging
{
    [CmdletBinding()]
    param
    (
        [switch] $noLogging,
        [switch] $noClobber,
        [string] $logfile,
        [switch] $showProgress = $false
    )

    if (-not $PSBoundParameters.ContainsKey('Verbose'))
    {
        $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
    }

    $message = Format-Message "Logging Started" 0
    Write-Verbose $message

    $Script:LogFileName = $null
    $Script:ShowProgress = $showProgress
    $Script:EnableLogging = !$noLogging
    if ($noLogging)
    {
        return
    }

    if ([string]::IsNullOrEmpty($logfile))
    {
        $d = Get-Date
        $r = (New-Object Random).Next()
        $Script:LogFileName = [string]::Format($Script:LogFileNameFormat, $d.Year, $d.Month, $d.Day, $d.Hour, $d.Minute, $d.Second, $r)
        if (Test-Path $HOME)
        {
            $Script:LogFileName = Join-Path $HOME $Script:LogFileName
        }
        else
        {
            $Script:LogFileName = ".\" + $Script:LogFileName
        }
    }
    else
    {
        $Script:LogFileName = $logfile
    }

    if (Test-Path -Path $Script:LogFileName)
    {
        if ($noClobber)
        {
            throw("Log file " + $Script:LogFileName + " already exits")
        }
        else
        {
            try
            {
                Remove-Item -Path $Script:LogFileName
            }
            catch
            {
                throw [string]::Format("You do not have sufficient permission to clear the log file: {0}", $Script:LogFileName)
            }
        }
    }

    # Write a header and also test if the log file is accessible.
    try
    {
        $message | Out-File -FilePath $Script:LogFileName -Append
    }
    catch
    {
        throw [string]::Format("You do not have sufficient permission to write the log file: {0}", $Script:LogFileName)
    }

    $XAMigrationToolVersionString | Out-File -FilePath $Script:LogFileName -Append
}

<#
    .Synopsis
        Log output to a file.
    .Parameter Message
        The message to be logged.
#>

[void]
function Write-LogFile
{
    [CmdletBinding()]
    param
    (
        [string] $message,
        [int] $indent = 0,
        [bool] $showProgress = $false,
        [switch] $isWarning = $false
    )

    if (-not $PSBoundParameters.ContainsKey('Verbose'))
    {
        $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
    }

    if ($isWarning)
    {
        $s = Format-Message ("WARNING: " + $message) $indent
        Write-Warning $message
    }
    else
    {
        $s = Format-Message $message $indent
    }

    if ($Script:EnableLogging -and ![string]::IsNullOrEmpty($Script:LogFileName))
    {
        $s | Out-File -FilePath $Script:LogFileName -Append
    }

    if (($showProgress -and $Script:ShowProgress) -and (!($isWarning -and ($WarningPreference -eq "Continue"))))
    {
        Write-Host ([string]::Format("{0}{1}", $space, $message))
    }

    Write-Verbose $s
}

[void]
function Stop-Logging
{
    [CmdletBinding()]
    param
    (
        [string] $message,
        [string] $exception = $null
    )

    if (![string]::IsNullOrEmpty($exception))
    {
        Write-LogFile $exception
        Write-Host $exception -ForegroundColor Red
    }

    Write-LogFile $message

    if (![string]::IsNullOrEmpty($exception))
    {
        Write-Host $message -ForegroundColor Red
    }
    else
    {
        Write-Host $message
    }

    if ($Script:ShowProgress)
    {
        if ($Script:EnableLogging -and ($Script:LogFileName -ne $null))
        {
            Write-Host ([string]::Format("Log has been saved to {0}", $Script:LogFileName))
        }
    }
}

[void]
function Print-Logo
{
    param
    (
        [bool] $nologo = $true
    )

    if (!$nologo)
    {
        Write-Host $XAMigrationToolVersionString
    }
}

Export-ModuleMember -Function Start-Logging
Export-ModuleMember -Function Stop-Logging
Export-ModuleMember -Function Write-LogFile
Export-ModuleMember -Function Print-Logo

# SIG # Begin signature block
# MIIZOwYJKoZIhvcNAQcCoIIZLDCCGSgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUIyHdYyk2DVSA7Q93Mrybig6k
# pBOgghQGMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
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
# SIb3DQEJBDEWBBSGap54rWOk2RHdIfquJLkag+C6MDA8BgorBgEEAYI3AgEMMS4w
# LKAYgBYAQwBpAHQAcgBpAHgAIABGAGkAbABloRCADnd3dy5jaXRyaXguY29tMA0G
# CSqGSIb3DQEBAQUABIIBAEjXbIQWOa3KxNmllTWzIlrbArw3Urve385tEN2zoztY
# AbkqVhMFcOW+XwaGXetsyuEVY/CCPcsH0f1R4bEMC9FWqndwRJV0gopKKMN2vHiH
# 5vAnxCez2TGkDagKyuJ24YW1+A7tv9olHCed/r2D8BvRNaQ+BJnbKmd4r1o9yCSy
# bgI83DdGfirrLL55CDzQFDskBdIonWcBEF9cAj6xcKd0IHy6J3plO7m8fJUBdwmc
# voecleDMezKeM+YZlMvFXOc7/BQB1TMyLKb7xZ+rj/pwq46MxgfyAMy9vDqSVc0O
# XqhxoguRg7G03bi4lFTAF9Q4JU02jeFKqIJmBi31bEmhggILMIICBwYJKoZIhvcN
# AQkGMYIB+DCCAfQCAQEwcjBeMQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50
# ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcg
# U2VydmljZXMgQ0EgLSBHMgIQDs/0OMj+vzVuBNhqmBsaUDAJBgUrDgMCGgUAoF0w
# GAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTQxMTI1
# MjAxNTM2WjAjBgkqhkiG9w0BCQQxFgQU2fNxi0c5Nzyi6l7aLXm1dtT5tqIwDQYJ
# KoZIhvcNAQEBBQAEggEAHM0BSvVKYW9NqXFfO2WPxyqWsh/qpJq8TQcmWCjpP9YR
# yG/UgUxC6QWB50qxWSFrLmVDBMUcebXNT+67S0l58+ZMufOhW7IBUY/MjSHxvSlX
# mOlT8CUSjHflCS60EkU1iaX97BPiCy/RN5bDe8Azhj1Sdn/8cTgKMghDBQtJSttR
# 3t7gm8z00fc8aZWkXogyFIivxzz5dvccB5jv/9QYKvRF4Yst0Tn4zPT4JEWqbBzD
# nA6Gfl7KdLiJ3WEW3nPizgbqfrffBBOPZD10oyedtm5UMxj46EbZFlybSzjNzg5y
# D36b19cfEqTBezNe32hwFZrHssNrFBfNQBBSNSaHrA==
# SIG # End signature block
