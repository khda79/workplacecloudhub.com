<#
.SYNOPSIS
    Detects whether the system drive is below the managed free-space threshold.

.VERSION
    1.6
#>
# Name: SmartM365-Disk-Space-Cleanup-Detection.ps1
# Version: 1.6
# Description: Detects whether the system drive is below the managed free-space threshold.

$ErrorActionPreference = "Stop"

$ScriptName = "SmartM365-Disk-Space-Cleanup-Detection"
$Version = "1.6"
$MinimumFreeSpaceGB = 50
$Windows10Only = $true

function ConvertTo-SingleLineValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return "Unknown"
    }

    $text = [string]$Value
    $text = $text -replace '[\r\n\t]+', ' '
    $text = $text -replace '\s{2,}', ' '
    return $text.Trim()
}

function Write-IntuneResult {
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Values)

    $parts = New-Object System.Collections.Generic.List[string]

    foreach ($key in $Values.Keys) {
        $parts.Add(("{0}={1}" -f $key, (ConvertTo-SingleLineValue -Value $Values[$key])))
    }

    Write-Output ($parts -join " ")
}

function Get-SystemDriveFreeSpaceGB {
    try {
        $systemDrive = $env:SystemDrive

        if ([string]::IsNullOrWhiteSpace($systemDrive)) {
            return $null
        }

        $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction Stop

        if ($null -eq $drive -or $drive.DriveType -ne 3) {
            return $null
        }

        return [math]::Round(($drive.FreeSpace / 1GB), 2)
    }
    catch {
        return $null
    }
}

try {
    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

    if ($Windows10Only -and $operatingSystem.Caption -notmatch "Windows 10") {
        Write-IntuneResult -Values ([ordered]@{ Status = "NotApplicable"; Reason = "NotWindows10"; Script = $ScriptName; Version = $Version })
        exit 0
    }

    $freeSpaceGB = Get-SystemDriveFreeSpaceGB

    if ($null -eq $freeSpaceGB) {
        Write-IntuneResult -Values ([ordered]@{ Status = "RemediationRequired"; Reason = "FreeSpaceUnknown"; RequiredFreeSpaceGB = $MinimumFreeSpaceGB; Script = $ScriptName; Version = $Version })
        exit 1
    }

    if ($freeSpaceGB -ge $MinimumFreeSpaceGB) {
        Write-IntuneResult -Values ([ordered]@{ Status = "Ready"; FreeSpaceGB = $freeSpaceGB; RequiredFreeSpaceGB = $MinimumFreeSpaceGB; Script = $ScriptName; Version = $Version })
        exit 0
    }

    Write-IntuneResult -Values ([ordered]@{ Status = "RemediationRequired"; FreeSpaceGB = $freeSpaceGB; RequiredFreeSpaceGB = $MinimumFreeSpaceGB; Script = $ScriptName; Version = $Version })
    exit 1
}
catch {
    Write-IntuneResult -Values ([ordered]@{ Status = "Error"; Script = $ScriptName; Version = $Version; Message = $_.Exception.Message })
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCsymxq798Gkey+
# jxRX6PS1FFnpkW4wy3P+uyo3I3mkQqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC/0X4WHJxaNJ8LJGVC
# nGE1DW6VKVpFh2U130gjoz7QkTANBgkqhkiG9w0BAQEFAASCAYAOxdwXUPgTC22o
# qarApWwpl1kIy4A9E393xK8uoZBdhi6dVA7+dACmczuc5byFf1AJ8Knzuvp5vivQ
# bJeguBHvF7mizslvchYkAoqMv/2V3+imLprm82eL5R1aQSXzTIYnTW87Pvtd+aoF
# cEyyofqAu3VKjSyGF5iLt/d3prh70hcN23XCgmtS6MVwIbvxCD/EPo1n2lhD7uKV
# a917PiQE7w7cUDCZPVqU4RHIO2IpYumyxiqH8yBSClAYtNeqaRsliPMPIfxvy9mw
# 3xKvBiqcrrXJ09IjiTriMmuwxk8IfJiu1Cr6YnsoEPTHhoBlsLSajShoz6O5UJRa
# xbdEWYKNmePxHv59acDz2vVdH79zwM3uNaiCYJA30jeREu6+G+PxdrT6Sl7hDy/0
# nLTsVC3eN5Ufnu5EDdIgf+K9QuYD3jNzOPVPUU0tjkF21VbiVXmpelp7dxk+EYpR
# hGq6lsudUQmdlWrLd2SVdDKSDIKENykMDBAuMuSQlbkQs6Gba9c=
# SIG # End signature block
