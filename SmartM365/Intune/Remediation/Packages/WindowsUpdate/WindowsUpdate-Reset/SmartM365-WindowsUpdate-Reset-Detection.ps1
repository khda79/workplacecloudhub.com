# Name: SmartM365-WindowsUpdate-Reset-Detection.ps1
# Version: 1.0
# Description: Detects Windows Update reset conditions such as legacy WSUS policies or recent Autopatch error 0x80244007.

$ErrorActionPreference = "Stop"

$ScriptName = "Detect-WindowsUpdate-Reset"
$Version = "1.0"
$ForceRemediation = $false

function Add-Issue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[string]]$Issues,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Issues.Contains($Message)) {
        $Issues.Add($Message)
    }
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $wuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    $wuAuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

    if (Test-Path -Path $wuPolicyPath) {
        $wuPolicy = Get-ItemProperty -Path $wuPolicyPath -ErrorAction Stop

        foreach ($name in @("WUServer", "WUStatusServer", "UpdateServiceUrlAlternate")) {
            if ($wuPolicy.PSObject.Properties.Name -contains $name) {
                $value = [string]$wuPolicy.$name

                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    Add-Issue -Issues $issues -Message "$name is configured."
                }
            }
        }
    }

    if (Test-Path -Path $wuAuPolicyPath) {
        $wuAuPolicy = Get-ItemProperty -Path $wuAuPolicyPath -ErrorAction Stop

        if ($wuAuPolicy.PSObject.Properties.Name -contains "UseWUServer" -and [int]$wuAuPolicy.UseWUServer -eq 1) {
            Add-Issue -Issues $issues -Message "UseWUServer=1."
        }
    }

    $reportingEventsPath = Join-Path -Path $env:windir -ChildPath "SoftwareDistribution\ReportingEvents.log"

    if (Test-Path -LiteralPath $reportingEventsPath) {
        $recentError = Select-String -Path $reportingEventsPath -Pattern "0x80244007|Windows Update Client failed to detect" -ErrorAction SilentlyContinue |
            Select-Object -Last 1

        if ($recentError) {
            Add-Issue -Issues $issues -Message "Recent Windows Update detection failure 0x80244007 found."
        }
    }

    if ($ForceRemediation) {
        if ($issues.Count -gt 0) {
            Write-Output ("{0} v{1}: forced remediation required. Indicators: {2}" -f $ScriptName, $Version, ($issues -join " | "))
        }
        else {
            Write-Output ("{0} v{1}: forced remediation required. No blocking indicator found." -f $ScriptName, $Version)
        }

        exit 1
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. {2}" -f $ScriptName, $Version, ($issues -join " | "))
        exit 1
    }

    Write-Output ("{0} v{1}: no remediation required." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBgeW5MBs6rLTs5
# kuutYKd5mK8LK4ET86smDU8zjO6lOKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCx5Cfo78KBLRLKgzR9
# ABpiXmLR/RF0RxlIUKExEUYNlTANBgkqhkiG9w0BAQEFAASCAYAnQIEbH2Pefwlc
# yvxXu3q85xZwkRQD+fJJRnxi/0Scun0MmqNHn+01SvhS7sJkzQPpOjhBnp4ahaES
# 6EeCn3qdD0/15NhhpMva1+OhzlQvwk5EIshz5Va1lff/fqXZ4MwMR/Ttr7IICcxZ
# 7bua3MN8SmBYTMt3EG5a2zrv3yGdzJE00+KNYny/b16r1PDpShu31LeQtA1kmMHL
# rNUWROVR6f/oszdoIwmzQZ9xN+hq/TaAxS0qyq2LvC/HthKRPKLKa2XuYIHVuqVB
# LeRXCsnBUV+WkcPVtR47EhQjaHyMoseI8n6ozWaVC5Y+9KSSBxW9Nd4bZ8FuMjSd
# kZrznFXVV20sDRS8XqHmecktpMN9l2kjZUM1jijkm14C5qV1Dxhz8rCYwfO6A53d
# kitI7wke+ayNQsaHt79jNOb19GUpRN+i1Zzj+nRgR4B4R95vdUK3cnTQDNgm/LBs
# rIDMFUWXHe9F2lL9spuXv0PIAYkwebftnEJ2i23Y70AWmhUHc3w=
# SIG # End signature block
