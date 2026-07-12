<#
    Name: SmartM365-WindowsUpdate-Service-And-Scan-Health-Detection.ps1
    Version: 1.0
    Description: Consolidated detection for Windows Update service health and recent scan activity.
#>

[CmdletBinding()]
param(
    [int]$MaxHoursSinceLastWindowsUpdateEvent = 24
)

$ErrorActionPreference = "Stop"
$ScriptName = "WindowsUpdate-Service-And-Scan-Health"
$RequiredServices = @("bits", "wuauserv", "dosvc", "cryptsvc", "UsoSvc")
$WindowsUpdateLogName = "Microsoft-Windows-WindowsUpdateClient/Operational"

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $healthyStates = New-Object System.Collections.Generic.List[string]

    foreach ($serviceName in $RequiredServices) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            $issues.Add("ServiceMissing=$serviceName")
            continue
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue

        if ($null -eq $serviceCim) {
            $issues.Add("ServiceCimUnavailable=$serviceName")
            continue
        }

        if ($serviceCim.StartMode -eq "Disabled") {
            $issues.Add("ServiceDisabled=$serviceName")
            continue
        }

        if ($service.Status -in @("StopPending", "PausePending", "Paused")) {
            $issues.Add("ServiceUnhealthy=$serviceName Status=$($service.Status)")
            continue
        }

        $healthyStates.Add("$serviceName=$($service.Status),StartupType=$($serviceCim.StartMode)")
    }

    $logInfo = Get-WinEvent -ListLog $WindowsUpdateLogName -ErrorAction Stop
    if (-not $logInfo.IsEnabled) {
        $issues.Add("WindowsUpdateEventLogDisabled")
    }
    else {
        $lastEvent = Get-WinEvent -LogName $WindowsUpdateLogName -MaxEvents 1 -ErrorAction SilentlyContinue

        if ($null -eq $lastEvent) {
            $issues.Add("WindowsUpdateEventLogEmpty")
        }
        else {
            $ageHours = (New-TimeSpan -Start $lastEvent.TimeCreated -End (Get-Date)).TotalHours

            if ($ageHours -ge $MaxHoursSinceLastWindowsUpdateEvent) {
                $issues.Add(("LastWUEventAgeHours={0:N1}" -f $ageHours))
            }
            else {
                $healthyStates.Add(("LastWUEventAgeHours={0:N1}" -f $ageHours))
            }
        }
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0}: remediation required. {1}" -f $ScriptName, ($issues -join "; "))
        exit 1
    }

    Write-Output ("{0}: healthy. {1}" -f $ScriptName, ($healthyStates -join "; "))
    exit 0
}
catch {
    Write-Output ("{0}: detection failed. Message={1}" -f $ScriptName, $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAvH7cgTLOFfzCJ
# TUATl6dMKJXrCX7kItE4fA5QdnarbqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCFKmii13iFAnulvXZg
# yp2+hQ87Cxmf4/17Jb5ZeUbItTANBgkqhkiG9w0BAQEFAASCAYBn601pkrQdfoRp
# pooPBa3EeFxMPCAlMxdJ7IVpoZthDOoTeA60x/ARxwVPz0yAn48vAzKdodg80m2D
# IXVqbAphTecY6vNgKdqf98+FECg/fys3/hV7ztayBIrMdfRFr6wPISNJo/GWvVeW
# mNn8WRtS6bDer8QC1PnZjauT3OClCJbNoMuFQlbkgJ6xFFobqnMN5NSGr/mA5Cuc
# LfH2oqojnlG/gVcHPWTDxPayC6vkg+NtXqy42mof/qCLlmfxPmxpoFI/oNMSRru5
# q3lKXFa7JwNPU6c0v1cdDLvjkHVjn8d6fWK3QobYemchf8mgIMJmlAbmpSWDoQcF
# EASzOkOZphDVguB4y+z8qygdlI8g9ofslUp9gYs8mU+FDZmvEscJEE59O4RaLX0L
# X2hMXavoTiOfZweFd8ARkwaIb23HbQgXT/fxbiPQUu4Bf9ztaMhdyX3S9I6sMPpO
# VOV4AM6Qsvfbr0zzeSS5gn/ggS6XlcWx2trOP7Ex3N6RCCn5EZ0=
# SIG # End signature block
