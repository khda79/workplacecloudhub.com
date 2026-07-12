
<#
.SYNOPSIS
    Version: 1.0
    Intune remediation script for Windows Hello Key Trust issue.
.DESCRIPTION
    This script:
    - Logs actions to C:\ProgramData\SmartM365\IntuneRemediation\Logs\Remediate-WindowsHello.log.
    - Stops NgcSvc service.
    - Takes ownership of NGC folder and deletes it.
    - Restarts NgcSvc.
    - Optionally forces dsregcmd /join if Hybrid Join is broken.
    - Runs silently and exits with code 0 for success.
.NOTES
    Run as SYSTEM via Intune.
#>

$ErrorActionPreference = "Stop"
$Scenario = "Reset-WindowsHello"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogPath = Join-Path -Path $LogRoot -ChildPath "$Scenario-Remediation.log"

function Write-SmartM365Log {
    param([string]$Message)
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    $ngcPath = "C:\\Windows\\ServiceProfiles\\LocalService\\AppData\\Local\\Microsoft\\NGC"
    $ngcService = Get-Service -Name NgcSvc -ErrorAction SilentlyContinue

    if ($ngcService -and $ngcService.Status -ne "Stopped") {
        Write-SmartM365Log "Stopping NgcSvc service."
        Stop-Service -Name NgcSvc -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path -LiteralPath $ngcPath) {
        Write-SmartM365Log "Taking ownership of NGC folder."
        takeown.exe /F $ngcPath /A /R | Out-Null
        icacls.exe $ngcPath /grant administrators:F /T | Out-Null
        Write-SmartM365Log "Deleting NGC folder."
        Remove-Item -LiteralPath $ngcPath -Recurse -Force -ErrorAction Stop
    }
    else {
        Write-SmartM365Log "NGC folder not found; nothing to delete."
    }

    if ($ngcService) {
        Write-SmartM365Log "Starting NgcSvc service."
        Start-Service -Name NgcSvc -ErrorAction SilentlyContinue
    }

    $dsreg = dsregcmd /status | Out-String
    if ($dsreg -notmatch "AzureAdJoined.*YES" -or $dsreg -notmatch "DomainJoined.*YES") {
        Write-SmartM365Log "Device join state is incomplete; requesting dsregcmd /join."
        dsregcmd /join | Out-Null
    }

    Write-SmartM365Log "RemediationCompleted"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    try { Write-SmartM365Log "RemediationFailed Message=$($_.Exception.Message)" } catch { Write-Output "LogWriteFailed=$($_.Exception.Message)" }
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBIorMPQduQ8GBT
# rp8kSRHb2VFu3QaxD7JqzDc4yzJ9raCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCF5Wa+cQUN4gJpCpaN
# pDuyPkWXFQfaiCeGBnzszKd8JTANBgkqhkiG9w0BAQEFAASCAYBiVuLhz/F7XNvx
# jnAsxZgauMlcjRLyY2M/ScV0K/hIm6Z2UPahJbZHaYkNVQCSb/Ax4su+TDINhirW
# eCZx3eNjqTNAGv38P1xA+QO57U3D+fU0In5S//i2cYoUNbXN2x8Fk6jHboDyhrBE
# Hkifm51esUZixk1t0pQsBiB/fUHD0z0l/zbwxZi45rHJVsFvSq1efVyCVSXBAMAd
# JqQ94jAqMaRi2tESlzKUl3MBW0bDIpP/p3Qabs6deBYEWqjAzV+z33sYIIrZMoFB
# pLEqITlRQKt7hlK0XtpD0BC5oReSFLyaFBegucVcM2KXbi5c25oZeoOmNyQDnGnx
# EEWGuF4jcMQDgOTrxn2jBgOJ6CkMZUziLTOLC4Okb355mWawTspkW+r283qSp2K8
# 2Nnjia7OjfMtBtbWcLnuBkiXybRe2DlZ2psMuuLSTky69gvLn6/t0sJBLbjFKLFV
# sAiZ9SVuTMzA41dc/Gj/sKiYa2hovQxb5z/rOBUoiXvelkQByhw=
# SIG # End signature block
