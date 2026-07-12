<#
.SYNOPSIS
    Version: 1.0
    Repairs inconsistent WinHTTP proxy state for WinHTTP-Proxy.
.DESCRIPTION
    Reads the current WinHTTP proxy state and resets it to direct access when the configuration is unreadable or inconsistent. This does not change user-level browser proxy settings.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$Scenario = 'WinHTTP-Proxy'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
function Write-SmartM365Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log 'Remediation started.'
    $proxyOutput = netsh winhttp show proxy 2>&1
    $proxyText = ($proxyOutput -join ' ')
    Write-SmartM365Log "Current WinHTTP proxy output: $proxyText"
    if ($LASTEXITCODE -ne 0 -or ($proxyText -notmatch 'Direct access|Acc.s direct|Acces direct' -and $proxyText -notmatch 'Proxy Server|Serveur proxy')) {
        netsh winhttp reset proxy | Out-Null
        Write-SmartM365Log 'WinHTTP proxy reset to direct access.'
    }
    else {
        Write-SmartM365Log 'WinHTTP proxy configuration is already valid; no reset required.'
    }
    exit 0
}
catch {
    Write-SmartM365Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}


# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB+DYbR+h6+ZWyb
# cclGliA4ic4GUh6GcMBqMY8zIA3ZaaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB6L+L0JJL59zzlmtbQ
# dT4H1wocnVnY2ueh1U6FoAdFLjANBgkqhkiG9w0BAQEFAASCAYBP6pvVy+jPZKp9
# xpjF+yevyCZJf/sB2CoXXUrZ/fwyw/4FnsGkmfcqqziRyEsR88z/ptiVBqerasCD
# 9Idl5Uaae5/YH3wxeboLw6ZqVM9qJoFJ2PJDw17+Wx0/JYzQbXf72U4OqOjFBE4M
# gKqo/UD6suJbBaKAPMSQL0Fa6EaiqHiJslR5WfqA4SwoAOHY+ZywlH2LxwIos3gL
# OUezaFZKCpbpa4x11/ZaGPFJFBwDgUHkgxqqUspAp/bOqr5tFg6aQaVvcdd+q1Nh
# dwO3e8Ms8/8X3/jeA+75Ccx0ASTQnyL4bHYZy2Jb+mR6p5/e8FagxpVqHQ6OIk27
# zcdJ3V+Xa78YRCZZ1ZNywSQo3LnZIZmpZFcPh6KU+k/jTvtOxHnJX7b9QRNguALO
# S5i2oLfm8yjuOV15fr9Yu7Oepq8gxKOqb5tJSO1YuXCTj2hZ5ipSSrGeQBCUCKwd
# qIVBfXa4U3M77/0yrMZe1ETi9AyJN+ob040VDhWUJ8c9TCGCvg4=
# SIG # End signature block
