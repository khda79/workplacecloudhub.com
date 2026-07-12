
<#
.SYNOPSIS
    Version: 1.0
    Intune detection script for Windows Hello for Business with multilingual AD connectivity check.
.DESCRIPTION
    This script:
    - Detects if Windows Hello is configured (NGC folder or registry keys).
    - Checks Hybrid Join status using dsregcmd.
    - Dynamically checks AD connectivity (supports English and French outputs).
    Exit codes:
      0 = Healthy or not applicable.
      1 = Windows Hello configured but AD unreachable, or technical error.
.NOTES
    Run as SYSTEM via Intune.
#>

$ErrorActionPreference = "Stop"
$Scenario = "Reset-WindowsHello"

function Write-DetectionResult {
    param([string]$Message)
    Write-Output "$Scenario $Message"
}

# Check NGC folder
$ngcPath = "C:\\Windows\\ServiceProfiles\\LocalService\\AppData\\Local\\Microsoft\\NGC"
$ngcConfigured = ((Test-Path -LiteralPath $ngcPath) -and ((Get-ChildItem -LiteralPath $ngcPath -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1).Count -gt 0))

# Check registry for PIN presence
$pinConfigured = $false
try {
    $regPath = "HKLM:\\SOFTWARE\\Microsoft\\PassportForWork\\"  # Windows Hello for Business key
    $keys = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
    if ($keys) {
        $pinConfigured = $true
    }
} catch {
    $pinConfigured = $false
}

$helloConfigured = $ngcConfigured -or $pinConfigured

# Check Hybrid Join status
$dsreg = dsregcmd /status | Out-String
$domainJoined = $dsreg -match "DomainJoined.*YES"
$azureJoined = $dsreg -match "AzureAdJoined.*YES"

# Dynamically check AD connectivity (supports FR and EN output)
$adReachable = $false
try {
    $nltestResult = nltest /dsgetdc: | Out-String
    if ($nltestResult -match "DC:" -or $nltestResult -match "Domain Name" -or 
        $nltestResult -match "Nom du domaine" -or $nltestResult -match "Contrôleur de domaine") {
        $adReachable = $true
    }
} catch {
    $adReachable = $false
}
# Determine exit code
if (-not $helloConfigured) {
    Write-DetectionResult "Status=NotApplicable Reason=WindowsHelloNotConfigured"
    exit 0
}

if (-not $domainJoined) {
    Write-DetectionResult "Status=Healthy Reason=WindowsHelloConfiguredWithoutHybridJoin DomainJoined=$domainJoined AzureAdJoined=$azureJoined"
    exit 0
}

if ($adReachable) {
    Write-DetectionResult "Status=Healthy Reason=WindowsHelloConfiguredAndADReachable DomainJoined=$domainJoined AzureAdJoined=$azureJoined"
    exit 0
}

Write-DetectionResult "Status=RemediationRequired Reason=WindowsHelloConfiguredButADUnreachable DomainJoined=$domainJoined AzureAdJoined=$azureJoined"
exit 1

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB13KAY29Iljxa9
# jBqyeNhS9UphoKiKo+2fzD39GNlbEKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAsZDFwomNX7lxsWowO
# c4pPbjZfIu3UT0gBLKm5hEUjOjANBgkqhkiG9w0BAQEFAASCAYCo/oS5gyO+vSP3
# BLCSph1NWqFQWLk//qH0dejCW7P+ua8UDTzqIie78Y5YVa00ir9emLliSWYV8q7x
# +OVJ7KR+iVtuzcSr1GLNoqQshNb3yo5tn5lc+mCxI5b0lf62UIy/1glLhMxHfWwv
# JntDeWHcJagQ3+2RSV34X3SiTx1ip0LKf+EyTRA9knS4I3iHG9sATg8UBPphR0zx
# fFq/FzCpQvQ45u2Ir8v3rsCUAyVZoSQhnUmwBPyUJ6MS8Xj5S72PhN9+QHLPrhKy
# qBbhMshRReNR9mnEfb974LaT5mstMvjUSaQjBC8IXVa+9hx6z9Hv1MTcOyq23stR
# 1zDMK8Ae5ubdIQ6hhmOy39l6nBDzCj7U9PbMr7nfJNA/jcBqrt8kqHVdneQZ+i1X
# DCIf9lslvXh7OaW8aVOIHThAXzLk0+L7pFd8A2+hetOZJwfhlFPRJzXQim1TibXl
# XDIxudyguY6FL/ukS+0THTMjA4LcA+Vtdka7D/Nnugsego9V5Bg=
# SIG # End signature block
