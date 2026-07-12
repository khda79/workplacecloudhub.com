# Name: SmartM365-WinHTTP-Proxy-Configuration-Detection.ps1
# Version: 1.0
# Description: Verifies whether the WinHTTP proxy configuration is readable and consistent

$ErrorActionPreference = "Stop"

try {
    # Retrieve WinHTTP proxy configuration
    $proxyOutput = netsh winhttp show proxy 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Output ("Failed to retrieve WinHTTP proxy configuration: " + ($proxyOutput -join " "))
        exit 1
    }

    $proxyText = ($proxyOutput -join "`n")

    # Direct access is a valid configuration
    if ($proxyText -match "Direct access|Acc.s direct|Acces direct") {
        Write-Output "WinHTTP proxy configuration is valid: direct access"
        exit 0
    }

    # Explicit proxy is also a valid configuration
    if ($proxyText -match "Proxy Server|Serveur proxy") {
        Write-Output "WinHTTP proxy configuration is valid: proxy server configured"
        exit 0
    }

    # Any other output is considered inconsistent or unexpected
    Write-Output ("WinHTTP proxy configuration is inconsistent or unexpected: " + ($proxyText -replace "`r?`n", " "))
    exit 1
}
catch {
    Write-Output ("Technical script error: " + $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBXpTJN1GzFFYLt
# fiPhpAjKR/CNBoRAal2PLKflzNJRpKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAbt+zGQcgGj0Q3/WKo
# fZMlibutcHlr8TyDZQELy1LD+TANBgkqhkiG9w0BAQEFAASCAYAmNsMAliiSbWb5
# eotoFMBJRNkn8SYAUvPbUk5yuiekc7WgaNI65BvwXxx8gJ8J4nm+iHVOGEeUgeOr
# +riPTw27W+K4LHLYtkKbPZ4LQ/4xLKcBkrdBRr/2sKBJ2VKAwtyD6w3Rz0h2if6U
# qv4uJ56Gd6HEjr2EQVKREJTnF25ckeuWyvjHXbvu9WH494GuI/LMI7LpjWLsUt3z
# hah/DImjGCBwCfqbumTi7IvKu7+3GcqPulfzP7KuLSBYXya/yjvGL8bYtSNZ7lus
# frcA5i9BGUjWfB/rGYkGM284ybHaD/p2kDkPeb0bdVoCzPDvS87cEH8n0BdsHO3I
# 8Mbp/vfxK8V7MuVUIMCUPo/uF2xX9fPf9bYHN35kCOJP3e9v6XTCXrpZGrc6qIaY
# Y/1TUV5eM1buhck5ek8jpcXaEG4dV3YfIy9uGzZAeks5iFtVYzgSkwbrZSicwtUY
# ewlx7/O7CnZ9ysruxglyny3tn0guhfN7wySLc9td2ZjiVvcRQKY=
# SIG # End signature block
