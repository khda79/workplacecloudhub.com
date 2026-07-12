# Name: SmartM365-Force-AllowOSUpgrade-Detection.ps1
# Version: 1.0
# Detection script - Check if AllowOSUpgrade is set correctly
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\OSUpgrade"
$regName = "AllowOSUpgrade"
$expectedValue = 1

if (Test-Path $regPath) {
    try {
        $actualValue = Get-ItemProperty -Path $regPath -Name $regName -ErrorAction Stop | Select-Object -ExpandProperty $regName
        if ($actualValue -eq $expectedValue) {
            Write-Output "Compliant"
            exit 0
        } else {
            Write-Output "Non-compliant: Value is $actualValue"
            exit 1
        }
    } catch {
        Write-Output "Non-compliant: Value not found"
        exit 1
    }
} else {
    Write-Output "Non-compliant: Registry path not found"
    exit 1
}
# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBIbb5Hix77y1vf
# XJli2KWkzR3cDhieZDFW/K54U9dkqqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDXU4IyTbBtgU9H+USi
# fzsFzFybiPqZqmGV7gEdRfzCgzANBgkqhkiG9w0BAQEFAASCAYBDOcpuUVy8kS02
# srftWqZmSUt+nxmmF4aVeiwGfPPlcbgTOi8iq2dZpXlMpMVIYJQYsEDp7RT8lJEu
# 0Jut0t8XoNWf9jRF+BywEdWziHF8eq89QKgu4lS7tLP9rIcuatquXx3LyN3s14jp
# VO+ucs8xbtuP2buS3ReMESPzmjnQG7kPPNHXSNrC4BnurjHfcsVCOyghnnu48bnA
# 1JachvrFYyrkzMgGqsbJccc6cHxWqJyU6l6CS2rQQQO5sPctNONybYsyFIFPdTAf
# i76OUSjm6/RrzxMUXY0XVNIs+d2raFLSIReiitgQyEBYQW548BIDA7SvaFf04Lfr
# Zzf7slh2BUca5p1IB6oSxHwY0OG6XTpa3s5DsKdwkNTNkU4a3ad0pBBlxopoJQUf
# d8DUte9YkZ2ykpjZqSJjgSR7YaUA7OzeWhvpgs+68hCh4J+8B3GEuw4jCVAZyw5U
# VBzt0AkY8kSHUTUCd2ABXOj5NFllXt5XxdefNkVl0tvs2ASyYVo=
# SIG # End signature block
