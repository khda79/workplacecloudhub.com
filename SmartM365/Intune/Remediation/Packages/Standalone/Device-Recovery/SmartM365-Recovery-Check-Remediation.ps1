# Name: SmartM365-Recovery-Check-Remediation.ps1
# Version: 1.0
$taskPath = "\Microsoft\Windows\Workplace Join"
$taskName = "Recovery-Check"

try {
    # Active la tâche
    Enable-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop

    # Exécute la tâche
    Start-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop
}
catch {
    Write-Output "Erreur lors de la remédiation : $($_.Exception.Message)"
    exit 1
}

exit 0

## Test
# powershell.exe -ExecutionPolicy Bypass -File .\SmartM365-Recovery-Check-Detection.ps1
# powershell.exe -ExecutionPolicy Bypass -File .\SmartM365-Recovery-Check-Remediation.ps1

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7tLlpV35V+FG0
# 9kt8vGuR7x+jSrRMTuW+9B/eSSqqmaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCsezkwKTgkjCLXshC5
# 2djiioNbKkB1LEYPiQpxoV/DUTANBgkqhkiG9w0BAQEFAASCAYBAEQrdFof6P7CF
# oZPxnqIqM5eGNCoKvMTn7lDPRc1Q13vQ26zUz9vF4ujXxAQjG26Beo5rdKBbw/lc
# 4yLsl6UZc8XjIVeFsczO9YgUCZUGspYSLuNYdQKrg+gyP3NsxgBexfTIQ88+XXCi
# LViOR4dlPuV2bYlbgZ4W6TLtTR5AWIKeU8l9H46M4SYbWGaJkoiQNnVrt5V+KftA
# fhKwa/DDvrYEbxebIEq0Jt/THzy3x+R4qo42gIAS9NkFZeWf64l7je3Ze4zcxChY
# PGJmbI6ERDXGo0Xx+fwQoGwDsUYR6nKAIKy5NIZnvOfTjQ1MJOXP9h8ZSeOdCoK8
# Q04IghlOSF8URxqqj2u4LetyiBtkGdbkjGsXTtDvE09pkG3PviA+j94aTiRla+FC
# 3K6z95VjxsQIH33MkkpbuIl61RnSssZQCLzfGVrWQ4ddPK46bvOHHsO41GJiyObX
# hDtpqQDtzG5OMnJIYS0r7XBCH4En6GCNe/RfwnLQry0yrjsEqlg=
# SIG # End signature block
