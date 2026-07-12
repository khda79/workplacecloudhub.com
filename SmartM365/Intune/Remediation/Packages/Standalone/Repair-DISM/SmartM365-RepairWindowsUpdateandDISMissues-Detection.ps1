# Name: SmartM365-RepairWindowsUpdateandDISMissues-Detection.ps1
# Version: 1.0
# Detection Script for Windows Update and DISM issues

$RemediationNeeded = $false

# Check Windows Update service
$WUService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
if ($WUService.Status -ne 'Running') {
    Write-Output "Windows Update service is not running."
    $RemediationNeeded = $true
}

# Check DISM health
$DismLog = "$env:windir\Logs\DISM\dism.log"
$DismError = $false
if (Test-Path $DismLog) {
    $DismError = Select-String -Path $DismLog -Pattern "error|failed" -SimpleMatch
    if ($DismError) {
        Write-Output "DISM log contains errors."
        $RemediationNeeded = $true
    }
}

# Check CBS log for update errors
$CBSLog = "$env:windir\Logs\CBS\cbs.log"
$CBSError = $false
if (Test-Path $CBSLog) {
    $CBSError = Select-String -Path $CBSLog -Pattern "error|failed" -SimpleMatch
    if ($CBSError) {
        Write-Output "CBS log contains errors."
        $RemediationNeeded = $true
    }
}

# Check Windows Update cache
$WUCache = "$env:windir\SoftwareDistribution\Download"
if (Test-Path $WUCache) {
    $CacheFiles = Get-ChildItem -Path $WUCache -Recurse -ErrorAction SilentlyContinue
    if ($CacheFiles.Count -gt 1000) {
        Write-Output "Windows Update cache contains excessive files."
        $RemediationNeeded = $true
    }
}

# Output for Intune Remediation
if ($RemediationNeeded) {
    Write-Output "Remediation required."
    exit 1
} else {
    Write-Output "No remediation required."
    exit 0
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDHjyKtzCKIFy1s
# GbKe7HmrriN2dq6XAVkME3litivK06CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBMKnfs0B3+z/Z+C4c1
# QgNnlhsJ1zUmn9ZP5tGHARCJXDANBgkqhkiG9w0BAQEFAASCAYAVVURH8Nszyc83
# bwFz2H3bkADCWtkiPgWSivxYWguQf3+XCyxtEDTc4T6tFeUQYuh7fKlSLGvsBwlX
# wohMSaZhu1HmSkQyRflqOjNG2NjrA7ENPbfVRxoe/qwupBycox0U46JRY5WQ7UcC
# gRlejGiGsXfF5lUQ8UuT8W7R5FWsDzjA2WUejBFubS5de80tSPF44YLgOEo23YeZ
# BlhviAa4I/Ayj0E6W1BVLK2K2elORfxBPJ0+0aDLPDYWkpkd+cwVeQdJpU9JQaAs
# ODG+de1pg4TrYqK1zoKjsRLggItrf4Cjc0qn+9TwxFmSHhwOXOyysjwWOxQkwyUn
# aWC1pPvRK/Lb5gHDf1MJFfjZ42hW9GV0BhtU1CXSH0oIWYH3PJnogESUPLg3TK+e
# SBlNhzfJ+2iJL3FHqQY8+uujOiSsDtBdXCMKZaoze3axBP4mAD/e0PvghvsYbPHG
# Cz88msnrB5G8O7W5qh7yd60GRG9dwZHTLncyx5BBg7bbUXh7Y0I=
# SIG # End signature block
