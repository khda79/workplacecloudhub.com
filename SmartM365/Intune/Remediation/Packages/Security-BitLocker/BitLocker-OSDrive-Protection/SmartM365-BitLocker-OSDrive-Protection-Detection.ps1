<#
.SYNOPSIS
    Version: 1.0
    Detects BitLocker protection health on the operating system drive.
.DESCRIPTION
    Validates that the operating system drive exists, BitLocker protection is enabled,
    encryption is complete, and a recovery password protector is present.
#>
[CmdletBinding()]
param(
    [string]$MountPoint = 'C:'
)

$ErrorActionPreference = 'Stop'

try {
    $volume = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction Stop

    if ($null -eq $volume) {
        Write-Output "Status=BitLockerVolumeNotFound"
        exit 1
    }

    $issues = New-Object System.Collections.Generic.List[string]

    if ($volume.ProtectionStatus -ne 'On') {
        $issues.Add("ProtectionStatus=$($volume.ProtectionStatus)")
    }

    if ($volume.EncryptionPercentage -lt 100) {
        $issues.Add("EncryptionPercentage=$($volume.EncryptionPercentage)")
    }

    $recoveryProtector = $volume.KeyProtector |
        Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } |
        Select-Object -First 1

    if (-not $recoveryProtector) {
        $issues.Add('RecoveryPasswordProtector=Missing')
    }

    if ($issues.Count -gt 0) {
        Write-Output "Status=BitLockerNonCompliant"
        Write-Output ("Issues=" + ($issues -join '; '))
        exit 1
    }

    Write-Output "Status=BitLockerProtected"
    Write-Output "MountPoint=$MountPoint"
    Write-Output "EncryptionPercentage=$($volume.EncryptionPercentage)"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAF5xjl8uFnFiTR
# Alyj6l5+o7Qt/LtGHwpoHxgNZTbhhqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDe0oaFoW8gpPOP0M/L
# 5TXUFLbdeNa063U2NBMuHn59BjANBgkqhkiG9w0BAQEFAASCAYA6WSpUvu7Pj22i
# jn/yI8y7p/NKha2DLpFBNQvkUHbJz0JVWrY7FhCjFgQ0AlLP8zKD21Hckne1zEdr
# bYUwFCsUjmShR2zTquNcUMUpHDZ6OKYrdxUa/L37aIN5XPUrZHFW+DQGOpFn6YCl
# DGQgV4JhH1HBvihNAq0+o4QWqp93/MYPyj2KAJaY1fagd8jRDaQgqADwWw8XNUy1
# YL0R2JmV6fuGRBqSHpJ6MWkpeuZJ3OwJ5S/lS33VA9TEBzNIRWLfHV5xSzjCKQ2M
# ELcv3ijpF3rZMWvorp+pkW++LxvnJPaoZ9bc4Nj0Pmc9ammjp7x9Hp8U04klPv6I
# /vSHRoe0w8WpTncnumjK2PhDiStSNZ+6USjSiwvIByGXO7UqaXwHdw8bUh0yEGPW
# DZTdOfiM4oUf9RGR9K+AfxVz1V/GMEQ/S2AeCAtXi9v9oJQ+tatqjYHzpyRMfdKd
# vK26VovzzpccM22m6NrWmY8izK12CLLpi6+GwcTNfL8cuVuSmCs=
# SIG # End signature block
