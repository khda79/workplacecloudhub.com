<#
.SYNOPSIS
    Version: 1.0
    Enables or repairs BitLocker protection for BitLocker-OSDrive-Protection.
.DESCRIPTION
    Validates TPM readiness, adds a recovery password protector when needed, enables BitLocker on the operating system volume with used-space-only encryption, resumes protection when suspended, and attempts to back up the recovery key to Entra ID.
#>
[CmdletBinding()]
param(
    [string]$MountPoint = 'C:'
)

$ErrorActionPreference = 'Stop'
$Scenario = 'BitLocker-OSDrive-Protection'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
function Write-SmartM365Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log 'Remediation started.'

    $tpm = Get-Tpm -ErrorAction Stop
    if (-not $tpm.TpmPresent -or -not $tpm.TpmReady) { throw 'TPM is not present or not ready. BitLocker cannot be enabled safely by this remediation.' }

    $volume = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction Stop
    if ($volume.ProtectionStatus -eq 'On' -and $volume.EncryptionPercentage -eq 100) {
        Write-SmartM365Log 'BitLocker is already fully enabled and protected.'
        exit 0
    }

    $recoveryProtector = $volume.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1
    if (-not $recoveryProtector) {
        $protector = Add-BitLockerKeyProtector -MountPoint $MountPoint -RecoveryPasswordProtector -ErrorAction Stop
        $recoveryProtector = $protector.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1
        Write-SmartM365Log 'Recovery password protector added.'
    }

    if ($volume.VolumeStatus -eq 'FullyDecrypted') {
        Enable-BitLocker -MountPoint $MountPoint -TpmProtector -UsedSpaceOnly -SkipHardwareTest -ErrorAction Stop
        Write-SmartM365Log 'BitLocker enablement started with TPM protector and used-space-only encryption.'
    }
    elseif ($volume.ProtectionStatus -ne 'On') {
        Resume-BitLocker -MountPoint $MountPoint -ErrorAction SilentlyContinue
        Write-SmartM365Log 'BitLocker protection resume requested.'
    }

    if ($recoveryProtector) {
        try {
            BackupToAAD-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $recoveryProtector.KeyProtectorId -ErrorAction Stop
            Write-SmartM365Log 'Recovery key backup to Entra ID requested.'
        }
        catch {
            Write-SmartM365Log "Recovery key backup to Entra ID failed or is not supported on this device: $($_.Exception.Message)"
        }
    }

    Write-SmartM365Log 'Remediation completed.'
    exit 0
}
catch {
    Write-SmartM365Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}


# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDHvLLgy7053Gci
# zuzR25rQ1a4PO5wk7wODETiimHXCzqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAubvEaXPFB5jzC+4TL
# Iv45sWcHvnKJS7B5f2tiqiSluDANBgkqhkiG9w0BAQEFAASCAYArHYmb+x8uZy8N
# vi+3A8mWvieN53Os1EVXSX4KDdqhRrhl3flQDW6IcWlhjYtYxM2jDMK12HuT5/02
# 7xrUKSryjuGDEmKMxTJzJXkoq4ckeEghouy4tJtdJv6aQ+qS4XWgX/pWm7d6SY2F
# uUrDvytOnv384gSs/oVm9UimGZLB53OSNvACvRmbvOcMQNsEb4ugJkKoLkHjrD0l
# DdaOj/U72riSd3qtbEcWkTArKeBbHDtKsR31ydJGAF5g6XJThL7Kgynh8oRfytVd
# kO0oYoC29LvG4BkwuO0OzuhCwtnSBX9JDCp4/NLAroM4NhxCgAkPqir2Rmn71Ie7
# /KoljEJu8UIKb3JaU3MCHJJzZm7WvPbsWRSrwQ/sIlA2bXWTMLzBCcy7QHV8H9QG
# yrpHhometGg/BnNAeSdvUrtOywY9kTPmRgTO6hZ7cmOcLPYih4OHcPeyaniiW0MF
# JjlXGOXccY8DzMpFJm+UyKRO+VhnyEmwKp6XMNJwQVuDbuKDVs8=
# SIG # End signature block
