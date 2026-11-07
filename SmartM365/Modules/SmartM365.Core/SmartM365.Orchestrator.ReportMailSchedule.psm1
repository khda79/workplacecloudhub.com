<#
.SYNOPSIS
    Calculates the earliest permitted report-mail delivery time.
.DESCRIPTION
    Pure scheduling helper. It does not send, persist, or enable any mail policy.
#>
function Get-SmartM365ReportMailNotBeforeUtc {
    [CmdletBinding()]
    [OutputType([datetimeoffset])]
    param(
        [Parameter(Mandatory)][datetimeoffset]$Now,
        [Parameter(Mandatory)][ValidatePattern('^(?:[01][0-9]|2[0-3]):[0-5][0-9]$')][string]$NotBefore,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$TimeZoneId
    )

    $zone = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
    $localNow = [TimeZoneInfo]::ConvertTime($Now, $zone)
    $parts = $NotBefore.Split(':')
    $scheduledLocal = [datetime]::SpecifyKind(
        $localNow.Date.AddHours([int]$parts[0]).AddMinutes([int]$parts[1]),
        [DateTimeKind]::Unspecified
    )

    # A wall-clock time skipped by a spring transition is moved to the first
    # valid minute. If repeated in autumn, select the later occurrence so that
    # delivery is never earlier than the configured local wall-clock time.
    while ($zone.IsInvalidTime($scheduledLocal)) {
        $scheduledLocal = $scheduledLocal.AddMinutes(1)
    }
    if ($zone.IsAmbiguousTime($scheduledLocal)) {
        $offset = @($zone.GetAmbiguousTimeOffsets($scheduledLocal) | Sort-Object | Select-Object -First 1)[0]
        $scheduledUtc = ([datetimeoffset]::new($scheduledLocal, $offset)).ToUniversalTime()
    } else {
        $scheduledUtc = [datetimeoffset]::new([TimeZoneInfo]::ConvertTimeToUtc($scheduledLocal, $zone))
    }
    $nowUtc = $Now.ToUniversalTime()
    if ($nowUtc -ge $scheduledUtc) { return $nowUtc }
    return $scheduledUtc
}

Export-ModuleMember -Function Get-SmartM365ReportMailNotBeforeUtc

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCANtoyG70D+6vAP
# CFL0tNmtHLUuTqfnzInSpkWcAkuGkaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAtKjALUmxiq+IsGdY8g5RP
# OvO8uzlArX9BHQY6ldATSTANBgkqhkiG9w0BAQEFAASCAYAhmYAVgz5lF32J7cCT
# djacyXqD2aDHCORqNRmiDoZ+oi4/rSfnKvlS+NX/RKhJk9LbzbmRYzjStTDIMa8q
# r3Enk0jFTUNk/nxkCbDzhrsgGh9Haiyd0ubFNW89tu9sqSFrt/Ohlih69TQPF+xT
# jBwJ1Jdxr0MQ11IFQneZYD8P4N7gJIG4P2T5N0OZmqEnaMcNcONO3Ll/AeDaYfEd
# 3OvUKc6GMq5Tw9k1VLqUyJgQO4gTRHIM0qKqLvgEzVzOCSD/i2cjOd8nhuFlZQlF
# 7ZB9jojHDcu0TbMfGI5YPSqCxXRF7oGRAPsCaLCQ6HbfvdQ/j6MDo46/cj3sptiz
# xF2nBWwOreYnjhRzZB9iHRT6GFxNEuI30Qe9ty4T0LCwZ9plvWg8/nSFk+ycfyPM
# zfn4fiXQo/P5rDAU/xxzFUqqqeD0hu/KOaJ8Vo2WTGAzVQBe7pZ/KsilxbOH0EKs
# atfVMpWA90dLHZvpcdghCS3np0TluONdl1kLVoEBTNdtS5k=
# SIG # End signature block
