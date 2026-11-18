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

function ConvertTo-SmartM365MailWindowUtc {
    param([Parameter(Mandatory)][datetime]$LocalTime, [Parameter(Mandatory)][TimeZoneInfo]$Zone)
    # Same rules as Get-SmartM365ReportMailNotBeforeUtc: a skipped wall-clock time
    # moves to the first valid minute; a repeated one uses its later occurrence.
    $scheduledLocal = [datetime]::SpecifyKind($LocalTime, [DateTimeKind]::Unspecified)
    while ($Zone.IsInvalidTime($scheduledLocal)) { $scheduledLocal = $scheduledLocal.AddMinutes(1) }
    if ($Zone.IsAmbiguousTime($scheduledLocal)) {
        $offset = @($Zone.GetAmbiguousTimeOffsets($scheduledLocal) | Sort-Object | Select-Object -First 1)[0]
        return ([datetimeoffset]::new($scheduledLocal, $offset)).ToUniversalTime()
    }
    return [datetimeoffset]::new([TimeZoneInfo]::ConvertTimeToUtc($scheduledLocal, $Zone))
}

function Get-SmartM365ReportMailWindowState {
    <#
    .SYNOPSIS
        Evaluates a weekly business-hours window [NotBefore, NotAfter) on the listed local days.
    .OUTPUTS
        IsOpen and NextOpenUtc (equal to Now when the window is open).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetimeoffset]$Now,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string[]]$Days,
        [Parameter(Mandatory)][ValidatePattern('^(?:[01][0-9]|2[0-3]):[0-5][0-9]$')][string]$NotBefore,
        [Parameter(Mandatory)][ValidatePattern('^(?:(?:[01][0-9]|2[0-3]):[0-5][0-9]|24:00)$')][string]$NotAfter,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$TimeZoneId
    )
    $daySet = [Collections.Generic.HashSet[DayOfWeek]]::new()
    # Name matching only (no numeric values); works the same on Windows PowerShell 5.1.
    $names = [Enum]::GetNames([DayOfWeek])
    foreach ($day in $Days) {
        $name = @($names | Where-Object { $_ -eq ([string]$day).Trim() })
        if ($name.Count -ne 1 -or -not $daySet.Add([DayOfWeek]$name[0])) {
            throw 'Mail window days must be unique English day names.'
        }
    }
    $start = [timespan]::ParseExact($NotBefore, 'hh\:mm', [Globalization.CultureInfo]::InvariantCulture)
    $end = if ($NotAfter -eq '24:00') { [timespan]::FromHours(24) } else { [timespan]::ParseExact($NotAfter, 'hh\:mm', [Globalization.CultureInfo]::InvariantCulture) }
    if ($end -le $start) { throw 'Mail window NotAfter must be later than NotBefore.' }
    $zone = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
    $nowUtc = $Now.ToUniversalTime()
    $localNow = [TimeZoneInfo]::ConvertTime($Now, $zone)
    $time = $localNow.TimeOfDay
    if ($daySet.Contains($localNow.DayOfWeek) -and $time -ge $start -and $time -lt $end) {
        return [pscustomobject]@{ IsOpen = $true; NextOpenUtc = $nowUtc }
    }
    for ($offset = 0; $offset -le 7; $offset++) {
        $date = $localNow.Date.AddDays($offset)
        if (-not $daySet.Contains($date.DayOfWeek)) { continue }
        $openUtc = ConvertTo-SmartM365MailWindowUtc -LocalTime $date.Add($start) -Zone $zone
        if ($openUtc -gt $nowUtc) { return [pscustomobject]@{ IsOpen = $false; NextOpenUtc = $openUtc } }
    }
    throw 'Mail window next opening could not be calculated.'
}

Export-ModuleMember -Function Get-SmartM365ReportMailNotBeforeUtc, Get-SmartM365ReportMailWindowState

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
