<#
.SYNOPSIS
    Offline tests for report-mail schedule calculation; no transport is invoked.
#>
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.Orchestrator.ReportMailSchedule.psm1') -Force

$zone = 'Romance Standard Time'
$cases = @(
    @{ Now = '2026-11-03T06:29:00Z'; Time = '07:30'; Expected = '2026-11-03T06:30:00Z' }
    @{ Now = '2026-11-03T06:30:00Z'; Time = '07:30'; Expected = '2026-11-03T06:30:00Z' }
    @{ Now = '2026-11-03T07:00:00Z'; Time = '07:30'; Expected = '2026-11-03T07:00:00Z' }
    @{ Now = '2026-07-01T02:00:00Z'; Time = '07:30'; Expected = '2026-07-01T05:30:00Z' }
    @{ Now = '2026-03-29T00:30:00Z'; Time = '02:30'; Expected = '2026-03-29T01:00:00Z' }
    @{ Now = '2026-10-25T00:15:00Z'; Time = '02:30'; Expected = '2026-10-25T01:30:00Z' }
    @{ Now = '2026-10-25T00:45:00Z'; Time = '02:30'; Expected = '2026-10-25T01:30:00Z' }
)
foreach ($case in $cases) {
    $actual = Get-SmartM365ReportMailNotBeforeUtc -Now ([datetimeoffset]::Parse($case.Now)) -NotBefore $case.Time -TimeZoneId $zone
    $expected = [datetimeoffset]::Parse($case.Expected)
    if ($actual -ne $expected) { throw "Expected $expected, got $actual for $($case.Now) / $($case.Time)." }
}
foreach ($badTime in @('7:30', '24:00', '12:60', 'abc')) {
    $failed = $false
    try { $null = Get-SmartM365ReportMailNotBeforeUtc -Now ([datetimeoffset]::UtcNow) -NotBefore $badTime -TimeZoneId $zone } catch { $failed = $true }
    if (-not $failed) { throw "Invalid time '$badTime' was accepted." }
}
$failed = $false
try { $null = Get-SmartM365ReportMailNotBeforeUtc -Now ([datetimeoffset]::UtcNow) -NotBefore '07:30' -TimeZoneId 'Invalid/TimeZone' } catch { $failed = $true }
if (-not $failed) { throw 'Invalid time zone was accepted.' }
Write-Host 'PASS: 12 schedule checks; no mail sent.'

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB+De+D1QO0CM4M
# 5xVpnEACrADlggmEKfV87NU7XEkFlqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCY3Ty+zFzgCnua/zqJq1tF
# HxUv52yv/jkX9yDK7ORBczANBgkqhkiG9w0BAQEFAASCAYApnn5w5w/32/9GCJvL
# vDurx785rvBJIeb9+iRfK7yLwD0YKEpm3Ac5aKEHRKbw9d4qA3jhVlgwB89hAhkW
# W8trvp0Fi3ofkJG9qWgY5hC4YFtc3ditwCJ5KlBXh49TH0sUIODV5fOsTF0R1iMO
# Aj/dbXaHdyyFqC/X1TqDhN4hYkaXTTZH3qzapSpaxeh6YsJj/hav5p0t9T1QX/EZ
# iKDSb3wbqq5dSttQT+s8W5w537YSOVcbMA5tK9OdYdFjxkwbzqRYQ2hcPCKmBQ5D
# 7XzAiwjcB8ycPTxRaPE7Q0LPJ7MLeS74qS/ZrIkAOr8cjNTHjdCC+LVV6SPSTeJU
# 1hjXjKybiCpxFyYQD7h5o+wHblNA6fncbyEc+/7x7rUjzh59B/o8+3W9kOuyvZFH
# 8OBUOxCZFk3lpSxJDxM91ppZ/i8EgyHyG/ZDUaNkV+nY7VaRYtK/oX0EFASZ85UJ
# YIFfgCIefsMuJVpvuOctAhNXgPVtBGBoEqgnHm5v6B7+wq0=
# SIG # End signature block
