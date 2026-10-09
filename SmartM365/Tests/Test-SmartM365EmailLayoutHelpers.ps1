[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$moduleFolder = Join-Path $PSScriptRoot '..\Modules\SmartM365.Core'
$manifest = if ($PSVersionTable.PSEdition -eq 'Desktop') { Join-Path $moduleFolder 'Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1' } else { Join-Path $moduleFolder 'SmartM365.Core.psd1' }
Import-Module $manifest -Force -ErrorAction Stop
$cards = @(
    [pscustomobject]@{Label='Objects';Value='12';Detail='collected';Background='#f8fafc';Border='#dbe3ef';Accent='#0f172a';Span=1},
    [pscustomobject]@{Label='Build';Value='16.0.10417.20198';Detail='known';Background='#eff6ff';Border='#bfdbfe';Accent='#1d4ed8';Span=2;FontSize=18}
)
$grid = New-SmartM365EmailKpiGridHtml -Cards $cards -RowSizes @(2)
$table = New-SmartM365EmailTableHtml -Headers @('Name','Status') -Rows @(@('A&B','Online'),@('C','Offline'))
$nowrapTable = New-SmartM365EmailTableHtml -Headers @('URL','Pool') -Rows @(,@('https://example.test/a','Pool Name')) -NoWrapColumns @(0,1)
if ($nowrapTable -notmatch '<td nowrap="nowrap"[^>]*>https://example\.test/a</td>' -or $nowrapTable -notmatch '<td nowrap="nowrap"[^>]*>Pool Name</td>') { throw 'Column-specific no-wrap rendering failed.' }
$html = New-SmartM365EmailBody -Title 'Layout fixture' -Category 'SmartM365 Fixture' -Severity Warning -Tenant 'fixture' -HostName 'fixture-host' -GeneratedAt '2026-10-09 00:00:00 +02:00' -StatusBadge 'PARTIAL' -Duration '00:00:31' -Message 'Fixture message.' -Sections @([pscustomobject]@{Title='Summary';Html=$grid},[pscustomobject]@{Title='Objects';Html=$table})
foreach ($expected in @('background:#0f172a','Tenant: fixture','Host: fixture-host','PARTIAL','Duration: 00:00:31','A&amp;B','16.0.10417.20198','border-collapse:collapse')) {
    if (-not $html.Contains($expected)) { throw "Email layout lacks $expected" }
}
if ($html.Contains('A&B')) { throw 'Email table did not HTML-escape a value.' }
$hasher = [Security.Cryptography.SHA256]::Create()
try { $hash = ([BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($html)))).Replace('-','') } finally { $hasher.Dispose() }
Write-Output "PASS: $($PSVersionTable.PSEdition) email layout, SHA256=$hash"

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAs5/38CC50m5TL
# F+E712d0Xw3s+6fZculpcFUBuil1NqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCW1X/B4Zdx6dzIxfiGke/l
# FVOLn4KubOr+/FYaGbyBujANBgkqhkiG9w0BAQEFAASCAYAnl6RpbNyPfNQOteej
# 7Qd+LkSnejMBrhcMpua0IFS+Oc1/8p0C6PMxZFKXSVKI/m1KkBMDrDf7FBahzUeQ
# Qa0acbUazlCkquIc9GezVSdwOqEUDmPxezXmgOo7UTlak580+TM1aqqPmaOiLVNX
# WPTiRdYLpxZC5CmrzCRyOBStVbuQo5HjXjuD1aMCT50oJyVguQPk7aEDbULZycQF
# IL/iUci8uXlGreb/G+KOPvWDwwlKE3vfXAq4DrZ1sZuPagxhwSdSMsQV0TbZCxXE
# WFqWLyLFJyyiwCO8QKHh44Yzz9ToXs+qKyqwBMTn+d2Tkn0+wTrKAvtCHfIOaWRX
# hugMmQXdkLq8ztd2taPEtP1ajuvMLUMo2te3MZTQsrtzMq7w20ctz9e3d05GKc9S
# v30yxBtcxEurZy1tJBhcrmbtk+PTqZ3zQLth/Aab4QPxn3739nJjwYAW3NS8R9dd
# zSQy8alKgsATzJGHATdS/hTfspOQVgE8LumIHJI7iF4RvFA=
# SIG # End signature block
