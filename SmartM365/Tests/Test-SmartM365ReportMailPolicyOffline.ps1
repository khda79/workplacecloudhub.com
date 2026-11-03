<# Offline policy, queue and mocked-transport integration test. #>
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.Orchestrator.ReportMailPolicy.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.Orchestrator.ReportMailQueue.psm1') -Force
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ReportMailPolicy-Test-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $testRoot)
$testRootFullPath = (Get-Item -LiteralPath $testRoot).FullName
$trust = { param($path) $path -eq $testRootFullPath }
$script:sent = 0
$send = { param($message) if ($message.Subject -ne 'Synthetic') { throw 'Unexpected payload' }; $script:sent++; return $true }
$now = [datetimeoffset]::Parse('2026-11-03T06:29:00Z')
$params = @{ Now=$now; Message=@{ Subject='Synthetic' }; SendAction=$send; QueuePath=$testRoot; AssertTrustedDirectory=$trust; TenantKey='demo'; DeduplicationKey='daily' }
try {
    $report = Submit-SmartM365ReportMail @params -Enabled $true -DeliveryClass Report
    if ($report.Status -ne 'Queued' -or $script:sent -ne 0) { throw 'Early report was not queued.' }
    foreach ($class in @('Operational','Error','Maintenance')) {
        $result = Submit-SmartM365ReportMail @params -Enabled $true -DeliveryClass $class
        if ($result.Status -ne 'Sent') { throw "$class was not immediate." }
    }
    $disabled = Submit-SmartM365ReportMail @params -Enabled $false -DeliveryClass Report
    if ($disabled.Status -ne 'Sent' -or $script:sent -ne 4) { throw 'Disabled policy did not preserve immediate sending.' }
    $before = Invoke-SmartM365ReportMailQueueBatch -QueuePath $testRoot -AssertTrustedDirectory $trust -TenantKey demo -Now $now -SendAction $send
    if ($before.Pending -ne 1 -or $script:sent -ne 4) { throw 'Batch delivered early.' }
    $after = Invoke-SmartM365ReportMailQueueBatch -QueuePath $testRoot -AssertTrustedDirectory $trust -TenantKey demo -Now $now.AddMinutes(1) -SendAction $send
    if ($after.Sent -ne 1 -or $script:sent -ne 5) { throw 'Batch did not deliver at the cutoff.' }
    Write-Host 'PASS: simulated policy admission, cutoff and mocked delivery; no mail sent.'
} finally {
    $resolved = (Get-Item -LiteralPath $testRoot -Force).FullName
    if ($resolved -ne $testRootFullPath -or -not (Split-Path -Leaf $resolved).StartsWith('SmartM365-ReportMailPolicy-Test-')) { throw 'Refusing to remove unexpected test path.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDhp6C+tRqvyJ2a
# yS0UrDvqGKeLl80IrdHPnvhDVCOfxqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD/hd1k1Q4e7ncq+tlvYjTO
# 8V2ba5n7E090gw1fzv3y0zANBgkqhkiG9w0BAQEFAASCAYA/CA7La1wjzlMDv6XS
# 1gQmsCHcchJcFHv/aUS6RYbEzO6P1ZBmvwNYBl7JiKHSCKFw0Vh4NoWjg97+0qst
# 4tAB5MID7GI5bX44r5E0he9jYNfTtJ1BizYZ8nZGezPMEr/cYOVmamMFo1zKmreD
# gZ9evsNHhpNeeYk448nbzZaAzfzYgFfH/Rao9+ff6OmJXgo0ZW7k+yAo/2+bUKsi
# AZw623F43lZmd0R5Ske6qFn9b/Jiiam0cZxi4e4jUpPxFxZak6iONmnbBzPfGmSC
# CerZAbIRNNpECYI/5/hdlRSU5mtwc41p0O3RV3roctKgXbs72Y30R7dCdLP20vLZ
# fPP1yaHb/NPpb9Om25FuWUAxTkd/nzGrOBGp79POmt/i4v9ZLNdEHGEX8RxO+DJ9
# N0SZUFanxz2IWGJRWvdoDj8MkBPqj+nPJBYYIp7ktpbNzzKHMgDSdwphlXrZvw1i
# 7xIMs3/dS/C9Z3TN5O+fHTiyZQt5WLNvPxNOJW3JkYNdGMY=
# SIG # End signature block
