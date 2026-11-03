<#
.SYNOPSIS
    Exercises the isolated report-mail queue with a mocked transport only.
#>
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.Orchestrator.ReportMailQueue.psm1') -Force
$queuePath = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ReportMailQueue-Test-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $queuePath)
$testQueueFullPath = (Get-Item -LiteralPath $queuePath).FullName
$trust = { param($path) $path -eq $testQueueFullPath }
$script:sent = 0
$send = { param($message) if ($message.Subject -ne 'Synthetic' -or [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($message.AttachmentBase64)) -ne 'synthetic attachment') { throw 'Unexpected message' }; $script:sent++; return $true }
$now = [datetimeoffset]::Parse('2026-11-03T06:29:00Z')
$due = $now.AddMinutes(1)
try {
    $message = @{ Subject = 'Synthetic'; AttachmentBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('synthetic attachment')) }
    $item = Add-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -TenantKey 'demo' -DeduplicationKey 'daily-report' -NotBeforeUtc $due -Message $message
    if ($item.Status -ne 'Pending' -or $item.Existing) { throw 'Initial enqueue failed.' }
    $duplicate = Add-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -TenantKey 'demo' -DeduplicationKey 'daily-report' -NotBeforeUtc $due -Message $message
    if (-not $duplicate.Existing -or $duplicate.Id -ne $item.Id) { throw 'Idempotent enqueue failed.' }
    $failed = $false
    try { $null = Add-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -TenantKey 'demo' -DeduplicationKey 'daily-report' -NotBeforeUtc $due -Message @{ Subject='Different' } } catch { $failed = $true }
    if (-not $failed) { throw 'A different message was silently discarded by deduplication.' }
    $status = Invoke-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -Id $item.Id -TenantKey 'demo' -Now $now -SendAction $send
    if ($status -ne 'Pending' -or $script:sent -ne 0) { throw 'Sent before due time.' }
    $status = Invoke-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -Id $item.Id -TenantKey 'demo' -Now $due -SendAction { throw 'Synthetic transport failure' }
    if ($status -ne 'Retry') { throw 'Failed transport was not retained.' }
    $status = Invoke-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -Id $item.Id -TenantKey 'demo' -Now $due.AddMinutes(4) -SendAction $send
    if ($status -ne 'Pending' -or $script:sent -ne 0) { throw 'Retry occurred too early.' }
    $status = Invoke-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -Id $item.Id -TenantKey 'demo' -Now $due.AddMinutes(5) -SendAction $send
    if ($status -ne 'Sent' -or $script:sent -ne 1) { throw 'Retry did not send.' }
    $sentRecord = Get-Content -LiteralPath (Join-Path $queuePath ($item.Id + '.json')) -Raw | ConvertFrom-Json
    if ($sentRecord.PayloadBase64 -ne '') { throw 'Delivered message content was retained.' }
    $status = Invoke-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -Id $item.Id -TenantKey 'demo' -Now $due.AddMinutes(6) -SendAction $send
    if ($status -ne 'Sent' -or $script:sent -ne 1) { throw 'Sent message was replayed.' }
    $failed = $false
    try { $null = Invoke-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -Id $item.Id -TenantKey 'other' -Now $due -SendAction $send } catch { $failed = $true }
    if (-not $failed) { throw 'Cross-tenant delivery was accepted.' }
    $failed = $false
    try { $null = Add-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory { $false } -TenantKey 'demo' -DeduplicationKey 'another' -NotBeforeUtc $due -Message $message } catch { $failed = $true }
    if (-not $failed) { throw 'Untrusted directory was accepted.' }
    $other = Add-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -TenantKey 'demo' -DeduplicationKey 'no-transport-confirmation' -NotBeforeUtc $due -Message $message
    $status = Invoke-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -Id $other.Id -TenantKey 'demo' -Now $due -SendAction { $false }
    if ($status -ne 'Retry') { throw 'Unconfirmed transport was treated as sent.' }
    $lockPath = Join-Path $queuePath ($other.Id + '.lock')
    $lock = [IO.File]::Open($lockPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try {
        $failed = $false
        try { $null = Invoke-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -Id $other.Id -TenantKey 'demo' -Now $due -SendAction $send } catch [IO.IOException] { $failed = $true }
        if (-not $failed) { throw 'A concurrent worker acquired an active lock.' }
    } finally { $lock.Dispose() }
    $corrupt = Add-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -TenantKey 'demo' -DeduplicationKey 'integrity-test' -NotBeforeUtc $due -Message $message
    $corruptPath = Join-Path $queuePath ($corrupt.Id + '.json')
    $record = Get-Content -LiteralPath $corruptPath -Raw | ConvertFrom-Json
    $record.PayloadBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"Subject":"changed"}'))
    [IO.File]::WriteAllText($corruptPath, ($record | ConvertTo-Json -Depth 32 -Compress), [Text.UTF8Encoding]::new($false))
    $failed = $false
    try { $null = Invoke-SmartM365ReportMailQueueItem -QueuePath $queuePath -AssertTrustedDirectory $trust -Id $corrupt.Id -TenantKey 'demo' -Now $due -SendAction $send } catch { $failed = $true }
    if (-not $failed -or $script:sent -ne 1) { throw 'A modified queue payload was accepted.' }
    $batch = Invoke-SmartM365ReportMailQueueBatch -QueuePath $queuePath -AssertTrustedDirectory $trust -TenantKey 'demo' -Now $due.AddMinutes(6) -SendAction $send
    if ($batch.Sent -ne 1 -or $batch.Invalid -ne 1 -or $script:sent -ne 2) { throw 'Batch did not process retry or report invalid record.' }
    $batch = Invoke-SmartM365ReportMailQueueBatch -QueuePath $queuePath -AssertTrustedDirectory $trust -TenantKey 'demo' -Now $due.AddMinutes(7) -SendAction $send
    if ($batch.Sent -ne 0 -or $script:sent -ne 2) { throw 'Batch replayed a delivered message.' }
    $batchPath = Join-Path $queuePath 'batch'
    [void](New-Item -ItemType Directory -Path $batchPath)
    $batchFullPath = (Get-Item -LiteralPath $batchPath).FullName
    $batchTrust = { param($path) $path -eq $batchFullPath }
    $first = Add-SmartM365ReportMailQueueItem -QueuePath $batchPath -AssertTrustedDirectory $batchTrust -TenantKey 'demo' -DeduplicationKey 'candidate-1' -NotBeforeUtc $due.AddHours(1) -Message $message
    $second = Add-SmartM365ReportMailQueueItem -QueuePath $batchPath -AssertTrustedDirectory $batchTrust -TenantKey 'demo' -DeduplicationKey 'candidate-2' -NotBeforeUtc $due.AddHours(1) -Message $message
    $laterId = @($first.Id, $second.Id | Sort-Object)[-1]
    $duePath = Join-Path $batchPath ($laterId + '.json')
    $record = Get-Content -LiteralPath $duePath -Raw | ConvertFrom-Json
    $record.NotBeforeUnixMs = $due.ToUnixTimeMilliseconds()
    [IO.File]::WriteAllText($duePath, ($record | ConvertTo-Json -Depth 32 -Compress), [Text.UTF8Encoding]::new($false))
    $batch = Invoke-SmartM365ReportMailQueueBatch -QueuePath $batchPath -AssertTrustedDirectory $batchTrust -TenantKey 'demo' -Now $due -SendAction $send -MaximumItems 1
    if ($batch.Sent -ne 1 -or $batch.Pending -ne 1 -or $script:sent -ne 3) { throw 'A future item prevented a due message from being processed.' }
    Write-Host 'PASS: queue admission, scheduling, retry, deduplication and isolation; no mail sent.'
} finally {
    $resolved = (Get-Item -LiteralPath $queuePath -Force).FullName
    if ($resolved -ne $testQueueFullPath -or -not (Split-Path -Leaf $resolved).StartsWith('SmartM365-ReportMailQueue-Test-')) { throw 'Refusing to remove unexpected test path.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCH8ARca6dMA2nN
# RMmaCV3DFvD1wRjq7V6wqeeeELEwWKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBxlJRX2FI0yW4GJD23tGC9
# ga+qMJY3N2Q2O9F/9qBp2DANBgkqhkiG9w0BAQEFAASCAYASHCXqUkPcIhdAYlnF
# JRN8UXy4R63lAITiGp1uw367bLhnaQaRk51+03NjdKTAU2KAQSCtYh2En/i4QYuu
# ApjPmxBKGJCe/une3+5zYpg7cmsEmt7YDnxVMF5rWJjNd/1m63OeiEV+/Ww6lBxX
# 1EJA5cXmjjeKNiORUhg8wrM5moqr90Q7UCHtXmVwtmXPoT77Oym6YwSgpSaDA3MY
# cA2nZobn/ttrGK6ChINpv4g1RvWPnrHG9ov6IAyipVTBp5b+rG6kg4UxtUEbozP0
# bv0NfXc44F0ZSbX2xumNoQanLVVFPYuAxbswX86DCNgYSSQuLybl2CE96YK9C0Y5
# bTXBueQcz+EKkLF9VQ8Z89OfxIk/FZWbbuepfwEWQymOqLYtfEBr6RL+zphboVvg
# Y0qqqnOoEJsrmQMk+B2i2LrTfmY1qKYporkyrRBdIRi/SnFGNItztNs9HST6h0Ws
# ERBImV7NUH0XDhM3K1Gki2/JctM4kqBHjDFFu04jwylDMBw=
# SIG # End signature block
