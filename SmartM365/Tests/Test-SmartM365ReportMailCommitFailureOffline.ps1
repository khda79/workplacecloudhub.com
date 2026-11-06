#Requires -Version 5.1
<#
.SYNOPSIS
    Reproduces a reader blocking state persistence after transport acceptance.
.VERSION
    1.0.0
#>
$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.Orchestrator.ReportMailQueue.psm1'
Import-Module $modulePath -Force
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-MailCommit-Test-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $testRoot
$rootFull = [IO.Path]::GetFullPath($testRoot)
$trust = { param($path) $path -eq $rootFull }
$now = [datetimeoffset]::UtcNow
$script:transportCalls = 0
$script:blockingReader = $null
try {
    $message = @{Subject='Synthetic commit failure test'}
    $item = Add-SmartM365ReportMailQueueItem -QueuePath $testRoot -AssertTrustedDirectory $trust -TenantKey demo -DeduplicationKey blocked-commit -NotBeforeUtc $now -Message $message
    $recordPath = Join-Path $testRoot ($item.Id + '.json')
    $script:blockingReader = [IO.File]::Open($recordPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $batch = Invoke-SmartM365ReportMailQueueBatch -QueuePath $testRoot -AssertTrustedDirectory $trust -TenantKey demo -Now $now -SendAction { $script:transportCalls++; $true }
    if ($batch.Busy -ne 1 -or $script:transportCalls -ne 0) { throw 'A failed pre-send checkpoint still called the transport.' }
    $script:blockingReader.Dispose(); $script:blockingReader = $null
    $record = Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json
    if ($record.Status -ne 'Pending' -or $record.Attempts -ne 0) { throw 'Pre-send failure changed the durable record.' }

    $send = {
        param($payload)
        $checkpoint = Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json
        if ($checkpoint.Status -ne 'Sending' -or $checkpoint.Attempts -ne 1) { throw 'Missing durable checkpoint before transport.' }
        $script:transportCalls++
        # Model another reader without FILE_SHARE_DELETE during the final commit.
        $script:blockingReader = [IO.File]::Open($recordPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        return $true
    }
    $batch = Invoke-SmartM365ReportMailQueueBatch -QueuePath $testRoot -AssertTrustedDirectory $trust -TenantKey demo -Now $now -SendAction $send -WarningVariable commitWarnings
    if ($batch.Uncertain -ne 1 -or $batch.Busy -ne 0 -or $script:transportCalls -ne 1 -or @($commitWarnings).Count -ne 1) { throw 'Accepted but uncommitted delivery was not reported as uncertain.' }
    $script:blockingReader.Dispose(); $script:blockingReader = $null
    $record = Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json
    if ($record.Status -ne 'Sending' -or $record.Attempts -ne 1 -or -not $record.PayloadBase64 -or -not $record.AttemptHost) { throw 'Recovery evidence was lost.' }

    # Reload the module to prove replay protection is persisted, not process-local.
    Remove-Module SmartM365.Orchestrator.ReportMailQueue -Force
    Import-Module $modulePath -Force
    $batch = Invoke-SmartM365ReportMailQueueBatch -QueuePath $testRoot -AssertTrustedDirectory $trust -TenantKey demo -Now $now.AddDays(1) -SendAction { $script:transportCalls++; $true }
    if ($batch.Uncertain -ne 1 -or $batch.Sent -ne 0 -or $script:transportCalls -ne 1) { throw 'An uncertain delivery was replayed after restart.' }

    # A reader that allows deletion does not block atomic replacement.
    $good = Add-SmartM365ReportMailQueueItem -QueuePath $testRoot -AssertTrustedDirectory $trust -TenantKey demo -DeduplicationKey shared-read -NotBeforeUtc $now -Message $message
    $goodPath = Join-Path $testRoot ($good.Id + '.json')
    if ([string]::CompareOrdinal($item.Id, $good.Id) -ge 0) { throw 'Fixture must place the uncertain record before the ready record.' }
    $batch = Invoke-SmartM365ReportMailQueueBatch -QueuePath $testRoot -AssertTrustedDirectory $trust -TenantKey demo -Now $now -MaximumItems 1 -SendAction {
        $script:transportCalls++
        $script:blockingReader = [IO.File]::Open($goodPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $true
    }
    $script:blockingReader.Dispose(); $script:blockingReader = $null
    if ($batch.Sent -ne 1 -or $batch.Uncertain -ne 1 -or $script:transportCalls -ne 2) { throw 'Uncertain record or compatible shared reader blocked delivery.' }
    $record = Get-Content -LiteralPath $goodPath -Raw | ConvertFrom-Json
    if ($record.Status -ne 'Sent' -or $record.PayloadBase64 -ne '' -or -not $record.SentAtUtc) { throw 'Successful delivery was not finalized.' }
    Write-Host 'PASS: pre-send failure, post-send commit failure, persisted replay protection and compatible shared reader; no mail sent.'
} finally {
    if ($script:blockingReader) { $script:blockingReader.Dispose() }
    $resolved = [IO.Path]::GetFullPath($testRoot)
    if ($resolved -ne $rootFull -or (Split-Path $resolved -Leaf) -notmatch '^SmartM365-MailCommit-Test-[a-f0-9]{32}$' -or (Split-Path $resolved -Parent).TrimEnd('\') -ne ([IO.Path]::GetTempPath()).TrimEnd('\')) { throw 'Unexpected test cleanup path.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBAGVSxpIQ1bUHu
# 80Y0vAZ3708wvOdrK13JhKL6OrCaH6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCltrv9dIpxxTKhGb42Is5r
# 9lvc5SB8Zv1BAI/02WCSdzANBgkqhkiG9w0BAQEFAASCAYBnMMOdJ/3DAcBQmBtf
# UrQyN2/6oOcV3eP5B1EQ2cUmPDSJHCNBiPK2h85h3rK1RMm2vJTKWIrjxYskVg68
# 4pIdp8CXfYmq0T8D0n/JlwhHCzUkQNlw+iFcJrNt9+rVZw2LKDGcLNzSJLA3GDuz
# rb3jBDaN5d6Eemphs9vyOFhRM1PMGCYSdYmW9pOne4eWVwXD10xLQBUYSnmQykbo
# s6p90NwhB/tfwb178bmWXTF3xsJnt5G3sRnBLF4AXRFwBFJPFFB8/k2lnQeWpUXq
# /ohFN7ogZME+P1CR8NxcSC4yeYFx3pNAOoesa8r8FxSyh0Q8jvX9hqIn+DLcyg8n
# irrE+VFqBKpVHKFhTqcpUvtvZAAW2lLHmTBHE9IHe6RLnurwu6t8rCQHbIINVrU6
# SOZd74TseIbU0T6v0gA0+2/ydNSNLL/+9NGjOaiiWIngdNzyFvvgNoGcW3HUrM13
# qYFoRvnPynOPcaaG5dsqCC01LTrvZArq7jg0wumYEBfVVl4=
# SIG # End signature block
