<#
.SYNOPSIS
Offline tests of the peer monitoring alerts (orchestrator 1.5.78): unconfirmed observations logged as
INFO, every issue confirmed on its own, one alert and one recovery mail per issue for the cluster.
.VERSION
1.0.0
.NOTES
Two observer servers run the real orchestrator functions, extracted from the script, in separate
module scopes with a shared temporary folder. The peer snapshot and the mail are stubbed: no mail is
sent and no real peer is read.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'), [ref]$null, [ref]$null)
function Get-FunctionText([string]$Name) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing orchestrator function: $Name" }
    $definition.Extent.Text
}
$jsonModule = Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1'
$functions = @('Invoke-OrchestratorPeerMonitoring', 'Get-OrchestratorPeerMonitoringState', 'Format-OrchestratorPeerIssue', 'Reset-OrchestratorPeerMonitoringIncident',
    'Get-OrchestratorPeerAlertPaths', 'Read-OrchestratorPeerAlertState', 'Save-OrchestratorPeerAlertState', 'ConvertFrom-OrchestratorPeerMonitorTime',
    'Enter-OrchestratorSharePointMirrorLock', 'Exit-OrchestratorSharePointMirrorLock')
$stubs = @'
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Logs.Add("[$Level] $Message") }
function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now, $Level) $script:Logs.Add("[WARN] $Message") }
function Write-FileAtomically { param($Path, $Content) [IO.File]::WriteAllText($Path, $Content, [Text.UTF8Encoding]::new($false)) }
function Get-OrchestratorPeerHealthSnapshot { param($Now) [pscustomobject]@{ CheckedAt = $Now; Issues = @($global:PeerTestIssues); Peers = @() } }
function Send-OrchestratorPeerMonitoringEmail { param($Snapshot, [switch]$Recovery, $PreviousSummary, $Category = 'Jobs')
    $global:PeerTestMails.Add([pscustomobject]@{ Observer = $env:COMPUTERNAME; Recovery = [bool]$Recovery; Category = $Category; Keys = @($Snapshot.Issues | ForEach-Object { [string]$_.Key }) })
    $true }
'@
$functionText = (@($functions | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs
$originalComputer = $env:COMPUTERNAME
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-PeerAlert-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root -Force | Out-Null
function New-Observer([string]$Name) {
    $module = New-Module -Name ('Observer-' + $Name) -ScriptBlock ([scriptblock]::Create($functionText))
    & $module {
        param($Root, $JsonModule)
        Import-Module $JsonModule -Force
        $script:Logs = [Collections.Generic.List[string]]::new()
        $script:State = @{ Jobs = @{} }
        $script:LastPeerMonitoringCheckTime = [datetime]::MinValue
        $script:Settings = [pscustomobject]@{
            SharedDataFolderPath = $Root; PeerMonitoringEnabled = $true; ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'SRV-C')
            PeerMonitoringCheckIntervalSeconds = 60; PeerMonitoringConfirmationChecks = 2; PeerAlertReminderMinutes = 240; PeerAlertMailRetryMinutes = 15
            PeerRecoveryEmailEnabled = $true; AtomicWriteRetrySeconds = 5
        }
    } $root $jsonModule
    $module
}
function Invoke-Check([string]$Server, $Module, [datetime]$Now) {
    $env:COMPUTERNAME = $Server
    try { & $Module { param($n) $script:Logs.Clear(); Invoke-OrchestratorPeerMonitoring -Now $n; @($script:Logs) } $Now }
    finally { $env:COMPUTERNAME = $originalComputer }
}
function New-Issue([string]$Key, [string]$Server, [string]$Type, [string]$Job = '') { [pscustomobject]@{ Key = $Key; Server = $Server; Job = $Job; Type = $Type } }

try {
    $global:PeerTestMails = [Collections.Generic.List[object]]::new()
    $a = New-Observer 'SRV-A'; $b = New-Observer 'SRV-B'
    $t = [datetime]'2026-10-30T08:00:00'
    $stale = New-Issue 'HeartbeatStale|SRV-C' 'SRV-C' 'HeartbeatStale'
    $job = New-Issue 'JobNotStarted|SRV-C|Job-1' 'SRV-C' 'JobNotStarted' 'Job-1'

    # 1. First observation: candidate, INFO only, no mail.
    $global:PeerTestIssues = @($stale)
    $logsA = Invoke-Check 'SRV-A' $a $t; $logsB = Invoke-Check 'SRV-B' $b $t
    Check ($global:PeerTestMails.Count -eq 0) 'An unconfirmed issue sends no mail.'
    Check (@(@($logsA) + @($logsB) | Where-Object { $_ -like '`[INFO`] Peer monitoring candidate issue: SRV-C: HeartbeatStale (1/2)' }).Count -eq 2 -and @(@($logsA) + @($logsB) | Where-Object { $_ -like '`[WARN`]*' }).Count -eq 0) 'Candidate issues are INFO, never WARN.'

    # 2. Confirmed: the first observer sends, the second only acknowledges.
    $t = $t.AddMinutes(1)
    $logsA = Invoke-Check 'SRV-A' $a $t; $logsB = Invoke-Check 'SRV-B' $b $t
    Check ($global:PeerTestMails.Count -eq 1 -and $global:PeerTestMails[0].Observer -eq 'SRV-A' -and $global:PeerTestMails[0].Category -eq 'Server') 'A confirmed issue sends one alert for the cluster.'
    Check (@($logsB | Where-Object { $_ -like '*alert already sent by another server: SRV-C: HeartbeatStale (by SRV-A)*' }).Count -eq 1) 'The second observer logs that the alert was already sent.'
    Check ((& $b { $script:State.PeerMonitoring.ActiveIssueKey }) -eq 'HeartbeatStale|SRV-C') 'The second observer still tracks the incident for its recovery.'
    $shared = Get-Content -LiteralPath (Join-Path $root 'Orchestrator-PeerAlert.json') -Raw | ConvertFrom-Json
    Check ($shared.Issues.'HeartbeatStale|SRV-C'.Observer -eq 'SRV-A') 'The shared state records who sent the alert.'

    # 3. A new issue does not restart the count of the confirmed one.
    $global:PeerTestIssues = @($stale, $job)
    $t = $t.AddMinutes(1)
    $logsA = Invoke-Check 'SRV-A' $a $t; $logsB = Invoke-Check 'SRV-B' $b $t
    Check ($global:PeerTestMails.Count -eq 1 -and @($logsA | Where-Object { $_ -like '*candidate issue: SRV-C/Job-1: JobNotStarted (1/2)' }).Count -eq 1 -and @($logsA | Where-Object { $_ -like '*HeartbeatStale (1/2)*' }).Count -eq 0) 'The confirmed issue keeps its count while a new one is a candidate.'
    $t = $t.AddMinutes(1)
    $null = Invoke-Check 'SRV-B' $b $t; $null = Invoke-Check 'SRV-A' $a $t
    Check ($global:PeerTestMails.Count -eq 2 -and $global:PeerTestMails[1].Observer -eq 'SRV-B' -and $global:PeerTestMails[1].Category -eq 'Jobs' -and (@($global:PeerTestMails[1].Keys) -join ',') -eq 'JobNotStarted|SRV-C|Job-1') 'Only the newly confirmed issue is alerted, once.'

    # 4. Recovery: one mail for the cluster, then the other observer closes silently.
    $global:PeerTestIssues = @()
    foreach ($step in 1..2) {
        $t = $t.AddMinutes(1)
        $logsA = Invoke-Check 'SRV-A' $a $t; $logsB = Invoke-Check 'SRV-B' $b $t
    }
    $recoveries = @($global:PeerTestMails | Where-Object Recovery)
    Check ($recoveries.Count -eq 1 -and $recoveries[0].Observer -eq 'SRV-A') 'One recovery mail for the cluster.'
    Check (@($logsB | Where-Object { $_ -like '*recovery email was already sent by another server*' }).Count -eq 1 -and [string]::IsNullOrEmpty((& $b { $script:State.PeerMonitoring.ActiveIssueKey }))) 'The other observer closes its incident without a mail.'
    $shared = Get-Content -LiteralPath (Join-Path $root 'Orchestrator-PeerAlert.json') -Raw | ConvertFrom-Json
    Check (@($shared.Issues.PSObject.Properties).Count -eq 0) 'The recovery clears the shared alerts.'

    # 5. A flapping observation (seen, cleared, seen) never confirms.
    $mailsBefore = $global:PeerTestMails.Count
    foreach ($issuesNow in @(@($stale), @(), @($stale), @())) {
        $global:PeerTestIssues = $issuesNow
        $t = $t.AddMinutes(1)
        $null = Invoke-Check 'SRV-A' $a $t
    }
    Check ($global:PeerTestMails.Count -eq $mailsBefore) 'A flapping observation that never lasts two checks sends nothing.'

    # 6. A stale shared entry (older than two reminder periods) does not block a new alert.
    $old = [ordered]@{ SchemaVersion = 1; Issues = [ordered]@{ 'HeartbeatStale|SRV-C' = [ordered]@{ SentAtUtc = $t.ToUniversalTime().AddHours(-9).ToString('o'); Observer = 'SRV-B' } } }
    [IO.File]::WriteAllText((Join-Path $root 'Orchestrator-PeerAlert.json'), ($old | ConvertTo-Json -Depth 5))
    $global:PeerTestIssues = @($stale)
    foreach ($step in 1..2) { $t = $t.AddMinutes(1); $null = Invoke-Check 'SRV-A' $a $t }
    Check ($global:PeerTestMails.Count -eq $mailsBefore + 1 -and $global:PeerTestMails[-1].Observer -eq 'SRV-A') 'An expired shared entry lets a new alert through.'
    Check (-not (Test-Path -LiteralPath (Join-Path $root 'Orchestrator-PeerAlert.lock'))) 'The shared lock is released.'
}
finally {
    $env:COMPUTERNAME = $originalComputer
    Remove-Variable -Name PeerTestMails, PeerTestIssues -Scope Global -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: peer monitoring alerts, $($script:passed) offline checks."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAScaMlNFPTehQ2
# 8po9OxS/Zi9Zt7zyDDsUMjrz2kg0iqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIAuf8PSOGNUfEtI8h5hE34z3QNppLWvMHPn1vcJ+Y9r3MA0GCSqG
# SIb3DQEBAQUABIIBgBi06ZVNE8T99DMQwAwi/BSwYPdBhyQ7G5GZekQJO6zIYCGL
# SQk0LzRQZjYAPGCJbR5d1RIoLe5o0NqWzitv6Be+yChp92dbfYp8I5NR3RwabYU0
# rJa2lBYtECotVvqJYxDShQfcG6JYCjPruVMRoMXYAgHtVlNnLOX/PHfiPX62s0o7
# sM868Dp0HwHw+ILvhe02oAAZYWst1z1hq7trAmJSbc2UsvDNeU+ibyOyn/OyQRUV
# ef+D19eKC/gPR6SwWdfzKKVdxtXV0P2iyVQDi8b6fBfV0YnWu2xLWXnLa8eBWnPf
# dtr7p5OVfjIQ/rE1sNUePL3YYBzlVxyX+NYaq/Fn+/+TvPMok8DBi3AMQAkoNLm2
# A2z3TaDLC1oUfS+sYc50MgdOZ74l2iyMcFTX0cX+/2Z3NJLrNUX32GxsmJ+3Z4Wm
# l+tldhOZbQ04r6FnjPjcUVoSqF/CpINCEgIm5IiIcqB2G4rIMeNP0BYlQgLOFRWW
# Hsiddz9I/o4jyfYD9aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAwODAz
# MzdaMC8GCSqGSIb3DQEJBDEiBCBVcZkmvbr+SWyx+57ySTez1PFw2AJlQhs7MCcf
# /A8E8zANBgkqhkiG9w0BAQEFAASCAgBXFH6JniFYT4EEk09ugTDgipzwe/2Q8IMc
# 86OPOUDDAlwNkE7ezhI5jfK28JtNeVhy0Rc8oX+LQdbaKFTdb1I6DHP/OexpTALX
# LdP+ASdOQdsyqvUwqhoDXsdMXbNkOhWXJZouNE+DGPY7Bu5uU/+6wOk6mVnnFrnA
# txbvKskoi/UPWqyknZfbUPdWAwt1jjZueAIJmETbfurGHTCufHOZl2/xkpkpTU5O
# rTTN26TEI5g8nkot/Wr5T+Yb9ObKrBOFXirzZeu7mxb1s+1jUkDppI9shw1pfA/w
# bogkMk0BeU7oqbco3OXZwTdGJLLGaykyExuAdqakn1M2Bk9y4o64RGJoFy9cErE2
# W6jBozETmTwa7SQRwx3913JupJFJons81VlYm4vthDfv/YZupGPAko6dRlD18cxZ
# HzNcDMMo0ZtD0FSuSX6RTgSBQMoer80Tb7LVshimvkm7GXXqaWwrme8i/6/q2rL+
# RDC+7GqC4ny14UAjC9zxHXGMwPGMC8BALF8lLlfGSUnsCjz3V3ljOSt5DZhgXKxf
# uNCnYbNZpiVSXkvC3JInngM9kR4xMIahKdN0ryIk0FMmruGFil6IN8dI+EY39Ei0
# dUgwt9dzKojnjcu3U8dTdNGq+ONYy78clmRUly4yOLlRUSHTNaQi7TAEQbLcm7JI
# Vykw/k/M/Q==
# SIG # End signature block
