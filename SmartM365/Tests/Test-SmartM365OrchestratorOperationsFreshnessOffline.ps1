#Requires -Version 7.0
<#
.SYNOPSIS
Checks that old heartbeats cannot make running or pending jobs look live.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-OperationsFreshness-' + [guid]::NewGuid().ToString('N'))
try {
    $serverPath = Join-Path $root 'WorkerA'
    $null = New-Item -ItemType Directory -Path $serverPath -Force
    $heartbeatPath = Join-Path $serverPath 'Orchestrator-Heartbeat.json.txt'
    $heartbeat = [ordered]@{
        Timestamp = [datetime]::UtcNow.AddMinutes(-30).ToString('o')
        Lifecycle = 'Starting'; JobStopProtocol = 1
        RunningJobs = @([pscustomobject]@{Name='RunningJob';Pid=123;StartTime=[datetime]::UtcNow.AddHours(-1).ToString('o')})
        PendingJobs = @([pscustomobject]@{Name='PendingJob';Reason='WaitingDependencies';FirstSeen=[datetime]::UtcNow.AddMinutes(-20).ToString('o')})
    }
    [IO.File]::WriteAllText($heartbeatPath, ($heartbeat | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Insights.psm1') -Force
    $cluster = [pscustomobject]@{ExpectedOrchestratorServers=@('WorkerA');PeerHeartbeatStaleMinutes=5}
    $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath $root -ClusterDocument $cluster
    if ($operations.Servers[0].State -ne 'Stale' -or $operations.Running[0].ServerState -ne 'Stale' -or $operations.Running[0].StopSupported -or $operations.Pending[0].ServerState -ne 'Stale') {
        throw 'A stale heartbeat was displayed as a live running or pending job.'
    }
    $heartbeat.Remove('Timestamp')
    [IO.File]::WriteAllText($heartbeatPath, ($heartbeat | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath $root -ClusterDocument $cluster
    if ($operations.Servers[0].State -ne 'No timestamp' -or $operations.Running[0].StopSupported) { throw 'An undated heartbeat was treated as live.' }
    'PASS: stale and undated heartbeats are explicit in job rows and cannot enable stop.'
}
finally {
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUEeb6Qz7SH1Vx
# Io4+m+mplOu6th5ITh05riDxSZkr+6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCChStePVKk4JjGv3wE6rlHV
# 22GWsziA5M3BRZpcMggKSjANBgkqhkiG9w0BAQEFAASCAYCnwLUWx5OLrN60roWU
# WasKLF3cNsKuVuLeWs/XMUE2JgXhkh/KQFuOVujDdW/vS5rhlmJsrimpTd0DHHQQ
# 2GVDC7u+z2k540BMrSwXTpkuRqgMQpnPKei+SOVRwghVWxjadpiYiXcjqQww7jDr
# lHySrABiZvAUt/oXC9BpVnq8u15e21ZCOR0TUQPT4FcyfZgpxrE6ZrBnhD7SwSWT
# 7x8Y/oI1/lo/ZDRFtoeYglbTFy8+/OhJOZKe7u5/f9kbspMbmKmKr4cKD/B5/yYe
# +TMBJZSFGDIOOiyydNUEZwM04qsbPn3oYjD89QGnhRMbAy/jNsbIhfQ9a+b9Gjbg
# JjnxKmLjp2bmtuMxbGm3QD5nTuV2n91BuBd/PZ73NnQ0XxpEx9RIig3gVTsEFz4e
# fnrzoWOcXZoRNVxKtufx89eWgVbYvU32+G17YWxEtrvZ6rb3hfpqj0arVQ6KEARv
# GztWkZDfm6SJy2EV+4Wo3pbVgzj7+LEmJDA5nx/DcLp0Bg8=
# SIG # End signature block
