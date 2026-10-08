#Requires -Version 7.0
<#
.SYNOPSIS
Checks that non-started scheduled occurrences remain visible in history.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-History-' + [guid]::NewGuid().ToString('N'))
try {
    $runsPath = Join-Path $root 'WorkerA/JobRuns'
    $null = New-Item -ItemType Directory -Path $runsPath -Force
    $date = (Get-Date).Date
    $scheduled = $date.AddHours(1)
    $rows = @(
        [pscustomobject]@{ JobName = 'Succeeded'; ScheduledTime = $scheduled.ToString('o'); StartTime = $scheduled.ToString('o'); EndTime = $scheduled.AddMinutes(1).ToString('o'); DurationSec = 60; ExitCode = '0'; Status = 'Success'; RetryCount = 0; LogPath = 'success.log' }
        [pscustomobject]@{ JobName = 'Skipped'; ScheduledTime = $scheduled.AddMinutes(2).ToString('o'); StartTime = ''; EndTime = ''; DurationSec = ''; ExitCode = ''; Status = 'Skipped'; RetryCount = 0; LogPath = '' }
        [pscustomobject]@{ JobName = 'Blocked'; ScheduledTime = $scheduled.AddMinutes(3).ToString('o'); StartTime = ''; EndTime = ''; DurationSec = ''; ExitCode = ''; Status = 'BlockedDependencyFailed'; RetryCount = 0; LogPath = '' }
    )
    $csvPath = Join-Path $runsPath ('Orchestrator_JobRuns_' + $date.ToString('yyyyMMdd') + '.csv')
    $rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation
    Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Management.psm1') -Force
    $history = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $root -From $date -To $date.AddDays(1).AddTicks(-1))
    if ($history.Count -ne 3) { throw "Expected 3 history rows; got $($history.Count)." }
    foreach ($name in @('Skipped', 'Blocked')) {
        $row = @($history | Where-Object JobName -eq $name)[0]
        if ($null -eq $row -or $null -ne $row.StartTime -or $row.EventTime -lt $date) { throw "Non-started $name occurrence was lost or assigned a false start time." }
    }
    $filtered = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $root -From $date -To $date.AddDays(1).AddTicks(-1) -Status BlockedDependencyFailed)
    if ($filtered.Count -ne 1 -or $filtered[0].JobName -ne 'Blocked') { throw 'Blocked status filter is incorrect.' }
    'PASS: non-started occurrences remain visible with scheduled event time and no false start.'
}
finally {
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBYf6u/+UDinamV
# DmKvj7cRtOV6ZGg3ZYVhovbpBvBBuaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCkb63/M/Jub0ic7KQVfxqF
# eRDHEw1t+euzJvAIgh5JLDANBgkqhkiG9w0BAQEFAASCAYB0MSjPVhx+5VwqzEax
# ZZ/K2Rg43ls7uydX6RP+sKc5yqsAffnaWGCGEQcy9b0vFszai3gUGNUjLUTBGxx7
# yesSmzP9zV4T1c2zFYsCwWjeYak2XvpVXM7yCAvsz/w57GIydcHpwUmA/g2EX2Me
# EoO3vVwEKxNtH/oWIvDcxRaTmCVTZEzvV7fTfy5h4qzAHtvmhbF9WpICIko8TR3Q
# AuRPMFgclRIcuWQRLFnU/gZvX6I8bc4YYRk8Je8rVUQyacLofMUCCrHwpfD9PJr9
# rjNp0DACUeId2cdl00nmenXDKaPRpW5CwoBmlpu4IYZwKJf7hMwh2XIRZQVpMOD4
# sTyWgH37f+NbXoVbCro2ii1tsOpUtCIaP8SzaR8NF/lMXnmBmkjsO/p3iE+CYfAC
# a2PHWkDv49u+GvNBtGUUgOYYS6b3HwPQQth5e3T9KZimswFSXCfH8+ljOUPzaxf8
# z5gHtqwjg+M/K7EB53onAUNjjIyxfXsO7pqouAsxGvIDsS8=
# SIG # End signature block
