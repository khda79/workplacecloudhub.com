<#
.SYNOPSIS
Offline tests of the server system information: each orchestrator reads OS, CPU, memory and drive
C: once at startup and publishes them in its heartbeat; the GUI Servers and versions view reads
them from the heartbeat and never queries the servers.
.VERSION
1.0.2
.NOTES
Synthetic heartbeats in a temporary folder. The collector runs read-only CIM queries on this
computer. No tenant initialization, collectors, scheduling or publication.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
function Get-FunctionText([string]$Name) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing orchestrator function: $Name" }
    $definition.Extent.Text
}

# Startup collection and heartbeat publication.
$source = [IO.File]::ReadAllText($orchestratorPath)
Check ($source -match '\$script:SystemInfo = Get-OrchestratorSystemInfo') 'The orchestrator no longer reads its system information at startup.'
Check ((Get-FunctionText 'Write-OrchestratorHeartbeat') -match 'SystemInfo = \$script:SystemInfo') 'The heartbeat no longer publishes the system information.'
Check (([regex]::Matches($source, 'Get-OrchestratorSystemInfo')).Count -eq 2) 'The system information must be read once, at startup only.'

# Real collector on this computer: values or a named error, never an exception.
. ([scriptblock]::Create((Get-FunctionText 'Get-OrchestratorSystemInfo')))
$info = Get-OrchestratorSystemInfo
Check ($null -ne $info -and $info.SystemDrive -eq 'C:' -and [datetimeoffset]::TryParse($info.CollectedAtUtc, [ref]([datetimeoffset]::MinValue))) 'The collector must return the drive and the collection time.'
Check (($info.Os -and $info.Cpu -and $info.MemoryGB -gt 0 -and $info.LogicalProcessors -gt 0) -or $info.Error) 'OS, CPU and memory must be read or the failure named.'
Check (($info.SystemDriveSizeGB -gt 0 -and $info.SystemDriveFreePercent -ge 0 -and $info.SystemDriveFreePercent -le 100) -or $info.Error -match 'Drive C:') 'Drive C: size and free percent must be read or the failure named.'

# GUI data: read from the heartbeat by the management module.
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SystemInfo-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
try {
    foreach ($server in @('SRV-A', 'SRV-B', 'SRV-C')) { New-Item -ItemType Directory -Path (Join-Path $root $server) -Force | Out-Null }
    $heartbeat = [ordered]@{
        Timestamp = (Get-Date).ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.68'
        SystemInfo = [ordered]@{ CollectedAtUtc = '2026-10-28T05:00:00.0000000Z'; Os = 'Windows Server 2022 Standard (10.0.20348, build 20348)'; Cpu = 'Xeon Gold'; LogicalProcessors = 8; MemoryGB = 16; SystemDrive = 'C:'; SystemDriveSizeGB = 127.4; SystemDriveFreePercent = 30.0; Error = '' }
        # Orchestrator 1.5.71: the 5-minute sample carries the current drive C: free space.
        ResourceUsage = [ordered]@{ TimeUtc = '2026-10-28T09:25:00.0000000Z'; Cpu = 12.5; Memory = 40.1; SystemDriveFreePercent = 8.5 }
    }
    [IO.File]::WriteAllText((Join-Path $root 'SRV-A/Orchestrator-Heartbeat.json'), ($heartbeat | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    # SRV-B runs an older orchestrator: no SystemInfo in its heartbeat.
    [IO.File]::WriteAllText((Join-Path $root 'SRV-B/Orchestrator-Heartbeat.json'), ([ordered]@{ Timestamp = (Get-Date).ToString('o'); ScriptVersion = '1.5.60' } | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    # SRV-C runs 1.5.69/1.5.70: CPU and memory samples without the drive value.
    $heartbeatC = [ordered]@{ Timestamp = (Get-Date).ToString('o'); ScriptVersion = '1.5.70'; SystemInfo = [ordered]@{ CollectedAtUtc = '2026-10-28T05:00:00.0000000Z'; SystemDriveSizeGB = 200; SystemDriveFreePercent = 22.2 }; ResourceUsage = [ordered]@{ TimeUtc = '2026-10-28T09:25:00.0000000Z'; Cpu = 5; Memory = 50 } }
    [IO.File]::WriteAllText((Join-Path $root 'SRV-C/Orchestrator-Heartbeat.json'), ($heartbeatC | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    $cluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'SRV-C'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; PeerHeartbeatStaleMinutes = 5 }
    $rows = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $root -ClusterDocument $cluster)
    $a = @($rows | Where-Object Server -eq 'SRV-A')[0]
    $b = @($rows | Where-Object Server -eq 'SRV-B')[0]
    $c = @($rows | Where-Object Server -eq 'SRV-C')[0]
    $usageLocal = ([datetimeoffset]::Parse('2026-10-28T09:25:00Z')).LocalDateTime.ToString('yyyy-MM-dd HH:mm')
    $collectedLocal = ([datetimeoffset]::Parse('2026-10-28T05:00:00Z')).LocalDateTime.ToString('yyyy-MM-dd HH:mm')
    Check ($a.Version -eq '1.5.68' -and $a.Os -like 'Windows Server 2022*' -and $a.Cpu -eq 'Xeon Gold (8 logical)' -and $a.MemoryGB -eq 16) 'Version, OS, CPU and RAM must come from the heartbeat.'
    Check ($a.SystemDriveSizeGB -eq 127.4 -and $a.SystemDriveFreePercent -eq 8.5 -and $a.SystemDriveLow -and $a.SystemInfoAt -eq $collectedLocal) "The live drive C: sample must win over the startup reading ($($a.SystemDriveFreePercent), $($a.SystemInfoAt))."
    Check ($a.UsageAt -eq $usageLocal -and $a.CpuPercent -eq 12.5 -and $a.MemoryPercent -eq 40.1) "Usage at must be the time of the latest sample ($($a.UsageAt))."
    Check ($c.SystemDriveFreePercent -eq 22.2 -and -not $c.SystemDriveLow -and $c.UsageAt -eq $usageLocal -and $c.CpuPercent -eq 5) 'A sample without drive value keeps the startup drive reading.'
    Check ($null -eq $b.CpuPercent -and -not $b.UsageAt) 'An orchestrator without samples shows no usage time.'
    Check ($b.Version -eq '1.5.60' -and -not $b.Os -and $null -eq $b.MemoryGB -and -not $b.SystemDriveLow -and -not $b.SystemInfoAt) 'An older orchestrator shows its version and empty system values.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

# GUI: Servers and versions is a sub-tab of Configuration with the system columns; Activity is removed.
$guiText = [IO.File]::ReadAllText((Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
Check ($guiText -match '<TabItem x:Name="ConfigurationServersTab" Header="Servers and versions">' -and $guiText -notmatch '<TabItem Header="Servers and versions">' -and $guiText -notmatch '<TabItem Header="Activity">') 'Servers and versions must be under Configuration and Activity removed.'
# Every control the GUI code reads must exist in the XAML: a removed control (GUI 1.6.13 ActivityBox)
# made the strict-mode GUI fail at startup.
$xamlNames = [Collections.Generic.HashSet[string]]::new([string[]]@([regex]::Matches($guiText, 'x:Name="([A-Za-z0-9_]+)"') | ForEach-Object { $_.Groups[1].Value }))
$controlRefs = @([regex]::Matches($guiText, '\$(?:script:)?Controls(?:\.|\[\x27)([A-Za-z0-9_]+)') | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -notin @('Keys', 'Values', 'ContainsKey', 'Count') } | Sort-Object -Unique)
$unknown = @($controlRefs | Where-Object { -not $xamlNames.Contains($_) })
Check ($controlRefs.Count -gt 100 -and $unknown.Count -eq 0) "GUI code reads controls absent from the XAML: $($unknown -join ', ')"
Check ($guiText -match 'Header="Usage at" Binding="\{Binding UsageAt\}"' -and $guiText -match 'Header="Started at \(static\)" Binding="\{Binding SystemInfoAt\}"' -and $guiText -notmatch 'Header="System info at"') 'Servers grid must show Usage at and Started at (static).'
foreach ($binding in @('Version', 'Os', 'Cpu', 'MemoryGB', 'SystemDriveSizeGB', 'SystemDriveFreePercent', 'SystemInfoAt', 'UsageAt')) {
    Check ($guiText -match ('Binding="\{Binding ' + $binding + '\}"')) "Servers grid column missing: $binding"
}
"PASS: server system information, $($script:passed) offline checks."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD1XBe1cwKewpS2
# X7gobDRMFvi23FMse0nb4vhPq6LYtqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMYzBRs46KbrbkXGoybIksJy4uMSlFTj+yEY7r07yRrOMA0GCSqG
# SIb3DQEBAQUABIIBgEeVb5Xq1IcQAXuN8YrBPwoGWoizWcBBMYT3UmZ34YKzwkSe
# iAgzs2no99T7v72nMbT84jcoBq1MbGF+s8UIY1xJH1UMLu9kQkCT5t4jHn8s+4Wy
# 4qu9maEjP/OBmDtaax+V2oFen6OCSqZJPT137UORgTjOwplEUJVwi8ER2ZcWyI8E
# 5ga7TleicwoO44WM5EzXWk4g2HHK5zgYisADgqfFpHOtMwptaTr+CKs/VrTYv0ah
# iJAN04jfD5uBmfpigwOCn1+g/IJqJinpVEBH8s8C8DcfzTlPQS/XK9toYz95btQh
# YhY6ZqMlFMK7w76Y1K9XMPu4aDTDpZbhlzdDNpnoD66fmoIXmG3TxMSfOe3b//8Q
# Fi2vd/qdWfr7lcpbIz6OhT21n6tTAW6QGAH0b2tXjsp6yKrD2FjK9z09XOHyXaO1
# ps7IH0kx4jPfBSPGe5+NIYQS6z7b3+AeGKv2gSoJ9fLuxEXNbpidum5ycAgcfSyc
# A7eKNXVCTSKZyLV596GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgyMTEz
# MzVaMC8GCSqGSIb3DQEJBDEiBCCCHGsiegKANWIL59ogvx2whhJg4RwXe7F7MKcT
# 7LJqATANBgkqhkiG9w0BAQEFAASCAgAD0sqoDRvueGvZGhMTGbzeYCnHaToM5L/y
# cYBIXJ/NAigMVH0k9957TZTNXqZUFPxZGRFVWaCWNCkVqcxhjPtUSfC2lyl/kLOm
# RKdqHwWWY4fYyccfKEA5To8ZcuAvpDIifR0u0OnblIRTg2GQlTogQ9+65q3zeyk2
# Gsh9tRC6ZFcyD4nleN3kNAxgG0b6VKwP4O4Ptu6PJeRFq1OaNAoQfter3DzW9h75
# k9AY20BgGWDGJKOaNICxzmAR5QTvq0egSjtDDujMUne0jrKOtTYZOzd/y1nwXfLJ
# WqjFJXKx0tXmC6g2JVIZkfqyjdXzN1sT538zVwkiqvrkzYZhJdIspI8qcaoYjJtz
# Dn3DLK8qf+F9w61YxBVdUiHj5GW8cADWsX62+AK+TS0I80XXEUHoLAbzpKz5sK2B
# RN/Mn+mfJ5LgAV267XEAdXe/zhKjN0CedeEIaCYoAXjlIMIocTUqIxNKu4Xz5zrD
# 7sAFEM0ZPQdJ9cu0XGer1uo8DeTqvC84tpau/SHXdpmk7qLVwT92MkVYgrWHe99K
# ZWEnoj1WTatdkCFJozvq7CSz9TXBii4IQ4K8uLJKdOCyDmM7MZI6p73GfZoWUDOV
# GasMB54m5s6T9dRUpL9YQS907cA4EzrUx90/eqNUvYjT1IL4dhspclQ6bTyJZQLL
# iM6lKjUGOw==
# SIG # End signature block
