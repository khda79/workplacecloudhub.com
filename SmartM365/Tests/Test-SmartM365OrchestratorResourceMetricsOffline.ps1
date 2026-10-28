<#
.SYNOPSIS
Offline tests of the CPU and memory curves: each orchestrator samples every 5 minutes, keeps only
the last 12 hours in Orchestrator-Metrics.json and publishes the latest sample in its heartbeat;
the GUI reads the file and draws one chart per server in the Dashboard and in Servers and versions.
.VERSION
1.0
.NOTES
Synthetic samples in a temporary folder; one real read-only CIM sample of this computer. No tenant
initialization, collectors, scheduling or publication.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
function Get-FunctionText($Tree, [string]$Name) {
    $definition = $Tree.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing function: $Name" }
    $definition.Extent.Text
}

# Wiring: sampled in the tick loop before the heartbeat, reloaded at startup, latest sample published.
$source = [IO.File]::ReadAllText($orchestratorPath)
Check ($source -match 'Update-OrchestratorResourceMetrics -Now \$now\s+Write-OrchestratorHeartbeat') 'The tick loop no longer samples CPU and memory before the heartbeat.'
Check ($source -match 'Import-OrchestratorResourceMetrics -Now \(Get-Date\)') 'Startup no longer reloads the previous samples.'
Check ((Get-FunctionText $ast 'Write-OrchestratorHeartbeat') -match 'ResourceUsage = ') 'The heartbeat no longer publishes the latest sample.'
Check ((Get-FunctionText $ast 'Write-FileAtomically') -match 'ElectionPlan\|Metrics') 'The metrics file is not written as a regenerated snapshot.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Metrics-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path (Join-Path $root 'SRV-A') -Force | Out-Null
try {
    $functions = @('Get-OrchestratorMetricsPath', 'Get-OrchestratorResourceSample', 'Remove-OrchestratorExpiredResourceSamples', 'Import-OrchestratorResourceMetrics', 'Update-OrchestratorResourceMetrics')
    $stubs = @'
function Write-FileAtomically { param([string]$Path, [string]$Content) [IO.File]::WriteAllText($Path, $Content, [Text.UTF8Encoding]::new($false)); $script:Writes++ }
function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now) $script:Warnings.Add($Message) }
'@
    $text = (@($functions | ForEach-Object { Get-FunctionText $ast $_ }) -join "`n") + "`n" + $stubs
    $module = New-Module -ScriptBlock ([scriptblock]::Create($text))
    Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Global
    function Reset-Orchestrator {
        & $module {
            param($Root)
            $script:Settings = [pscustomobject]@{ HeartbeatPath = (Join-Path $Root 'SRV-A/Orchestrator-Heartbeat.json') }
            $script:ResourceSamples = [Collections.Generic.List[object]]::new()
            $script:LastResourceSampleAt = [datetime]::MinValue
            $script:ResourceSampleMinutes = 5; $script:ResourceRetentionHours = 12
            $script:Writes = 0; $script:Warnings = [Collections.Generic.List[string]]::new()
        } $root
    }
    Reset-Orchestrator
    $now = Get-Date

    # A real sample of this computer: percentages between 0 and 100.
    $sample = & $module { param($n) Get-OrchestratorResourceSample -Now $n } $now
    Check ($null -ne $sample.Cpu -and $sample.Cpu -ge 0 -and $sample.Cpu -le 100 -and $sample.Memory -gt 0 -and $sample.Memory -le 100) "CPU and memory must be read as percentages ($($sample.Cpu), $($sample.Memory))."

    # Samples older than 12 hours are dropped when the file is rewritten.
    & $module {
        param($n)
        $script:ResourceSamples.Add([pscustomobject]@{ TimeUtc = $n.ToUniversalTime().AddHours(-13).ToString('o'); Cpu = 10; Memory = 20 })
        $script:ResourceSamples.Add([pscustomobject]@{ TimeUtc = $n.ToUniversalTime().AddHours(-1).ToString('o'); Cpu = 30; Memory = 40 })
        Update-OrchestratorResourceMetrics -Now $n
    } $now
    $metricsPath = Join-Path $root 'SRV-A/Orchestrator-Metrics.json'
    $document = Get-Content -LiteralPath $metricsPath -Raw | ConvertFrom-Json
    Check (@($document.Samples).Count -eq 2 -and $document.RetentionHours -eq 12 -and $document.SampleMinutes -eq 5 -and (& $module { $script:Writes }) -eq 1) 'The file must keep only the last 12 hours plus the new sample.'
    & $module { param($n) Update-OrchestratorResourceMetrics -Now $n.AddMinutes(4) } $now
    Check ((& $module { $script:Writes }) -eq 1 -and (& $module { $script:ResourceSamples.Count }) -eq 2) 'A sample must not be taken before 5 minutes.'
    & $module { param($n) Update-OrchestratorResourceMetrics -Now $n.AddMinutes(5) } $now
    Check ((& $module { $script:Writes }) -eq 2 -and (& $module { $script:ResourceSamples.Count }) -eq 3) 'A sample must be taken every 5 minutes.'
    Check ((& $module { $script:Warnings.Count }) -eq 0) "Unexpected warning: $(& $module { $script:Warnings -join '; ' })"

    # After a restart the curve continues from the file; stale samples are dropped on import.
    Reset-Orchestrator
    & $module { param($n) Import-OrchestratorResourceMetrics -Now $n.AddHours(11.5) } $now
    Check ((& $module { $script:ResourceSamples.Count }) -eq 2) 'Import must reload the samples still within 12 hours.'

    # Heartbeat values and file reading on the GUI side.
    $heartbeat = [ordered]@{ Timestamp = (Get-Date).ToString('o'); ScriptVersion = '1.5.69'; ResourceUsage = [ordered]@{ TimeUtc = $now.ToUniversalTime().ToString('o'); Cpu = 42.5; Memory = 61.2 } }
    [IO.File]::WriteAllText((Join-Path $root 'SRV-A/Orchestrator-Heartbeat.json'), ($heartbeat | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
    Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force
    $cluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; PeerHeartbeatStaleMinutes = 5 }
    $status = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $root -ClusterDocument $cluster)[0]
    Check ($status.CpuPercent -eq 42.5 -and $status.MemoryPercent -eq 61.2) 'Servers and versions must show the latest sample of the heartbeat.'
    $metrics = Get-SmartM365OrchestratorServerMetrics -ServerFolder (Join-Path $root 'SRV-A') -Server 'SRV-A' -Now $now.AddMinutes(6)
    Check ($metrics.Server -eq 'SRV-A' -and @($metrics.Samples).Count -eq 3 -and $metrics.SampleMinutes -eq 5 -and @($metrics.Samples)[0].Time -is [datetime]) 'The GUI must read the samples of the last 12 hours.'
    $later = Get-SmartM365OrchestratorServerMetrics -ServerFolder (Join-Path $root 'SRV-A') -Server 'SRV-A' -Now $now.AddHours(11.5)
    Check (@($later.Samples).Count -eq 2) 'The GUI must not show samples older than 12 hours.'
    $missing = Get-SmartM365OrchestratorServerMetrics -ServerFolder (Join-Path $root 'SRV-X') -Server 'SRV-X' -Now $now
    Check (@($missing.Samples).Count -eq 0) 'A server without samples must give an empty curve.'

    # Drawing: one panel per server, a line per series, broken by a gap.
    $guiPath = Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'
    $guiAst = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$null, [ref]$null)
    foreach ($name in @('ConvertTo-OrchestratorBrush', 'Update-OrchestratorServerMetricsChart')) { . ([scriptblock]::Create((Get-FunctionText $guiAst $name))) }
    $clock = [datetime]'2026-10-18 09:30'
    $samples = @(for ($minute = 720; $minute -ge 0; $minute -= 5) { if ($minute -le 300 -and $minute -gt 240) { continue }; [pscustomobject]@{ Time = $clock.AddMinutes(-$minute); Cpu = 30; Memory = 60 } })
    $canvas = [System.Windows.Controls.Canvas]::new(); $canvas.Width = 900; $canvas.Height = 130
    $canvas.Measure([System.Windows.Size]::new(900, 130)); $canvas.Arrange([System.Windows.Rect]::new(0, 0, 900, 130))
    $canvas.Tag = @([pscustomobject]@{ Server = 'SRV-A'; From = $clock.AddHours(-12); To = $clock; SampleMinutes = 5; Samples = $samples }, [pscustomobject]@{ Server = 'SRV-B'; From = $clock.AddHours(-12); To = $clock; SampleMinutes = 5; Samples = @() })
    Update-OrchestratorServerMetricsChart -Canvas $canvas
    $lines = @($canvas.Children | Where-Object { $_ -is [System.Windows.Shapes.Polyline] })
    Check ($lines.Count -eq 4 -and @($lines | Where-Object { $_.Points.Count -lt 2 }).Count -eq 0) "A one-hour gap must split CPU and RAM into two lines each ($($lines.Count))."
    $titles = @($canvas.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and $_.FontWeight -eq [System.Windows.FontWeights]::SemiBold } | ForEach-Object Text)
    Check ($titles.Count -eq 2 -and $titles[0] -like 'SRV-A  CPU 30% *' -and $titles[1] -eq 'SRV-B  no sample') "Each server needs its own titled panel ($($titles -join ' | '))."
    $canvas.Tag = @()
    Update-OrchestratorServerMetricsChart -Canvas $canvas
    Check (@($canvas.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and $_.Text -like 'No CPU or memory sample yet*' }).Count -eq 1) 'An empty chart must explain why.'

    # Both views host the chart.
    $guiText = [IO.File]::ReadAllText($guiPath)
    Check ($guiText -match 'x:Name="DashboardMetricsCanvas"' -and $guiText -match 'x:Name="ServersMetricsCanvas"' -and $guiText -match 'DashboardMetricsCanvas\.Add_SizeChanged' -and $guiText -match 'ServersMetricsCanvas\.Add_SizeChanged') 'The Dashboard and Servers and versions must both show the chart.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: CPU and memory curves, $($script:passed) offline checks."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDP/V3OUNIEAc2f
# q5pfXtmsul4Osi40EJAUjiYEgek9pqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBE6C3sZwkj2mODAXigVZAvsD2KBm86IwjNELkeXloZkMA0GCSqG
# SIb3DQEBAQUABIIBgGUlvn3X5XJplwPtRUom9ICo8SJyvLl6CCzKC8BFZxz1ZZ+6
# W7S6OsHITdnZlPnpo5Df6SIjnjVtnTkSenlYWszQ4rXFfs6CU7/9ihcWSytFbnXY
# Ex05MKbbugwRdb70xsc3nwuzJc3BMo7reO5T7mCsWtWz6OVTmaUcoDnHuNV00R89
# VMe5HnrLcOuUxaD0Z86fx/2kxp1xh6yJxkEVFIOWKTt/m+oNsGmth94i1mpQ8XEl
# pgu/yhSIP7TwsLYfptctYxtTHvvMQmOgTuz/6x7c6d4Yliy1u612M+MPVF2KRy9l
# E7D4FB+IAVdXlPEwii93ZGTw8obKgYUy0suunx/u7YOZ0aYotm2jm6vcT37D/VbH
# u2tJfh7phn2XBRmYXySilojfB2b6qaD9UsPEil0IM8yd2odvWpAp9h1S+rnrj8vS
# h5YOh37UJPYNl565uL44IY8xKo+8VoQH6WXV2J+wfmid9sS9iBLIfqfI9ZNkbuQG
# ZXvr3B+4cB0RojvGIKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgyMTA2
# MThaMC8GCSqGSIb3DQEJBDEiBCDclJlU66e3WOL1p06QW5MdWblbC2a7RX+TyN89
# p1RcUjANBgkqhkiG9w0BAQEFAASCAgCp97TGq5Q7IuCwyblF40JEq1I2fDmzg/Da
# vnYYrN/r8wi6K2mPqgAAPYuDcO/9DVdJWqRBYzqjxuyteGgRHQnGF2+r4WHvrEH0
# C6E+VaDvYq1eOX3Sjq9fat1XWbGg4EhttBClO9XMqbL3jaiH0by+/wD0VEB3GRBm
# t59erbE3bksb4Jrc7QszGhHi4TrCFAk21x9JfmR2zs+omtHotfw2JyA/HsaKTq2s
# ffrb6fPgUiVK8uYd4J72KH3n1mCKZHFm59JeMeCwPx/j5p4grkZhcOVTzE3fE9vj
# H60DqYVNIVt0cAttMFwJ8mkmTXAhAq5zt1mmwHrofClX/hNycDSnbh6TGrmcgAMB
# jyeVPALRIX8O8VRk9bf4LtFJaOOYiadKh2T3HcesN25jqnLyDYsrwbZPfq7kNu5D
# 2He7Ex/Un/Xth+C7kOx9t62m6VCEESxA9esHIBQP4N3xKBVDxzSVgj82DPYE2SxR
# xxgACt/78tuok1nqCVccSUu07qgePTiAp0VP0qz6vHW2SXOg8/83zMzzHf9gFIw9
# qT78U9f1lyd3wvHNStoYlGUPjv2qGPt/+CqXhfzjj4dBSkxfHjWlpcLoDBI80zq7
# SskmBNTtPGjHFYFBjYKr2JQ/gxJLLFA1vg+MVCGCn2KUGKrkuivVRebhHnnITxW5
# TB5ZlH816Q==
# SIG # End signature block
