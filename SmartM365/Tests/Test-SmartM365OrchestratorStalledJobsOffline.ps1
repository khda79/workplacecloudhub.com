<#
.SYNOPSIS
Offline tests of the stalled-job alert: a running job whose log stays silent beyond MaxSilentMinutes
is reported once per run (log line, email, GUI and Dashboard) and is never stopped automatically.
.VERSION
1.0.0
.NOTES
Synthetic job log and running state in a temporary folder; one short child pwsh of this test
measures the process tree CPU. No tenant initialization, collectors, scheduling or publication.
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

# Wiring: checked for every supervised job after the timeout handling; never a stop or a kill.
$update = Get-FunctionText 'Update-RunningJobs'
Check ($update -match 'Update-RunningJobActivity -JobName \$name -RunInfo \$info -Now \$Now') 'Update-RunningJobs no longer checks the job activity.'
$activity = Get-FunctionText 'Update-RunningJobActivity'
Check ($activity -notmatch 'Stop-ProcessTree|taskkill|Stop-Process|Complete-JobRun|TimeoutRequested') 'The stalled-job check must never stop a job.'
Check ((Get-FunctionText 'Write-OrchestratorHeartbeat') -match 'Stalled = ' -and (Get-FunctionText 'Write-OrchestratorHeartbeat') -match 'CpuSecondsPerMinute = ') 'The heartbeat no longer publishes the stalled-job signals.'
Check ((Get-FunctionText 'Restore-RunningJobs') -match 'StallAlertSent = ') 'The once-per-run alert flag is not restored after an orchestrator restart.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Stalled-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root -Force | Out-Null
$child = $null
try {
    $functions = @('Get-OrchestratorProcessTreeCpuTime', 'Get-OrchestratorJobLogActivity', 'Update-RunningJobActivity')
    $stubs = @'
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Logs.Add(('{0} {1}' -f $Level, $Message)) }
function Get-JobState { param($JobName) $script:JobState }
function Save-OrchestratorState { $script:Saves++ }
function Send-JobStalledEmail { param($JobName, $RunInfo) $script:Mails.Add($JobName); $true }
'@
    $module = New-Module -ScriptBlock ([scriptblock]::Create((@($functions | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs))
    & $module {
        $script:Logs = [Collections.Generic.List[string]]::new(); $script:Mails = [Collections.Generic.List[string]]::new(); $script:Saves = 0
        $script:JobState = @{ Running = @{} }
        $script:JobActivityCheckSeconds = 60
        $script:Manifest = [pscustomobject]@{ JobsByName = @{ 'Slow-Job' = [pscustomobject]@{ MaxSilentMinutes = 60 }; 'Quiet-Job' = [pscustomobject]@{ MaxSilentMinutes = 0 } } }
    }

    # Process tree CPU: the launcher and its child pwsh are both counted.
    $child = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList '-NoProfile', '-Command', '$end=(Get-Date).AddSeconds(2); while((Get-Date) -lt $end){}; Start-Sleep -Seconds 60' -PassThru -WindowStyle Hidden
    Start-Sleep -Seconds 3
    $treeCpu = & $module { param($p) Get-OrchestratorProcessTreeCpuTime -RootPid $p } $PID
    $selfCpu = (Get-Process -Id $PID).TotalProcessorTime.TotalSeconds
    Check ($treeCpu -ge $selfCpu + 1) "The process tree CPU must include the child process ($treeCpu s, this process $selfCpu s)."

    # Log silence: the size read through a shared handle drives the last activity.
    $logPath = Join-Path $root 'Job-Slow-Job_SRV-A_20261029_043000.log'
    [IO.File]::WriteAllText($logPath, "start`n")
    $start = (Get-Date).AddMinutes(-10)
    [IO.File]::SetLastWriteTimeUtc($logPath, $start.ToUniversalTime().AddMinutes(2))
    $runInfo = @{ StartTime = $start; LogPath = $logPath; Process = $child }
    $t0 = Get-Date
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t0 $runInfo
    Check (-not $runInfo.Stalled -and [math]::Abs($runInfo.SilentMinutes - 8) -lt 0.2 -and $runInfo.MaxSilentMinutes -eq 60 -and $null -ne $runInfo.CpuSeconds) "First check: silence measured from the last log write ($($runInfo.SilentMinutes) min)."
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t0.AddSeconds(30) $runInfo
    Check ([math]::Abs($runInfo.SilentMinutes - 8) -lt 0.2) 'The check must run at most once a minute.'

    # 61 minutes later without any write: stalled, one WARN line and one email, never a stop.
    $t1 = $t0.AddMinutes(53)
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t1 $runInfo
    $logs = @(& $module { $script:Logs })
    Check ($runInfo.Stalled -and $runInfo.SilentMinutes -ge 60 -and @($logs | Where-Object { $_ -like 'WARN Job Slow-Job: stalled*Not stopped*' }).Count -eq 1) "A silent job must be reported as stalled ($($logs -join ' | '))."
    Check ((& $module { $script:Mails.Count }) -eq 1 -and $runInfo.StallAlertSent -and (& $module { $script:JobState.Running.StallAlertSent }) -and (& $module { $script:Saves }) -eq 1) 'The alert must be emailed once and persisted with the running state.'
    Check ($null -ne $runInfo.CpuSecondsPerMinute -and $runInfo.CpuSecondsPerMinute -ge 0) 'The CPU progression per minute must be measured.'
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t1.AddMinutes(5) $runInfo
    Check ((& $module { $script:Mails.Count }) -eq 1 -and @(& $module { $script:Logs } | Where-Object { $_ -like 'WARN*stalled*' }).Count -eq 1) 'A still-stalled job must not be alerted again.'

    # The log moves again: the state clears with an INFO line; a second silence in the same run sends no new email.
    [IO.File]::AppendAllText($logPath, "progress`n")
    $t2 = $t1.AddMinutes(10)
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t2 $runInfo
    Check (-not $runInfo.Stalled -and $runInfo.SilentMinutes -eq 0 -and @(& $module { $script:Logs } | Where-Object { $_ -like 'INFO Job Slow-Job: log activity resumed*' }).Count -eq 1) 'Log activity must clear the stalled state.'
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t2.AddMinutes(61) $runInfo
    Check ($runInfo.Stalled -and (& $module { $script:Mails.Count }) -eq 1) 'The email is sent once per run, even after a second silence.'

    # MaxSilentMinutes = 0 disables the alert; a restored run that already alerted sends no email.
    $quiet = @{ StartTime = (Get-Date).AddHours(-5); LogPath = $logPath; Process = $null }
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Quiet-Job' -RunInfo $r -Now $n } (Get-Date).AddHours(3) $quiet
    Check (-not $quiet.Stalled -and $quiet.MaxSilentMinutes -eq 0) 'MaxSilentMinutes = 0 must disable the alert.'
    $restored = @{ StartTime = (Get-Date).AddHours(-3); LogPath = (Join-Path $root 'missing.log'); Process = $null; StallAlertSent = $true }
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } (Get-Date) $restored
    Check ($restored.Stalled -and (& $module { $script:Mails.Count }) -eq 1) 'A restored run without log is measured from its start and keeps its once-per-run flag.'

    # Manifest: default 60, negative refused by the orchestrator and by the GUI publication check.
    $parse = Get-FunctionText 'ConvertTo-NormalizedJob'
    Check ($parse -match '\$maxSilentMinutes = 60' -and $parse -match 'MaxSilentMinutes cannot be negative') 'The manifest must default MaxSilentMinutes to 60 and refuse a negative value.'
    $management = [IO.File]::ReadAllText((Join-Path $base 'SmartM365.Orchestrator.Management.psm1'))
    Check ($management -match "'MaxSilentMinutes'") 'The GUI publication check must refuse a negative MaxSilentMinutes.'

    # GUI side: heartbeat signals become Running rows and a Critical Dashboard item.
    New-Item -ItemType Directory -Path (Join-Path $root 'Shared/SRV-A'), (Join-Path $root 'Shared/SRV-B') -Force | Out-Null
    $now = Get-Date
    $heartbeat = [ordered]@{ Timestamp = $now.ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.72'; JobStopProtocol = 1
        RunningJobs = @(
            [ordered]@{ Name = 'Slow-Job'; Pid = 4242; StartTime = $now.AddMinutes(-110).ToString('o'); ScheduledOccurrence = $now.AddMinutes(-110).ToString('o'); LastLogWriteUtc = $now.ToUniversalTime().AddMinutes(-75).ToString('o'); SilentMinutes = 75; MaxSilentMinutes = 60; Stalled = $true; CpuSeconds = 58.2; CpuSecondsPerMinute = 0.3 }
            [ordered]@{ Name = 'Busy-Job'; Pid = 4243; StartTime = $now.AddMinutes(-5).ToString('o'); ScheduledOccurrence = $now.AddMinutes(-5).ToString('o'); SilentMinutes = 0.5; MaxSilentMinutes = 60; Stalled = $false; CpuSeconds = 40; CpuSecondsPerMinute = 9.5 }
        ) }
    [IO.File]::WriteAllText((Join-Path $root 'Shared/SRV-A/Orchestrator-Heartbeat.json'), ($heartbeat | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    # SRV-B runs an older orchestrator without the signals.
    $old = [ordered]@{ Timestamp = $now.ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.71'; RunningJobs = @([ordered]@{ Name = 'Old-Job'; Pid = 5000; StartTime = $now.AddMinutes(-3).ToString('o') }) }
    [IO.File]::WriteAllText((Join-Path $root 'Shared/SRV-B/Orchestrator-Heartbeat.json'), ($old | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Global
    Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force
    $cluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B'); PeerHeartbeatStaleMinutes = 5 }
    $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath (Join-Path $root 'Shared') -ClusterDocument $cluster -Now $now
    $slow = @($operations.Running | Where-Object Job -eq 'Slow-Job')[0]
    $busy = @($operations.Running | Where-Object Job -eq 'Busy-Job')[0]
    $oldRow = @($operations.Running | Where-Object Job -eq 'Old-Job')[0]
    Check ($slow.Activity -eq 'Stalled' -and $slow.SilentMinutes -eq 75 -and $slow.CpuSeconds -eq 58.2 -and $slow.CpuSecondsPerMinute -eq 0.3) 'A stalled run must reach the Running rows with its silence and CPU.'
    Check ($busy.Activity -eq 'Active' -and $oldRow.Activity -eq '' -and $null -eq $oldRow.SilentMinutes) 'An active run and an older orchestrator must be shown as such.'
    $insights = [IO.File]::ReadAllText((Join-Path $base 'SmartM365.Orchestrator.Insights.psm1'))
    Check ($insights -match "Add-DashboardAction 'Critical' 'Stalled job'") 'The Dashboard must list a stalled job as Critical.'

    $guiText = [IO.File]::ReadAllText((Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
    foreach ($binding in @('Activity', 'SilentMinutes', 'CpuSeconds', 'CpuSecondsPerMinute')) { Check ($guiText -match ('Binding="\{Binding ' + $binding + '\}"')) "Running jobs column missing: $binding" }
    Check ($guiText -match '<DataTrigger Binding="\{Binding Activity\}" Value="Stalled">') 'A stalled run must be highlighted in Running jobs.'
}
finally {
    if ($null -ne $child) { try { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue } catch { } }
    Remove-Module -Name SmartM365.Orchestrator.Insights -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: stalled-job alert, $($script:passed) offline checks."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDH4IcTi2WTESLR
# Mmf7syIvyr4nXTKav+3WWuq8yNgQZKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILHo8sRnP8HGNQnUo3FoBZP8rEQAqXDdv4HY1G/x73I4MA0GCSqG
# SIb3DQEBAQUABIIBgCSXxuwnr8YcZjfpZGxbpPtLU5lEROijh5dbIa+EgJZRyk4L
# FaxPEToFHYHxVLVzpz056+Z5YWl/H/t7r+OlRXHyCoZoN9aMfE+/BMqopy6Qm8SW
# w4h/ljzclo3eKkgJ+WGQn8C+N7TcN/FhT0vNePmsyZSZSGcoHYA6RrErgQOmsEXh
# Bh1z3A7SMCCRpkknJB0lvy3yQMzovyJHgiWWEa72kSDHsybCjkFMKHClv3YG3Y+C
# Kghi8CLAsgsYHyQBJzEvGohlwQD+ZXHa1GDgZZHYT/wIC/sITf0/yFA/tukHqnIk
# Di5hnUjAoMWq/Y3ajs4N2Q8jJY7/vF3+4rkngLr8LUQBC4Fqp3ZohHlTD0Ll6Tiq
# gj271ewvk6q1NP4Mip/sGywyVamvIhlmWWIACucZUicKvlyBUbysM8bRryhKbeWs
# P7Y+ByxSccGfOAoNt5O7d823i8IcsfejTsohUk1IXK+gJwD8tTyd8HVojGfwlq82
# gL7Gmi3qk88//hvF9aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkwOTMw
# NTBaMC8GCSqGSIb3DQEJBDEiBCDmsUBm9hEVr0pYKZ8Z9YpcXD7OIzLncxRiwYAc
# WvbftjANBgkqhkiG9w0BAQEFAASCAgCVDZKeiQKR6hdhXTIfCRFS0mPD7toiK2+7
# SwSg6R57ESIhyEXuGjASqeFEpSYohicDaKyXAe0TOLfUr25zWQEq5/ojtrAG8BeF
# gChmUiw7Revko5xDOROmmfkfuvv/Zn79eowa55uTLl02tn2giJdYIle3B+LAuy10
# xvdxYHx1Qhp+aD3QoQ3rp7I12svGhC1Z5HBlcFxscP5X/1/mwgeeaqKVP9whrygp
# /8lTxv2WLcasl9NYMnC7nEqQrD9H0jNoecCyLMi6/EdmtyFabv2Qy/Vpuc2Nm4Xe
# 8xzcw3N5xVVELalFm3+bLSVvqcbMvwJqqM2WtNiZ+H+qMrwh+51cZj0zNrC+nKee
# XW8UuDHAQLOtPMrtjDjkKL4t8BzcKqy6YcnaKH0ukzK9mFBjYHlZR2od668UuGu5
# 1s395zrJdLNRlH+QNYmvq3PtrpoOjl5ChN+i4rKRfhM36Fz8Ijq2qUU6TXnVHDi/
# Qqn7gYZ2xB3iyBD1xw6bPQe3AK4kptZ87K7ataZsBIYWOvUELhBDe2kCmNVhf/c8
# 1sj1McfWmCRge+txY55k5yzOJENYfTh1pCndZMdgkme5roAyFfzfl+btiZSwBNe6
# DnDEsAn+3IU+UJ2u8fc1N84XbtiiwPSlaUBmgZf5QFnLVt2R0iIqEfWbJN4Y6GFo
# Iw9xVoaGIg==
# SIG # End signature block
