<#
.SYNOPSIS
Offline checks for pinned FreshSuccess dependencies across runtime, pipeline and GUI.
.VERSION
1.0.2
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force
$parseErrors=$null; $tokens=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'),[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count) { throw 'Orchestrator does not parse.' }
$definitions=foreach($name in @('Get-OrchestratorEffectiveDependencyMode','Get-OrchestratorDependencyFreshStatus','Get-OrchestratorExternalDependencyStatus','Get-OrchestratorDependencyMaxAge','Get-JobOccurrencesInWindow','Get-OrchestratorDependencyActivityText')) {
    $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true).Extent.Text
}
# Execute the real scheduled branch in a minimal loop; never execute the entry script.
$scheduled=$ast.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -like '*depJob.AssignmentMode*Pinned*FreshSuccess*'},$true)
if (-not $scheduled) { throw 'Pinned scheduled dependency gate is absent.' }
$definitions += 'function ScheduledStatus($job,$depJob,$Now) { $deferred=$false; $blockingDependencies=[Collections.Generic.List[string]]::new(); foreach($dep in @($depJob.Name)) { $dependencyMode=Get-OrchestratorEffectiveDependencyMode -Job $job -DependencyName $dep; ' + $scheduled.Extent.Text + ' }; if($deferred){"Waiting"}else{"Ready"} }'
$definitions += 'function Write-OrchestratorRuntimeUpdateWarning { param($Key,$Message,$Now) throw $Message }'
$definitions += '$script:DependencyWaitReason = @{}'
$gate=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-PinnedDeps-'+[guid]::NewGuid().ToString('N'))
$shared=Join-Path $root 'DATA-ALL/Orchestrator'
$ownerFolder=Join-Path $shared 'WorkerA'
$heartbeatPath=Join-Path $ownerFolder 'Orchestrator-Heartbeat.json.txt'
$statePath=Join-Path $ownerFolder 'Orchestrator-State.json.txt'
$now=[datetime]::new(2026,10,15,10,0,0,[DateTimeKind]::Utc)
$parent=[pscustomobject]@{Name='Synthetic-Inventory';Enabled=$true;AssignmentMode='Pinned';AllowedServers=@('WorkerA');ConcurrencyKey='Synthetic-Inventory';ContinueOnError=$true;Schedule=[pscustomobject]@{Type='Daily';Times=@('22:00');DaysOfWeek=@()}}
$consumer=[pscustomobject]@{Name='Consumer';DependsOn=@($parent.Name);DependencyMode='FreshSuccess';DependencyMaxAgeHours=48}
$jobs=[pscustomobject]@{Jobs=@($parent,$consumer)}
$oldProfile=Get-Variable SmartM365ProfileKey -Scope Global -ValueOnly -ErrorAction SilentlyContinue
$script:checks=0
function Check($Value,$Message) { if(-not $Value){throw $Message}; $script:checks++ }
function Write-Fixture {
    [IO.File]::WriteAllText($heartbeatPath,($heartbeat | ConvertTo-Json -Depth 8))
    [IO.File]::WriteAllText($statePath,($state | ConvertTo-Json -Depth 8))
}
function Verify-Gates([string]$Expected,[string]$Label) {
    $before=@(Get-ChildItem -LiteralPath $shared -File -Recurse | Sort-Object FullName | ForEach-Object { $_.FullName + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash })
    $runtime=& $gate {param($p,$n) Get-OrchestratorDependencyFreshStatus -DependencyJob $p -MaxAgeHours 48 -Now $n} $parent $now
    Check ($runtime -eq $Expected) "Runtime: $Label ($runtime)."
    $external=& $gate {param($c,$p,$n) Get-OrchestratorExternalDependencyStatus -Job $c -DependencyName $p.Name -Now $n} $consumer $parent $now
    Check ($external -eq $Expected) "Pipeline external dependency: $Label ($external)."
    $scheduledResult=& $gate {param($c,$p,$n) ScheduledStatus $c $p $n} $consumer $parent $now
    Check ($scheduledResult -eq $Expected) "Scheduled dependency: $Label ($scheduledResult)."
    $gui=@(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $shared -JobsDocument $jobs -JobName Consumer -Now $now)
    Check ($gui.Count -eq 1 -and $gui[0].State -eq $Expected) "GUI: $Label ($($gui[0].State))."
    $after=@(Get-ChildItem -LiteralPath $shared -File -Recurse | Sort-Object FullName | ForEach-Object { $_.FullName + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash })
    Check (($before -join '|') -ceq ($after -join '|')) "Read-only gate modified shared evidence: $Label."
}
try {
    New-Item -ItemType Directory -Path $ownerFolder -Force | Out-Null
    $global:SmartM365ProfileKey='synthetic'
    & $gate {param($path,$p) $script:Settings=[pscustomobject]@{SharedDataFolderPath=$path;PeerHeartbeatStaleMinutes=5}; $script:Manifest=[pscustomobject]@{JobsByName=@{$p.Name=$p}}; $script:DependencyFreshCache=@{}} $shared $parent
    $heartbeat=@{Timestamp=$now.AddMinutes(-1).ToString('o');Tenant='synthetic';Lifecycle='Running';StatePersistenceHealthy=$true;RunningJobs=@()}
    $state=@{SchemaVersion=1;Jobs=@{$parent.Name=@{LastRunStart=$now.AddHours(-2).ToString('o');LastRunEnd=$now.AddHours(-1).ToString('o');LastStatus='Success';Running=$null;PendingRetry=$null}}}
    $heartbeatBase=$heartbeat|ConvertTo-Json -Depth 8; $stateBase=$state|ConvertTo-Json -Depth 8
    foreach($case in @(
        @{Name='fresh success';Expected='Ready';Change={}},
        @{Name='warnings';Expected='Ready';Change={$state.Jobs[$parent.Name].LastStatus='CompletedWithWarnings'}},
        @{Name='heartbeat running';Expected='Waiting';Change={$heartbeat.RunningJobs=@(@{Name=$parent.Name})}},
        @{Name='state running';Expected='Waiting';Change={$state.Jobs[$parent.Name].Running=@{Pid=123}}},
        @{Name='retry';Expected='Waiting';Change={$state.Jobs[$parent.Name].PendingRetry=@{Attempt=1}}},
        @{Name='failed';Expected='Waiting';Change={$state.Jobs[$parent.Name].LastStatus='Failed'}},
        @{Name='timed out';Expected='Waiting';Change={$state.Jobs[$parent.Name].LastStatus='TimedOut'}},
        @{Name='cancelled';Expected='Waiting';Change={$state.Jobs[$parent.Name].LastStatus='Cancelled'}},
        @{Name='interrupted';Expected='Waiting';Change={$state.Jobs[$parent.Name].LastStatus='Interrupted'}},
        @{Name='stale success';Expected='Waiting';Change={$state.Jobs[$parent.Name].LastRunStart=$now.AddHours(-51).ToString('o');$state.Jobs[$parent.Name].LastRunEnd=$now.AddHours(-50).ToString('o')}},
        @{Name='future success';Expected='Waiting';Change={$state.Jobs[$parent.Name].LastRunEnd=$now.AddHours(1).ToString('o')}},
        @{Name='offsetless time';Expected='Waiting';Change={$state.Jobs[$parent.Name].LastRunEnd='2026-10-15T09:00:00'}},
        @{Name='stale heartbeat';Expected='Waiting';Change={$heartbeat.Timestamp=$now.AddMinutes(-6).ToString('o')}},
        @{Name='future heartbeat';Expected='Waiting';Change={$heartbeat.Timestamp=$now.AddMinutes(2).ToString('o')}},
        @{Name='unhealthy state';Expected='Waiting';Change={$heartbeat.StatePersistenceHealthy=$false}},
        @{Name='wrong profile';Expected='Waiting';Change={$heartbeat.Tenant='other'}},
        @{Name='recycling owner';Expected='Waiting';Change={$heartbeat.Lifecycle='Recycling'}},
        @{Name='unknown job';Expected='Waiting';Change={$state.Jobs=@{}}}
    )) {
        $heartbeat=$heartbeatBase|ConvertFrom-Json -AsHashtable; $state=$stateBase|ConvertFrom-Json -AsHashtable
        & $case.Change; Write-Fixture
        Verify-Gates $case.Expected $case.Name
    }
    $heartbeat=$heartbeatBase|ConvertFrom-Json -AsHashtable; $state=$stateBase|ConvertFrom-Json -AsHashtable; Write-Fixture
    $parent.AllowedServers=@('WorkerA','WorkerB'); Verify-Gates Waiting 'ambiguous owner'; $parent.AllowedServers=@('WorkerA')
    [IO.File]::WriteAllText($statePath,'{'); Verify-Gates Waiting 'invalid state'; Write-Fixture
    Remove-Item -LiteralPath $statePath; Verify-Gates Waiting 'missing state'; Write-Fixture
    Remove-Item -LiteralPath $heartbeatPath; Verify-Gates Waiting 'missing heartbeat'; Write-Fixture
    # Synthetic lease is never acquired through the operational scheduler.
    $leaseFolder=Join-Path $shared 'Election/Concurrency'
    New-Item -ItemType Directory -Path $leaseFolder -Force | Out-Null
    $leasePath=Join-Path $leaseFolder 'Synthetic-Inventory.json.txt'
    [IO.File]::WriteAllText($leasePath,'{}'); Verify-Gates Waiting 'lease before state update'
    Remove-Item -LiteralPath $leasePath
    Verify-Gates Ready 'lease cleared'
    # A cached successful decision must not override a newly running pinned producer.
    & $gate {param($n) $script:DependencyFreshCache['Synthetic-Inventory|48']=@{Status='Ready';Until=$n.AddHours(1)}} $now
    $heartbeat.RunningJobs=@(@{Name=$parent.Name}); Write-Fixture; Verify-Gates Waiting 'cached success during new run'
    $heartbeat.RunningJobs=@(); Write-Fixture
    $parent.Enabled=$false
    $ignored=& $gate {param($c,$p,$n) Get-OrchestratorExternalDependencyStatus -Job $c -DependencyName $p.Name -Now $n} $consumer $parent $now
    Check ($ignored -eq 'Ready') 'Disabled external dependency behavior changed.'
    "Pinned dependency offline checks passed: $script:checks"
}
finally {
    $global:SmartM365ProfileKey=$oldProfile
    if([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($root)).TrimEnd('\') -eq [IO.Path]::GetTempPath().TrimEnd('\') -and [IO.Path]::GetFileName($root) -like 'SmartM365-PinnedDeps-*') {
        Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDst5sxEkt0YpVp
# 43xTGM5NHDwGOp2Kqlhrr8U4iO1uX6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJjXyrR0rAv6CShrXD3HLbJbvKfdcOfygnNHaECITGHjMA0GCSqG
# SIb3DQEBAQUABIIBgKuMr6vTIxG5q1ZFte+BFSnkDaW2fnREr3OkMn2kOkjJ+IhU
# IBQgM+WKxlgAYwSipXMgULOBfEA3DoYJmcbwTpsyhOwIkrGtA+df8BlXep8oFX6u
# 7Lrv+q4SXy/XycxxIlTfRE34+t6lhn2Edd758XDcMSFc/NGYkc4PwREaMKevaBqY
# Y4/mc/FhOuXyXQbh9mxAKNqKvGIQoYak4rxlGb7fU8VVAk+7hJdamNY4CRmwRo8u
# o23+YS8uSrdx4IGiljOfY0WnAqbmOveYjIRRwrXzfxZd261SULjdSFabETL/jFsJ
# tNyjHbNoUqGQj4miVYZhRYRvxjbTlvmEc18yOx4Sqtn1Yx9OQji9ijlPBMJIadvg
# ZrzKDu9iX+inP4fDLPLRCjBute6pQMcVI14V07TXi0iq5GLJbOgoTmr6xHq6r4Yj
# ksybrplRXvmoFS2j0rPnTdDhyihnlop8dx9rAqgXLObTFAT2zKn8dMAQVFb/L4e0
# lcoahRbJzoJHdOSwSqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcyMTEx
# MDRaMC8GCSqGSIb3DQEJBDEiBCDflhx0iirhS4qfzbTsiN1QhTkaXrEAbSF2/lR6
# yQKhDTANBgkqhkiG9w0BAQEFAASCAgCL27ozHFQ9g5lVUYocfXyKbDbuzRS8yziI
# 0sXCxKiNwQGtRdojEZuj32jisS202bUdbVU7BPFOI/WHINLhnIsKGpycZLE75DFg
# qVmj60BERF9hg8jgUF+mK7FXvdFNPR0atj4VZLxwyoJmmaQhFP0rrj8jl40MDnbl
# LosF4M7AS8R8T2D0dCUmwcJL2Xc9aC4WzVqkaK78ON3BYRsXuh10f/nbWo/QDtAJ
# VKhxKPuypdT6/8+A0eOSoVIlBKdd+8EKmmXK7856fjdjuvMzw5hjwuEB61qA0h15
# X6ivsn93mTt/qX5c5vDl7M8/FS+f+rScHH6/cZYdpJVtrB39amclVt2DJ48m3yaX
# wxuz/c6zdzBN2IkZ8kxY9ms+JjUvPfzGFp4l7vIVti3/RddxKPZ0bqs38XlswabG
# 0kEnHHFaSD2RVl0ck94RbKgSNgTrc796PJIx/gqsX29lPpu2bo35OibNQ9hrdlCm
# u7lmeHh4SFjQTXM2ohJfKWEDjICmzszgJRUIuvghahMpXsvxECnhosV6LX3tRzyt
# cA/ZZNalvP3r6NLImPXYZAUkU6Mv0siAZ+TgR2wPzvEPkVI7Qbd8y2mj4BhvRuuh
# f8P89L1HeeyJfEWTUHZmrzFVWE+rsSlQZUkEfGKRyR6H7Xrjrk7vmohkV/7KTZh0
# V6XIjNQIUQ==
# SIG # End signature block
