<#
.SYNOPSIS
Offline tests of the wait reason of a pipeline request job that waits for its dependencies: the
reason names each blocking dependency with its server and start time, is written to the request
job status only when it changes, and is published in the heartbeat pending jobs.
.VERSION
1.0
.NOTES
Synthetic jobs, claims and request status files in a temporary folder. No tenant initialization,
collectors, scheduling, authentication or publication.
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

# The pipeline gate records the wait and forgets it when the job proceeds or is blocked.
$launchPhase = Get-FunctionText 'Invoke-LaunchPhase'
Check ($launchPhase -match 'Set-OrchestratorPipelineDependencyWait -JobName \$name -PipelineInfo \$pipelineInfo') 'The pipeline gate no longer records the dependency wait.'
Check (([regex]::Matches($launchPhase, '\$script:PipelineWaitState\.Remove\(\$name\)')).Count -ge 3) 'The pipeline wait is not cleared on launch, failed dependency and timeout.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-PipelineWait-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root | Out-Null
try {
    $functions = @('ConvertTo-StateTime', 'Write-OrchestratorRuntimeUpdateWarning', 'Get-OrchestratorExternalDependencyStatus', 'Get-OrchestratorDependencyActivityText',
        'Get-OrchestratorPipelineDependencyWaitText', 'Set-OrchestratorPipelineDependencyWait', 'Get-OrchestratorPipelineDependencyStatus', 'Get-OrchestratorPendingJobSnapshot')
    $stubs = @'
function Write-OrchestratorLog { param([string]$Message, [string]$Level = 'INFO') }
function Get-OrchestratorEffectiveDependencyMode { param($Job, $DependencyName) $script:Mode }
function Get-OrchestratorDependencyMaxAge { param($DependencyJob, $OverrideHours, $Now) 24 }
function Get-OrchestratorDependencyActivity { param($DependencyJob, $DownstreamOccurrence, $Now) $script:Activity }
function Get-OrchestratorDependencyFreshStatus { param($DependencyJob, $MaxAgeHours, $Now) $script:Fresh }
function Get-OrchestratorSharedDependencyStatus { param($Job, $Now) 'Waiting' }
function Test-JobAllowedOnServer { param($Job, [switch]$AllowManual) $true }
function Get-JobState { param($JobName) [pscustomobject]@{ PendingRetry = $null; LastScheduledOccurrence = '' } }
function Get-DueOccurrence { param($Job, $LastOccurrence, $Now) $null }
function Get-SmartM365OrchestratorPipelineJobStatusPath { param($SharedDataFolderPath, $BatchId, $JobName) Join-Path $SharedDataFolderPath "$BatchId-$JobName.json" }
function Set-OrchestratorPipelineJobStatus { param($BatchId, $JobName, $Status, [int]$Attempt = 0, $Detail = '') $script:StatusWrites.Add([pscustomobject]@{ BatchId = $BatchId; JobName = $JobName; Status = $Status; Detail = $Detail }) }
'@
    $text = (@($functions | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs
    $module = New-Module -ScriptBlock ([scriptblock]::Create($text))
    & $module {
        param($Root)
        $script:Settings = [pscustomobject]@{ SharedDataFolderPath = $Root; DistributedSchedulingEnabled = $true; MaxConcurrency = 4 }
        $script:StatusWrites = [Collections.Generic.List[object]]::new()
        $script:RuntimeUpdateWarnings = @{}
        $script:DependencyWaitReason = @{}
        $script:PipelineWaitState = @{}
        $script:DependencyWaitLogState = @{}
        $script:RunningJobs = @{}
        $script:MaintenanceHealthy = $true
        $script:MaintenanceControl = [pscustomobject]@{ Enabled = $false }
        $scope = [pscustomobject]@{ Name = 'Dep-Scope'; Enabled = $true; AssignmentMode = 'Elected'; ContinueOnError = $false }
        $inventory = [pscustomobject]@{ Name = 'Dep-Inventory'; Enabled = $true; AssignmentMode = 'Elected'; ContinueOnError = $false }
        $script:PrepareJob = [pscustomobject]@{ Name = 'Prepare'; Enabled = $true; AssignmentMode = 'Elected'; DependencyMaxAgeHours = 0; ConcurrencyKey = 'Prepare'; DependsOn = @('Dep-Scope', 'Dep-Inventory') }
        $script:Manifest = [pscustomobject]@{ JobsByName = @{ 'Dep-Scope' = $scope; 'Dep-Inventory' = $inventory; 'Prepare' = $script:PrepareJob }; OrderedJobs = @($script:PrepareJob) }
    } $root

    $now = [datetime]'2026-10-28 21:00'
    $claimedAt = '2026-10-28T17:16:00.0000000Z'
    $claimedLocal = ([datetimeoffset]::Parse($claimedAt)).LocalDateTime.ToString('yyyy-MM-dd HH:mm')

    # Dependency outside the request, running on another server: server and start time.
    $status = & $module {
        param($n, $c)
        $script:Mode = 'FreshSuccess'
        $script:Activity = [pscustomobject]@{ State = 'Waiting'; Detail = 'Running for occurrence 2026-10-28 18:15, may rewrite DATA-LAST during the run'; OwnerServer = 'SRV-B'; ClaimedAtUtc = $c; Status = 'Running' }
        Get-OrchestratorExternalDependencyStatus -Job $script:PrepareJob -DependencyName 'Dep-Scope' -Now $n
    } $now $claimedAt
    Check ($status -eq 'Waiting') 'A running external dependency must keep the request job waiting.'
    $reason = & $module { $script:DependencyWaitReason['Dep-Scope'] }
    Check ($reason -eq "Running on SRV-B since $claimedLocal") "The external reason must name the server and the start time: $reason"

    # Not running but not fresh: the freshness rule is named.
    & $module { param($n) $script:Activity = [pscustomobject]@{ State = 'Clear'; Detail = 'No active execution' }; $script:Fresh = 'Waiting'; [void](Get-OrchestratorExternalDependencyStatus -Job $script:PrepareJob -DependencyName 'Dep-Scope' -Now $n) } $now
    Check ((& $module { $script:DependencyWaitReason['Dep-Scope'] }) -eq 'no successful run in the last 24 h') 'A stale external dependency must name the freshness rule.'
    & $module { param($n) $script:Fresh = 'Ready'; [void](Get-OrchestratorExternalDependencyStatus -Job $script:PrepareJob -DependencyName 'Dep-Scope' -Now $n) } $now
    Check (-not (& $module { $script:DependencyWaitReason.ContainsKey('Dep-Scope') })) 'A ready dependency must not keep an old wait reason.'

    # Dependency of the same request: its request status and server.
    [IO.File]::WriteAllText((Join-Path $root 'B1-Dep-Inventory.json'), '{"Status":"Running","OwnerServer":"SRV-C"}')
    Check ((& $module { Get-OrchestratorPipelineDependencyStatus -BatchId 'B1' -JobName 'Dep-Inventory' }) -eq 'Running') 'The request dependency status must still be returned.'
    Check ((& $module { $script:DependencyWaitReason['Dep-Inventory'] }) -eq 'in request: Running on SRV-C') 'The request dependency reason must name its status and server.'

    # Wait text: each blocking dependency with its reason; unknown ones stay readable.
    & $module { param($n, $c) $script:Activity = [pscustomobject]@{ State = 'Waiting'; Detail = 'x'; OwnerServer = 'SRV-B'; ClaimedAtUtc = $c; Status = 'Running' }; [void](Get-OrchestratorExternalDependencyStatus -Job $script:PrepareJob -DependencyName 'Dep-Scope' -Now $n) } $now $claimedAt
    $waitText = & $module { Get-OrchestratorPipelineDependencyWaitText -BlockingDependencies @('Dep-Scope (outside request)', 'Dep-Inventory') }
    Check ($waitText -eq "Waiting for 2 dependencies: Dep-Scope (outside request: Running on SRV-B since $claimedLocal); Dep-Inventory (in request: Running on SRV-C).") "Unexpected wait text: $waitText"
    $single = & $module { Get-OrchestratorPipelineDependencyWaitText -BlockingDependencies @('Unknown-Job') }
    Check ($single -eq 'Waiting for 1 dependency: Unknown-Job (not finished).') "Unexpected single wait text: $single"

    # Request job status: written as Pending with the reason, once per change; a retry keeps its status.
    $info = @{ BatchId = 'B1' }
    & $module { param($i, $t, $n) Set-OrchestratorPipelineDependencyWait -JobName 'Prepare' -PipelineInfo $i -Text $t -Now $n; Set-OrchestratorPipelineDependencyWait -JobName 'Prepare' -PipelineInfo $i -Text $t -Now $n.AddMinutes(1) } $info $waitText $now
    $writes = & $module { @($script:StatusWrites) }
    Check ($writes.Count -eq 1 -and $writes[0].Status -eq 'Pending' -and $writes[0].Detail -eq $waitText -and $writes[0].BatchId -eq 'B1') 'The request job must receive the wait reason once.'
    & $module { param($i, $n) Set-OrchestratorPipelineDependencyWait -JobName 'Prepare' -PipelineInfo $i -Text 'Waiting for 1 dependency: Dep-Scope (x).' -Now $n.AddMinutes(5) } $info $now
    $state = & $module { $script:PipelineWaitState['Prepare'] }
    Check ((& $module { $script:StatusWrites.Count }) -eq 2 -and $state.FirstSeen -eq $now) 'A changed reason must be rewritten and keep the first wait time.'
    & $module { param($i, $n) Set-OrchestratorPipelineDependencyWait -JobName 'Prepare' -PipelineInfo $i -Text 'retry text' -IsRetry -Now $n } $info $now
    Check ((& $module { $script:StatusWrites.Count }) -eq 2 -and (& $module { $script:PipelineWaitState['Prepare'].Text }) -eq 'retry text') 'A scheduled retry must not be rewritten as Pending.'

    # Heartbeat: the waiting request job is listed with its request and reason.
    & $module { param($t) $script:PipelineWaitState['Prepare'].Text = $t; $script:PipelinePending = @{ 'Prepare' = @{ BatchId = 'B1' } } } $waitText
    $rows = @(& $module { param($n) Get-OrchestratorPendingJobSnapshot -Now $n } $now)
    Check ($rows.Count -eq 1 -and $rows[0].Name -eq 'Prepare' -and $rows[0].Reason -eq 'WaitingDependencies' -and $rows[0].Details -eq "Request B1: $waitText" -and $rows[0].FirstSeen -eq $now.ToString('o')) 'The heartbeat must list the waiting request job with its reason.'
    & $module { $script:MaintenanceControl = [pscustomobject]@{ Enabled = $true; Reason = 'Patching' } }
    $rows = @(& $module { param($n) Get-OrchestratorPendingJobSnapshot -Now $n } $now)
    Check ($rows.Count -eq 1 -and $rows[0].Reason -eq 'WaitingDependencies') 'Maintenance does not hide a request job that waits for its dependencies.'
    & $module { $script:PipelinePending = @{} }
    $rows = @(& $module { param($n) Get-OrchestratorPendingJobSnapshot -Now $n } $now)
    Check (@($rows | Where-Object Reason -eq 'WaitingDependencies').Count -eq 0) 'A finished or cancelled request must not stay in the heartbeat.'

    # Distributed module: the active claim gives its owner and claim time.
    Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
    $claims = Join-Path $root 'Claims'
    $occurrenceUtc = [datetime]::UtcNow.AddHours(-1)
    $folder = Join-Path $claims 'Dep-Scope'
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $claim = [ordered]@{ SchemaVersion = 1; JobName = 'Dep-Scope'; OccurrenceUtc = $occurrenceUtc.ToString('o'); OwnerServer = 'SRV-B'; Status = 'Running'; CreatedAtUtc = $claimedAt; UpdatedAtUtc = $claimedAt; SafeUntilUtc = [datetime]::UtcNow.AddHours(2).ToString('o') }
    [IO.File]::WriteAllText((Join-Path $folder ($occurrenceUtc.ToString('yyyyMMddTHHmmssfffZ') + '.json')), ($claim | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    $activity = Get-SmartM365OrchestratorDependencyActivity -ClaimsRootPath $claims -DependencyJob ([pscustomobject]@{ Name = 'Dep-Scope'; EstimatedDurationMinutes = 240 }) -DownstreamOccurrence (Get-Date) -Now (Get-Date)
    Check ($activity.State -eq 'Waiting' -and $activity.OwnerServer -eq 'SRV-B' -and ([datetimeoffset]::Parse($activity.ClaimedAtUtc, [Globalization.CultureInfo]::InvariantCulture)) -eq [datetimeoffset]::Parse($claimedAt, [Globalization.CultureInfo]::InvariantCulture) -and $activity.Status -eq 'Running') 'The dependency activity must return the owner and claim time of the active claim.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: pipeline wait reason, $($script:passed) offline checks."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD4G5ABeb08Cezk
# otV3AQKBryJ0U+vWMfDVl/hDeuxOwKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID2vL+gSuGP0UYMgJoogAiTbPZL41Xp1GG3Ebxe1DA9oMA0GCSqG
# SIb3DQEBAQUABIIBgG2YrE9od1256nTYco+5aF+MeO+pYRwo4rUjTJRGAK54QzeT
# YWSimSwBzgrho/csZzeoIGY1BLJUy/mRN3rO6PQ9G6WbAQ9mzXhtBJKf6GULGS9x
# yEMwYHwnM0RRYjtu8cPC3waZfApCzmCT35ijOWTfdvBBY51PNvpcuwbevCgEIfkM
# PNwl2rzYtHdfPxjO0YfTpEhUjhPhOkh5RB6L0Wr4odJmZbfGHehq64R2zGElElsy
# GxxYkKtuvMX1xh3h8uZ02YF9gC//DEfcKshmgUUYp7m8JphYzl4DdfW0R4+XvHsr
# mbuAu1tFoXiJl5O95w9CQhzYehpi+UVIGp493GCtygcfsYLi+Hi6Bs03Q7tmpObe
# PuRagUBIuHpodlkcXTXyhs3u/X5VMZXUqx4EudAcTjTZLrblUOZUS6c4+EGCo80F
# kzw2PZAyck/UhXo1oS+rq3tztGBlbJX/fdJJ8VnVQB9aWpGxgON6kY8ovbdzDNXB
# o0KuR17Flv7aRZC9kqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgyMDQz
# MzlaMC8GCSqGSIb3DQEJBDEiBCBHc4KEHnlbQXt/U5RbuF5u3VHGHr0nsVwlNMlo
# 4onO7TANBgkqhkiG9w0BAQEFAASCAgB3szGeeVRIFLfq182JZLyMCi6Qqrz8+0Lr
# zZyMpu2EZoR0mhriJr4tGzV5XwnOltWqKh2KcHfcMYp0ioVclqx9sHIBODbBMmK6
# msEX3UFLg0m8PIcXynfdLH5S0Ulf7PaOo9UVprRiGnXJkgwUbbdj6sFo1nnb4uvB
# ObEI+An/GyDVFx4FNunxKE7vAl1qZFbU4kezBTFs9pDzidHECGYYUMaThTmKOlbY
# aIT0JWQ53vN5oo5dp5CHuC1CXOR9vRRcORdR+F06lBrslSegRi/XKM0FTRJ9OSi9
# lUXLBDpUt/olr8vvdbNQ/DxC1PV6rYghNk7pI+Yy3IDAnN5dRxZi8PKVF6pyhWS6
# BZ6lSIg+lj3ASX9njHO+NiL4QetF8CH+InJjWGdvgQY8WFXK+hbCayoIiDxs9Ja0
# GO05KN6OgdLQGOWgKyNIe4OsG5XRhVQmyUTgeHV8rgdco0rWJiJsWtEIq/7w6JCD
# gYPjUyxHIBb5xrVF7sRaBEEZ+jl1tis3VtS6RfRUBD3sG6/VoC2tP09xeWup86hC
# IzFWYaqQgCHBsv8U5iBjYmpelviH8HJMeD9Z68CTq59Gdc7FFvZCl4wF3+5Ga/1U
# UNqEf5RDs7AGc4TX4SJLsZbNbbDmZm2jcQOtKdiVGi26T35UiHFpb26K+LQtJxK5
# pk8s09zWqA==
# SIG # End signature block
