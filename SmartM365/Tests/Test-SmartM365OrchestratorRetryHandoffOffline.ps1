<#
.SYNOPSIS
Offline test of the retry handoff: an elected job that failed on SRV-A is retried by a free,
eligible server, and exactly one server runs the retry.
.VERSION
1.0.1
.NOTES
Real Distributed module claims in a temporary folder; the orchestrator offer and adoption
functions run in one module instance per simulated server. No tenant initialization,
collectors, scheduling or publication.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
$transport = Get-Module SmartM365.JsonTransport
& $transport { function script:Get-SmartM365JsonTransportPolicy { @{ Mode = 'Readers'; QualifiedUncRoots = @(); QualifiedLocalDataRoots = @() } } }

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
function Get-FunctionText([string]$Name) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing orchestrator function: $Name" }
    $definition.Extent.Text
}

# Wiring: offered only for elected jobs once the claim is RetryScheduled; the failing server waits.
$complete = Get-FunctionText 'Complete-JobRun'
Check ($complete -match "claimStatus -eq 'RetryScheduled'.*AssignmentMode -eq 'Elected'" -and $complete -match 'Write-OrchestratorRetryOffer') 'Only an elected retry with a RetryScheduled claim may be offered.'
$launch = Get-FunctionText 'Invoke-LaunchPhase'
Check ($launch -match 'Invoke-RetryHandoffAdoption -Now \$Now' -and $launch -match '\$Now -lt \$notBefore\.AddSeconds\(\$script:RetryHandoffGraceSeconds\)') 'Peers must adopt first and the failing server must wait its grace period.'
Check ($launch -match '-RetryHandoffFrom \$retryHandoffFrom' -and $launch -match 'Remove-OrchestratorRetryOffer -JobName \$name -OfferId \$retryOfferId') 'The launch must take the claim over atomically and remove the offer.'

$functions = @('ConvertTo-OrchestratorUtcTime', 'Get-OrchestratorRetryOfferPath', 'Write-OrchestratorRetryOffer', 'Get-OrchestratorRetryOffers', 'Remove-OrchestratorRetryOffer', 'Test-OrchestratorRetryHandoffEligible', 'Invoke-RetryHandoffAdoption', 'Test-OrchestratorLocalServerActive', 'Test-JobAllowedOnServer', 'ConvertFrom-StateTime', 'ConvertTo-StateTime')
$stubs = @'
function Get-JobState { param($JobName) if (-not $script:State.Jobs.ContainsKey($JobName)) { $script:State.Jobs[$JobName] = @{ PendingRetry = $null } }; $script:State.Jobs[$JobName] }
function Save-OrchestratorState { }
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Logs.Add($Message) }
function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now, $Level) $script:Logs.Add($Message) }
function Test-JobSelected { param($JobName) $true }
function Get-ElectedJobOwner { param($JobName) 'SRV-A' }
'@
$functionText = (@($functions | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-RetryHandoff-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
$claims = Join-Path $root 'Claims'
New-Item -ItemType Directory -Path $claims -Force | Out-Null
$originalComputer = $env:COMPUTERNAME
$job = [pscustomobject]@{ Name = 'Graph-Job'; AssignmentMode = 'Elected'; RequiredCapabilities = @('Graph'); RequiredGraphAppRoles = @(); TimeoutMinutes = 60 }
function New-Server([string]$Name, [string[]]$Ready, [bool]$Maintenance = $false) {
    $module = New-Module -Name ('Server-' + $Name) -ScriptBlock ([scriptblock]::Create($functionText))
    & $module {
        param($Root, $Claims, $Job, $Ready, $Maintenance)
        $script:Logs = [Collections.Generic.List[string]]::new()
        $script:State = @{ Jobs = @{} }
        $script:RunningJobs = @{}
        $script:Settings = [pscustomobject]@{ SharedDataFolderPath = $Root; ElectionClaimsPath = $Claims; DistributedSchedulingEnabled = $true; MaxConcurrency = 2; ServerJobPolicies = @{} }
        $script:ElectionPlan = [pscustomobject]@{ PlanId = 'plan-test' }
        $script:ServerCapabilities = [pscustomobject]@{ ReadyCapabilities = $Ready; GraphAppRoles = @() }
        $script:Manifest = [pscustomobject]@{ JobsByName = @{ $Job.Name = $Job } }
        $script:MaintenanceControl = [pscustomobject]@{ Enabled = $Maintenance }
    } $root $claims $job $Ready $Maintenance
    $module
}
function Invoke-As([string]$Server, $Module, [scriptblock]$Block, [object[]]$Arguments = @()) {
    $env:COMPUTERNAME = $Server
    try { & $Module $Block @Arguments } finally { $env:COMPUTERNAME = $originalComputer }
}
function New-FailedOccurrence($ServerA, [datetime]$Occurrence, [datetime]$NotBefore, [string]$BatchId = 'BATCH-1') {
    # SRV-A ran the job, it failed with a retry left: claim RetryScheduled, local PendingRetry, offer.
    Invoke-As 'SRV-A' $ServerA {
        param($Occurrence, $NotBefore, $BatchId, $Claims)
        $claim = Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $Claims -JobName 'Graph-Job' -Occurrence $Occurrence -OwnerServer 'SRV-A' -PlanId 'plan-test'
        $null = Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $claim.ClaimPath -OwnerServer 'SRV-A' -Status RetryScheduled -Attempt 0
        $pending = @{ NotBefore = ConvertTo-StateTime -Value $NotBefore; Attempt = 1; ScheduledOccurrence = ConvertTo-StateTime -Value $Occurrence; PipelineBatchId = $BatchId; PipelineStatusPath = '' }
        $pending['HandoffOfferId'] = Write-OrchestratorRetryOffer -JobName 'Graph-Job' -PendingRetry $pending -ClaimPath $claim.ClaimPath
        $pending['HandoffOffered'] = $true
        $script:State.Jobs['Graph-Job'] = @{ PendingRetry = $pending }
        [pscustomobject]@{ ClaimPath = $claim.ClaimPath; OfferId = $pending.HandoffOfferId }
    } @($Occurrence, $NotBefore, $BatchId, $claims)
}
function Enter-Claim($Module, [string]$Server, [datetime]$Occurrence, [string]$HandoffFrom = '') {
    Invoke-As $Server $Module { param($Occurrence, $Server, $HandoffFrom, $Claims) Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $Claims -JobName 'Graph-Job' -Occurrence $Occurrence -OwnerServer $Server -PlanId 'plan-test' -RetryHandoffFrom $HandoffFrom } @($Occurrence, $Server, $HandoffFrom, $claims)
}
try {
    $serverA = New-Server 'SRV-A' @('Graph')
    $serverB = New-Server 'SRV-B' @('Graph')
    $serverC = New-Server 'SRV-C' @('AD')
    $serverD = New-Server 'SRV-D' @('Graph')
    $now = Get-Date
    $occurrence = $now.Date.AddHours(5).AddMinutes(10).ToUniversalTime()

    # Failure on SRV-A; the retry delay is still running: nobody adopts it yet.
    $failed = New-FailedOccurrence $serverA $occurrence $now.AddMinutes(5)
    Invoke-As 'SRV-B' $serverB { param($n) Invoke-RetryHandoffAdoption -Now $n } @($now)
    Check ($null -eq (& $serverB { (Get-JobState -JobName 'Graph-Job').PendingRetry })) 'A retry must not be adopted before RetryDelaySeconds.'

    # After the delay: SRV-B (free, eligible) adopts; SRV-C lacks Graph; SRV-A ignores its own offer.
    $later = $now.AddMinutes(6)
    foreach ($pair in @(@('SRV-B', $serverB), @('SRV-C', $serverC), @('SRV-D', $serverD), @('SRV-A', $serverA))) { Invoke-As $pair[0] $pair[1] { param($n) Invoke-RetryHandoffAdoption -Now $n } @($later) }
    $adoptedB = & $serverB { (Get-JobState -JobName 'Graph-Job').PendingRetry }
    Check ($null -ne $adoptedB -and $adoptedB.HandoffFrom -eq 'SRV-A' -and $adoptedB.Attempt -eq 1 -and $adoptedB.PipelineBatchId -eq 'BATCH-1' -and $adoptedB.HandoffOfferId -eq $failed.OfferId) 'SRV-B must adopt the retry with its attempt and pipeline batch.'
    Check ($null -eq (& $serverC { (Get-JobState -JobName 'Graph-Job').PendingRetry })) 'A server without the required capability must not adopt the retry.'
    Check ((Invoke-As 'SRV-B' $serverB { Test-JobAllowedOnServer -Job $script:Manifest.JobsByName['Graph-Job'] }) -and -not (Invoke-As 'SRV-C' $serverC { Test-JobAllowedOnServer -Job $script:Manifest.JobsByName['Graph-Job'] })) 'Only the adopting server may launch the elected job it does not own.'

    # Race: SRV-B and SRV-D both adopted; the atomic claim lets exactly one server run the retry.
    $claimB = Enter-Claim $serverB 'SRV-B' $occurrence 'SRV-A'
    $claimD = Enter-Claim $serverD 'SRV-D' $occurrence 'SRV-A'
    $claimA = Enter-Claim $serverA 'SRV-A' $occurrence
    $winners = @(@($claimB, $claimD, $claimA) | Where-Object Acquired)
    $stored = (Get-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName 'Graph-Job' -Occurrence $occurrence).Claim
    Check ($winners.Count -eq 1 -and $claimB.Acquired -and $stored.OwnerServer -eq 'SRV-B' -and $stored.Status -eq 'Claimed' -and $stored.HandoffFrom -eq 'SRV-A') "Exactly one server must win the retry (B=$($claimB.Acquired), D=$($claimD.Acquired), A=$($claimA.Acquired))."
    # The losers drop their copy of the retry when their claim is refused (Invoke-LaunchPhase).
    & $serverD { $script:State.Jobs['Graph-Job'].PendingRetry = $null }

    # The launcher removes the offer; peers then ignore the occurrence.
    Invoke-As 'SRV-B' $serverB { param($id) Remove-OrchestratorRetryOffer -JobName 'Graph-Job' -OfferId $id } @($failed.OfferId)
    Check (@(Invoke-As 'SRV-D' $serverD { Get-OrchestratorRetryOffers }).Count -eq 0) 'The offer must be removed once the retry is launched.'

    # No free eligible peer: SRV-A takes its own retry back atomically, and a late peer is refused.
    $occurrence2 = $occurrence.AddDays(-1)
    $null = New-FailedOccurrence $serverA $occurrence2 $now.AddMinutes(-3) ''
    $ownA = Enter-Claim $serverA 'SRV-A' $occurrence2
    $lateB = Enter-Claim $serverB 'SRV-B' $occurrence2 'SRV-A'
    $stored2 = (Get-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName 'Graph-Job' -Occurrence $occurrence2).Claim
    Check ($ownA.Acquired -and -not $lateB.Acquired -and $stored2.OwnerServer -eq 'SRV-A' -and $stored2.Status -eq 'Claimed') 'Without a peer, the failing server must keep its retry and a late peer must be refused.'

    # The offer of that occurrence is now stale (claim no longer RetryScheduled): any server removes it.
    Invoke-As 'SRV-D' $serverD { param($n) Invoke-RetryHandoffAdoption -Now $n } @($later)
    Check (@(Invoke-As 'SRV-D' $serverD { Get-OrchestratorRetryOffers }).Count -eq 0 -and $null -eq (& $serverD { (Get-JobState -JobName 'Graph-Job').PendingRetry })) 'A stale offer must be cleaned and never adopted.'

    # Maintenance: a scheduled (non-pipeline) retry is not adopted while maintenance is on.
    $serverM = New-Server 'SRV-M' @('Graph') $true
    $occurrence3 = $occurrence.AddDays(-2)
    $null = New-FailedOccurrence $serverA $occurrence3 $now.AddMinutes(-3) ''
    Invoke-As 'SRV-M' $serverM { param($n) Invoke-RetryHandoffAdoption -Now $n } @($later)
    Check ($null -eq (& $serverM { (Get-JobState -JobName 'Graph-Job').PendingRetry })) 'Maintenance must keep a scheduled retry from being adopted.'
}
finally {
    $env:COMPUTERNAME = $originalComputer
    Get-Module -Name 'Server-*' | Remove-Module -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: retry handoff, $($script:passed) offline checks."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBARaWIxjitsnw7
# EGheWx9JOz6JxjApY9UvpO+7oDL9vaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKvbsN0Oel8Nurl0Mq6TaqxlU0ztk/7Kx1NRaz4LNGxuMA0GCSqG
# SIb3DQEBAQUABIIBgH079QuIZWcQqENhj/DfTmpx8M7JziEBZWXIjdY7hPS+n8a+
# AJx/wKMt2J9OOCo78IedD28bva9c3RrYYaJs/F7/qdeW2x996sNX45ownqQpkocv
# hpzfOmwtQnqAArBLu1lmit6ujvV56ipYdCcQ2/jCBVfoiLzlNR0bNVpZvNKDKyWh
# xSuzEIPWFXDksFm23OHxOhKwyv5B9YIxWjij8XWkjYezIWbwNqiem1WtnOaI/p7A
# LiUyJsNQTuOywnr6C8iBD/TsYpfvbm/BDPZ5TtjuB3733OyO4zK3EeOK4HOufrvQ
# /u6AKQtJGX3fCcCz9Cir/j68lPuVP5+naLh7nlNQGLVPnAbj4jOWPsLE6tfHUsif
# Smy+yBPF+5UbMCCeORBtMP/i0/fqlgEwp6qbE3y1NsUL51Xg/e0ly3jlbTCHU5/l
# BUH84J52hFUpsCPjuup4htutO4w5JT9YDLtaDPIVAhkp8DWGXi+NScvqVZwpI+dS
# /Is1wV62iP/Uf6Ju6KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkyMjAw
# MjFaMC8GCSqGSIb3DQEJBDEiBCANTrkmY4aYQCtSOIjTPZePJr3A1USlYPqRgRXB
# NG5VeTANBgkqhkiG9w0BAQEFAASCAgCWhDFbSiMFgsY4Km/LmJnrS0pef1LkLlDH
# yi9wgBjldLhuiH20Sy/HV17U+ImxBVB/fVVANSp/r711tN+IN2P/EIPWL8F7YrQe
# fCyrCfgmigfjzmFfk/qGEduQkLp/H3rxbYFXXtUHgRnlwkQ9Kyq27CkvREVeSFwB
# p473+jgZGdplqCqNdqwYyDICKlyUVNtyL9NxCm5L0VWHFOYEVt79BvDpho933Zq8
# x/pZARoSVq8A876k8jxkxfe/hBZp/7gua9mRlj3CDHtw4ZrU0RDh+gAGJHvm5IQx
# 1H/iO75d52fQ+mQsKYZA/9RGyA4GpScFvgTJdF+0Q6Va/Debde8RlsS+FvIhKdjo
# D7lxmcj2A9ADAH6CyEY3CoF/tEyC/F5bosHJ0UDH6YL97y8dT3CjDWPy9N+AYAKU
# NW2pJ9olzNmRI71dv7gNl1UktL4WMzP7ueTAQK3WBxvottfVDCz6KDpsUkOojuOX
# yxhA0MaNzUNHdptLQCv3Z9NOnazpMT8J3DMmt3hDU77K+Ra424jfVkfrP5SauQou
# fOwRrvqanN2gtqCAywvB8umPL2YQQU2e1yC2mhNHT6NkiZMtBpzpeY/0EcyKp8so
# 90Qf7v+cA0BbnjNTYQrGGeducfLHxQC2CHFoDR9YQ9FT01zpE6G6mVsuByPGE6L9
# /NjePGrWTw==
# SIG # End signature block
