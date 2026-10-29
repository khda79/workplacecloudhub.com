<#
.SYNOPSIS
Synthetic licensing dependency policy checks. No scheduler entry point or collector is executed.
.VERSION
1.0.3
#>
#requires -Version 7.0
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic tenant key consumed by the production receipt reader; restored in finally.')]
param()
$ErrorActionPreference='Stop'
$base=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Orchestrator parse failed.'}
$names=@('Get-OrchestratorEffectiveDependencyMode','Get-OrchestratorExternalDependencyStatus','Get-OrchestratorSharedDependencyStatus','Get-LatestPastOccurrence','Get-JobOccurrencesInWindow','ConvertTo-NormalizedJob','Get-OrchestratorDependencyActivityText')
$defs=foreach($name in $names){$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true).Extent.Text}
# Real latest-occurrence claim reader; the independent freshness gate is a stand-in.
$defs+=@'
function Get-OrchestratorLaterDependencyClaimStatus { return $null }
function Test-OrchestratorOccurrenceCoveredByOverlap { return $false }
function Get-OrchestratorDependencyMaxAge { return 48 }
function Get-OrchestratorDependencyActivity { [pscustomobject]@{State='Clear'} }
function Get-OrchestratorDependencyFreshStatus { $script:FreshCalls++; return 'Ready' }
function Write-OrchestratorRuntimeUpdateWarning { param($Key,$Message,$Now) throw $Message }
$script:DependencyWaitReason = @{}
'@
$gate=New-Module -ScriptBlock ([scriptblock]::Create($defs -join "`n"))
$root=Join-Path $env:TEMP ('SmartM365-LicenseDeps-'+[guid]::NewGuid().ToString('N'))
$shared=Join-Path $root 'DATA-ALL/Orchestrator'
$claims=Join-Path $shared 'Election/Claims'
$checks=0
$previousTenant=Get-Variable SmartM365TenantKey -Scope Global -ValueOnly -ErrorAction SilentlyContinue
$global:SmartM365TenantKey='synthetic'
function Check($Value,$Message){if(-not $Value){throw $Message};$script:checks++}
$versionDefinitions=foreach($name in @('Get-OrchestratorFileFingerprint','Get-OrchestratorRuntimeSnapshot','Test-OrchestratorRuntimeUpdate')){
 $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true).Extent.Text
}
# Exercise the real disk-version detector, without running the resident process or sending mail.
$versionDefinitions+=@'
function Write-OrchestratorLog { param($Message,$Level) }
function Write-OrchestratorRuntimeUpdateWarning { param($Key,$Message,$Now) throw $Message }
function Test-OrchestratorAuthenticodeFile { param($Path,$Role) [pscustomobject]@{IsValid=$true} }
'@
$versionGate=New-Module -ScriptBlock ([scriptblock]::Create($versionDefinitions -join "`n"))
$document=Get-Content -LiteralPath (Join-Path $base 'Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
$licences=@($document.Jobs | Where-Object Name -eq 'M365-Licences-Inventory')[0]
$prepare=@($document.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')[0]
$usage=@($document.Jobs | Where-Object Name -eq 'M365-Usage-Inventory')[0]
$devices=@($document.Jobs | Where-Object Name -eq 'Intune-Devices-Inventory')[0]
$now=[datetime]::new(2026,10,27,9,0,0,[DateTimeKind]::Local)
function Write-Claim($Job,[datetime]$When,[string]$Status){
 $folder=Join-Path $claims $Job.Name;New-Item -ItemType Directory -Path $folder -Force|Out-Null
 $claim=@{SchemaVersion=1;ClaimId='synthetic';JobName=$Job.Name;OccurrenceUtc=$When.ToUniversalTime().ToString('o');Status=$Status;OwnerServer='Synthetic';PlanId='synthetic';OrchestratorPid=1;Attempt=0;CreatedAtUtc=$When.ToUniversalTime().ToString('o');UpdatedAtUtc=$When.ToUniversalTime().AddMinutes(1).ToString('o');SafeUntilUtc=$When.AddDays(2).ToUniversalTime().ToString('o')}
 [IO.File]::WriteAllText((Join-Path $folder ($When.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')),($claim|ConvertTo-Json))
}
function Status($Consumer,$Parent){& $gate {param($j,$p,$n) Get-OrchestratorExternalDependencyStatus -Job $j -DependencyName $p.Name -Now $n} $Consumer $Parent $now}
try {
 $snapshot=& $versionGate {
  param($b)
  $script:Settings=[pscustomobject]@{
   OrchestratorScriptPath=(Join-Path $b 'SmartM365-Inventory-Orchestrator.ps1')
   CoreModulePath=(Join-Path $b '../../Modules/SmartM365.Core/SmartM365.Core.psd1')
   DistributedModulePath=(Join-Path $b 'SmartM365.Orchestrator.Distributed.psm1')
   ManagementModulePath=(Join-Path $b 'SmartM365.Orchestrator.Management.psm1')
   PipelineModulePath=(Join-Path $b 'SmartM365.Orchestrator.Pipeline.psm1')
   AutoRecycleOnRuntimeUpdate=$true;RuntimeUpdateCheckIntervalSeconds=60
   RuntimeUpdateStableChecks=2;MonitorCoreModuleVersion=$false;SendRuntimeUpdateEmail=$false
  }
  $current=Get-OrchestratorRuntimeSnapshot
  $script:RuntimeUpdateBaseline=$current.PSObject.Copy()
  $script:RuntimeUpdateBaseline.ScriptVersion=[version]'1.5.64'
  $script:RuntimeUpdateBaseline.ScriptFingerprint='previous-runtime'
  $script:LastRuntimeUpdateCheckUtc=[datetime]::MinValue
  $script:RuntimeUpdateCandidateIdentity='';$script:RuntimeUpdateCandidateStableCount=0
  $current
 } $base
 $headerVersion=[regex]::Match($ast.Extent.Text,'(?m)^\.VERSION\s*\r?\n(?<Version>\d+\.\d+\.\d+)').Groups['Version'].Value
 Check ([string]$snapshot.ScriptVersion -ceq $headerVersion) 'Header and executable ScriptVersion differ: automatic recycle would use the wrong version.'
 Check ($snapshot.ScriptVersion -gt [version]'1.5.64') 'The delivered runtime is not newer than the synthetic 1.5.64 baseline.'
 Check (-not (& $versionGate {param($n) Test-OrchestratorRuntimeUpdate -Now $n} $now)) 'Runtime update bypassed its first stability check.'
 Check (& $versionGate {param($n) Test-OrchestratorRuntimeUpdate -Now $n} $now.AddSeconds(61)) 'The update from the synthetic 1.5.64 baseline was not accepted after two stable checks.'
 New-Item -ItemType Directory -Path $claims -Force|Out-Null
 $jobsByName=@{};foreach($j in $document.Jobs){$jobsByName[$j.Name]=$j}
 & $gate {param($c,$j) $script:Settings=[pscustomobject]@{DistributedSchedulingEnabled=$true;ElectionClaimsPath=$c};$script:Manifest=[pscustomobject]@{JobsByName=$j};$script:FreshCalls=0} $claims $jobsByName
 Check (Test-SmartM365OrchestratorJobsDocument -Document $document).Valid 'Template dependency graph is invalid.'
 Check ((@($licences.DependsOn) -join ',') -ceq 'EXO-Mailboxes-Inventory-Fast,M365-Usage-Inventory,Intune-Devices-Inventory') 'Licences must retain EXO and wait for both producers.'
 Check ($licences.DependencyMode -eq 'LatestOccurrence') 'Licences lost its daily occurrence rule.'
 $normalizationErrors=[Collections.Generic.List[string]]::new()
 $normalized=& $gate {param($j,$e) ConvertTo-NormalizedJob -RawJob $j -Errors $e} $prepare $normalizationErrors
 Check ($normalizationErrors.Count -eq 0 -and (@($normalized.LatestOccurrenceDependencies) -join ',') -ceq $licences.Name) 'Runtime normalization lost the override.'
 foreach($p in @($usage,$devices)) {
  Check ([timespan]::Parse($licences.Schedule.Times[0]) -ge [timespan]::Parse($p.Schedule.Times[0])) ('Licences is due before '+$p.Name+' and could use yesterday data.')
 }

 foreach($parent in @($usage,$devices,$licences)){
  $consumer=if($parent.Name -eq $licences.Name){$prepare}else{$licences}
  $when=& $gate {param($p,$n) Get-LatestPastOccurrence -Job $p -Now $n} $parent $now
  Write-Claim -Job $parent -When $when.AddDays(-1) -Status 'Success'
  Check ((Status $consumer $parent) -eq 'Waiting') ('Yesterday success released '+$parent.Name)
  foreach($state in @('Running','RetryScheduled','Failed','Success')){
   Write-Claim -Job $parent -When $when -Status $state
   $expected=if($state -eq 'Success'){'Ready'}elseif($state -eq 'Failed' -and -not $parent.ContinueOnError){'Failed'}else{'Waiting'}
   Check ((Status $consumer $parent) -eq $expected) ('Wrong dependency gate for '+$parent.Name+' '+$state)
  }
 }
 Check ((& $gate {$script:FreshCalls}) -eq 0) 'Latest licensing occurrence fell back to a fresh older receipt.'
 Check ((Status $prepare $devices) -eq 'Ready') 'Other Prepare dependencies must retain FreshSuccess.'
 Check ((& $gate {$script:FreshCalls}) -eq 1) 'Prepare freshness rule was lost for other producers.'
 $rows=@(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $shared -JobsDocument $document -JobName $prepare.Name -Now $now)
 Check (@($rows | Where-Object {$_.Dependency -eq $licences.Name -and $_.Rule -eq 'LatestOccurrence' -and $_.State -eq 'Ready'}).Count -eq 1) 'Insights differs from the runtime licence rule.'
 Check (@($rows | Where-Object {$_.Dependency -eq $devices.Name -and $_.Rule -eq 'FreshSuccess'}).Count -eq 1) 'Insights lost other freshness rules.'
 foreach($invalid in @(@('unknown'),@($licences.Name,$licences.Name),@(''))){
  $copy=$document | ConvertTo-Json -Depth 40 | ConvertFrom-Json
  ($copy.Jobs | Where-Object Name -eq $prepare.Name).LatestOccurrenceDependencies=$invalid
  Check (-not (Test-SmartM365OrchestratorJobsDocument -Document $copy).Valid) 'Invalid occurrence override accepted.'
  $normalizationErrors=[Collections.Generic.List[string]]::new()
  & $gate {param($j,$e) ConvertTo-NormalizedJob -RawJob $j -Errors $e} ($copy.Jobs | Where-Object Name -eq $prepare.Name) $normalizationErrors | Out-Null
  Check ($normalizationErrors.Count -gt 0) 'Runtime accepted an invalid occurrence override.'
 }
 $source=$ast.Extent.Text
 Check ($source.Contains('LatestOccurrenceDependencies = $latestOccurrenceDependencies')) 'Normalized jobs discarded the policy.'
 Check ($source.Contains('Get-OrchestratorEffectiveDependencyMode -Job $job -DependencyName $dep')) 'Scheduled path does not apply per-dependency policy.'
 foreach($retired in @('Exchange-HybridIdentity-Issues','Intune-Windows11-Readiness-Issues')){Check (@($document.Jobs | Where-Object {$_.Name -eq $retired -or $retired -in @($_.DependsOn)}).Count -eq 0) 'Lot 2 retirement regressed.'}
 "License dependency offline checks passed: $checks"
} finally {
 $global:SmartM365TenantKey=$previousTenant
 $full=[IO.Path]::GetFullPath($root);$temp=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\'
 if($full.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $full -Leaf) -match '^SmartM365-LicenseDeps-[a-f0-9]{32}$' -and (Test-Path -LiteralPath $full)){Remove-Item -LiteralPath $full -Recurse -Force}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDajEj9P4FnfTWP
# HDs0sre6Q8FyondgVqi/VPgxdkWXPKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHaLDZ6SAreaXDqU5QnD4AozFxktCd+6y0g0fMSAjEqjMA0GCSqG
# SIb3DQEBAQUABIIBgEWarKrT4/LQSqutDMRJXq2bQ8BKv+mY545FG4vs71BxfvTa
# fdWU4iH89YOGq9pP1jtiyUEjJ2F8lCOAf8gRYQfEU/sM8baqU3V0iPw/W1bFdvLY
# fsyLT4DDBRCISja140OqwAUPlbzDcqlOGTVpjTs4WbfTrBwji/aBbAvA3bLpGZBy
# IVstz5dkoRfgeVKqhzvVwuRF8ggdfbUjg3ELWEI7LXhtzkWh/8duYn3+Dp7ypmgI
# VBP+2LEITAQBUv6Yd3eAChPY0/518pDPlCRenh8bFQWeowm+act7R84fgArdaenJ
# NFLdRzw6pxrtUJcfef/lJy4WriwN5VIGAq8DR8Kznz25ptvmYq5q0WHkKwA27StI
# 0xMGVHTrP1BWnZi4GL3v+li/Q4nKeUrBas9L+xYCXkJvGH1gWIcKRxgCClLRXe6z
# nOz/HHcz8iMLwO+rYlcCupQ3RY1/JBnBHWF5H4nhdB5hm2qW+2ssE+9XgQb/LvFc
# t+8ZdKoKKM5PIg7zh6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkyMTIz
# MjRaMC8GCSqGSIb3DQEJBDEiBCCsjctEZRpfpDcmo1rHDZE1nSBFAze062zlIYKH
# 3dCQxjANBgkqhkiG9w0BAQEFAASCAgBfYGyT6PRHl68+pQ88H3H4J9thNla6190k
# NE2RGWx1yWo5cZ0NVsRFRkEbTfL6vHvSYWxHom7eGmpiJBes5MKUIs/vGpLDhIwZ
# vykyG0bM8p+pq0kOg1XHfjFW3Hff4WVzdEeZgUedxwBW4H5XmLV5ITB/h1NCacdD
# 3rgu2xcdQr0tttE0AwuuYhT73ww/vOsnLaZ97feJyVskgR/IvNL9LlaF3tv8yBT7
# a1sMiKdeJPuch7LMBOIk3TFg0AhtTpjquw3NAhTp8GuU2T/VP3kXeWYvVuJrVfif
# yY9AYyVDpl4y1bsNMzIpoCFX3PWfnQipJXCtDOjnNXxk1a8tbgpmGBHP0HoGQzxK
# Ax03+Ljj65FofXBXHzOX++5TlHUqvdf6LZP755jdZdLUY8ydv3rt9PNJloTHNSdZ
# QIxyv8DEEPon5O3J6EM5M/ARFu9kyONTeoLPgaDITUko54fvE9yKmhk0+rTqyq2a
# dBsezQqB6nYb7yKER1vK2IZqUQyxdKoq3729qkKCfQJ3YF+zvH3U5EOr4v7PGgNP
# lv4O/wyWD4qiMB1SL3I/R96h0ntSx9lxIxEMcqa3ew2MNbiVGn7Nh60r23dwl8vQ
# lJ5Jsj5CAFhKdJsFe6qDa+TYjHOAqzifcUKEG+orHpyvvwVwDhVuHZx/9YbcT8sq
# z2TduA/rFQ==
# SIG # End signature block
