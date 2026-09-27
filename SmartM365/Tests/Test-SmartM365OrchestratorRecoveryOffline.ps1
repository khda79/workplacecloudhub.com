<#
.SYNOPSIS
Synthetic recovery and resident-lock tests; no scheduler entry point.
.VERSION
1.0.0
#>
[CmdletBinding()]
param([string]$SourceRoot,[string]$ResultPath)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
if(-not $SourceRoot){$SourceRoot=Split-Path $PSScriptRoot -Parent}
$source=Join-Path $SourceRoot 'SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$t=$null;$e=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$t,[ref]$e)
if($e.Count){throw 'Source parse failed'}
$definitions=foreach($name in @('Enter-OrchestratorLock','Exit-OrchestratorLock','Test-ProcessMatchesRecord','Restore-RunningJobs','Update-RunningJobs','Get-RunningJobTimeoutWindow','ConvertFrom-StateTime','ConvertTo-StateTime','Write-OrchestratorHeartbeat')){
 $n=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
 if(-not $n){throw "Missing $name"};$n.Extent.Text
}
if(-not ('SyntheticRecoveryProcess' -as [type])){
Add-Type -TypeDefinition @'
using System;
public class SyntheticRecoveryProcess {
 public int Id = 9999999;
 public string ProcessName = "pwsh";
 public DateTime Started = DateTime.Now;
 public bool DenyStart;
 public DateTime StartTime {get {if(DenyStart) throw new UnauthorizedAccessException("Synthetic StartTime denied"); return Started;}}
 public bool HasExited;
 public int ExitCode;
 public DateTime ExitTime = DateTime.Now;
 public IntPtr Handle {get {return IntPtr.Zero;}}
 public void Refresh() {}
 public bool WaitForExit(int delay) {return HasExited;}
}
'@
}
$results=New-Object 'System.Collections.Generic.List[object]'
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-Recovery-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($root)
$modules=New-Object 'System.Collections.Generic.List[object]'
function New-CaseModule {
 param([string]$Folder)
 [void][IO.Directory]::CreateDirectory($Folder)
 $m=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join [Environment]::NewLine))
 & $m {
  param($folder)
  $script:Settings=@{LockPath=(Join-Path $folder 'Orchestrator.lock');ElectionClaimGraceMinutes=5;ConcurrencyLeasesPath=$folder;SharedDataFolderPath=$folder;PeerHeartbeatStaleMinutes=5}
  $script:LockOwned=$false;$script:RunningJobs=@{};$script:Completed=0;$script:Renewed=0
  $script:ProcessMode='missing';$script:ReadHook=$null
  $script:MockProcess=New-Object SyntheticRecoveryProcess
  $script:Manifest=@{OrderedJobs=@();JobsByName=@{}}
  $script:State=@{Jobs=@{Synthetic=@{Running=@{Pid=9999999;StartTime=$script:MockProcess.Started.ToString('o');ScheduledOccurrence=$script:MockProcess.Started.ToString('o');LogPath='synthetic.log';Attempt=0;TimeoutMinutes=60;ClaimPath='synthetic-claim';ConcurrencyLeasePath='synthetic-lease.json';ConcurrencyLeaseId='synthetic-id'};PendingRetry=$null}}}
  function script:Get-Process {
   [CmdletBinding()]param([int]$Id)
   if($script:ProcessMode -eq 'denied'){Write-Error 'Synthetic access denied' -Category PermissionDenied -ErrorId SyntheticDenied;return}
   if($script:ProcessMode -eq 'missing'){Write-Error 'Synthetic PID absent' -Category ObjectNotFound -ErrorId NoProcessFoundForGivenId;return}
   return $script:MockProcess
  }
  function script:Get-Content {
   [CmdletBinding()]param($LiteralPath,[switch]$Raw)
   $value=Microsoft.PowerShell.Management\Get-Content -LiteralPath $LiteralPath -Raw -ErrorAction Stop
   if($script:ReadHook){$hook=$script:ReadHook;$script:ReadHook=$null;& $hook}
   return $value
  }
  function script:Write-OrchestratorLog {param($Message,$Level)}
  function script:Write-OrchestratorRuntimeUpdateWarning {param($Key,$Message,$Now)}
  function script:Save-OrchestratorState {}
  function script:Get-OrchestratorPendingJobSnapshot {param($Now) return @()}
  function script:Write-FileAtomically {param($Path,$Content) $script:Heartbeat=$Content|ConvertFrom-Json}
  $script:StartTime=Get-Date;$script:LifetimeDeadline=(Get-Date).AddHours(1)
  function script:Sync-RunningJobConcurrencyLease {param($JobName,$RunInfo,$Now) $script:Renewed++}
  function script:Enter-SmartM365OrchestratorConcurrencyLease {param($LeasesRootPath,$ConcurrencyKey,$JobName,$Occurrence,$OwnerServer,$SafeMinutes,$HeartbeatRootPath,$HeartbeatStaleMinutes) return @{Acquired=$true;LeasePath='synthetic-lease.json';Lease=@{LeaseId='synthetic-id'}}}
  function script:Set-SmartM365OrchestratorConcurrencyLease {param($LeasePath,$LeaseId,$OwnerServer,$SafeUntilUtc) return $true}
  function script:Complete-JobRun {param($JobName,$RunInfo,$StatusHint,$ExitCode,$EndTime,$ErrorText) $script:Completed++;$script:Status=$StatusHint;$script:State.Jobs[$JobName].Running=$null;$script:RunningJobs.Remove($JobName)}
  function script:Stop-ProcessTree {throw 'Unexpected process kill'}
  function script:Test-OrchestratorStatePersistenceReady {return $true}
 } $Folder
 $modules.Add($m);return $m
}
function Assert-Case {param([bool]$Value,[string]$Message) if(-not $Value){throw $Message}}
function Case {param([string]$Name,[scriptblock]$Body)
 try {& $Body;$results.Add([pscustomobject]@{Name=$Name;Passed=$true;Error=''})}
 catch {$results.Add([pscustomobject]@{Name=$Name;Passed=$false;Error=$_.Exception.Message})}
}
try {
 foreach($mode in @('denied','start-denied','name-denied')){
  Case "$mode is uncertainty, not absence" {
   $m=New-CaseModule (Join-Path $root $mode);$caught=$false
   try {& $m {param($mode) $script:ProcessMode=$mode;if($mode -eq 'start-denied'){$script:ProcessMode='found';$script:MockProcess.DenyStart=$true};if($mode -eq 'name-denied'){$script:ProcessMode='found';$script:MockProcess.ProcessName=$null};Test-ProcessMatchesRecord 9999999 $script:MockProcess.Started} $mode|Out-Null}catch{$caught=$true}
   Assert-Case $caught 'Access denial returned no process.'
  }
 }
 Case 'Missing and reused PID remain rejected; matching identity accepted' {
  $m=New-CaseModule (Join-Path $root 'identity')
  $p=& $m {Test-ProcessMatchesRecord 9999999 $script:MockProcess.Started}
  Assert-Case ($null -eq $p) 'Missing PID not recognized.'
  $p=& $m {$script:ProcessMode='found';Test-ProcessMatchesRecord 9999999 $script:MockProcess.Started.AddMinutes(-1)}
  Assert-Case ($null -eq $p) 'Reused PID accepted.'
  $p=& $m {Test-ProcessMatchesRecord 9999999 $script:MockProcess.Started}
  Assert-Case ($null -ne $p) 'Matching identity rejected.'
 }
 Case 'Uncertain re-adoption retains running gate and lease' {
  $m=New-CaseModule (Join-Path $root 'restore')
  & $m {$script:ProcessMode='denied';Restore-RunningJobs}
  $s=& $m {@{Completed=$script:Completed;Running=$script:RunningJobs.ContainsKey('Synthetic');Record=$script:State.Jobs.Synthetic.Running;Renewed=$script:Renewed}}
  Assert-Case ($s.Completed -eq 0 -and $s.Running -and $null -ne $s.Record -and $s.Renewed -gt 0) 'Uncertain process lost supervision.'
 }
 Case 'Inspection recovery re-adopts on a later tick' {
  $m=New-CaseModule (Join-Path $root 'recover')
  & $m {$script:ProcessMode='denied';Restore-RunningJobs;$script:ProcessMode='found';Update-RunningJobs (Get-Date)}
  $s=& $m {@{Completed=$script:Completed;Process=$script:RunningJobs.Synthetic.Process}}
  Assert-Case ($s.Completed -eq 0 -and $null -ne $s.Process) 'Later inspection did not resume supervision.'
 }
 foreach($timeout in @($false,$true)){
  Case "Confirmed disappearance after uncertainty; TimeoutRequested=$timeout" {
   $m=New-CaseModule (Join-Path $root ('gone-'+$timeout))
   & $m {param($timeout) $script:State.Jobs.Synthetic.Running.TimeoutRequested=$timeout;$script:ProcessMode='denied';Restore-RunningJobs;if($script:Completed){throw 'Premature completion'};$script:ProcessMode='missing';Update-RunningJobs (Get-Date);Update-RunningJobs (Get-Date)} $timeout
   $s=& $m {@{Completed=$script:Completed;Status=$script:Status}}
   $expected=if($timeout){'TimedOut'}else{'Interrupted'}
   Assert-Case ($s.Completed -eq 1 -and $s.Status -eq $expected) 'Confirmed disappearance completion changed.'
  }
 }
 Case 'Heartbeat retains recorded PID during uncertain recovery' {
  $m=New-CaseModule (Join-Path $root 'heartbeat')
  & $m {$script:ProcessMode='denied';Restore-RunningJobs;Write-OrchestratorHeartbeat}
  $s=& $m {$script:Heartbeat}
  Assert-Case ($s.RunningJobs.Count -eq 1 -and $s.RunningJobs[0].Pid -eq 9999999) 'Uncertain recovery lost heartbeat PID.'
 }
 Case 'Repeated denied inspection keeps ownership and renews supervision' {
  $m=New-CaseModule (Join-Path $root 'repeat')
  & $m {$script:ProcessMode='denied';Restore-RunningJobs;Update-RunningJobs (Get-Date);Update-RunningJobs (Get-Date)}
  $s=& $m {@{Completed=$script:Completed;Renewed=$script:Renewed;Running=$script:RunningJobs.ContainsKey('Synthetic')}}
  Assert-Case ($s.Completed -eq 0 -and $s.Running -and $s.Renewed -eq 3) 'Repeated inspection lost protection.'
 }
 Case 'Legacy live owner lock remains protected' {
  $folder=Join-Path $root 'legacy-live';$m=New-CaseModule $folder
  [IO.File]::WriteAllText((Join-Path $folder 'Orchestrator.lock'),'{"Pid":9999999}')
  $value=& $m {$script:ProcessMode='found';$script:LockOwned=Enter-OrchestratorLock;return $script:LockOwned}
  Assert-Case (-not $value) 'Live legacy owner was displaced.'
 }
 Case 'Abandoned owner handles allow stale recovery' {
  $folder=Join-Path $root 'abandoned';$a=New-CaseModule $folder;$b=New-CaseModule $folder
  & $a {$script:LockOwned=Enter-OrchestratorLock;if($script:ResidentLockStream){$script:ResidentLockStream.Dispose()};if($script:ResidentLockGuard){$script:ResidentLockGuard.Dispose()};$script:LockOwned=$false}
  $value=& $b {$script:LockOwned=Enter-OrchestratorLock;return $script:LockOwned}
  Assert-Case $value 'Stale recovery failed after handles were abandoned.'
 }
 Case 'Malformed lock payload is preserved' {
  $folder=Join-Path $root 'corrupt';$m=New-CaseModule $folder
  [IO.File]::WriteAllText((Join-Path $folder 'Orchestrator.lock'),'{unfinished')
  $acquired=& $m {$script:LockOwned=Enter-OrchestratorLock;return $script:LockOwned}
  Assert-Case (-not $acquired -and [IO.File]::ReadAllText((Join-Path $folder 'Orchestrator.lock')) -eq '{unfinished') 'Unknown lock was replaced.'
 }
 Case 'Resident handle excludes contenders despite failed PID lookup' {
  $folder=Join-Path $root 'resident';$a=New-CaseModule $folder;$b=New-CaseModule $folder
  $first=& $a {$script:LockOwned=Enter-OrchestratorLock;return $script:LockOwned}
  $second=& $b {$script:LockOwned=Enter-OrchestratorLock;return $script:LockOwned}
  Assert-Case ($first -and -not $second) 'Both residents acquired the lock.'
 }
 Case 'Stale reader cannot remove newly acquired owner lock' {
  $folder=Join-Path $root 'race';$a=New-CaseModule $folder;$b=New-CaseModule $folder
  [IO.File]::WriteAllText((Join-Path $folder 'Orchestrator.lock'),'{"Pid":9999998}')
  & $a {param($other) $script:Other=$other;$script:ReadHook={$script:ContenderAcquired=& $script:Other {$script:LockOwned=Enter-OrchestratorLock;return $script:LockOwned}}} $b
  $first=& $a {$script:LockOwned=Enter-OrchestratorLock;return $script:LockOwned}
  $second=& $a {$script:ContenderAcquired}
  Assert-Case ($first -and -not $second) 'Both stale-recovery contenders acquired.'
 }
 Case 'Release permits new owner; duplicate exit cannot remove its lock' {
  $folder=Join-Path $root 'release';$a=New-CaseModule $folder;$b=New-CaseModule $folder
  & $a {$script:LockOwned=Enter-OrchestratorLock;Exit-OrchestratorLock}
  $second=& $b {$script:LockOwned=Enter-OrchestratorLock;return $script:LockOwned}
  & $a {Exit-OrchestratorLock}
  Assert-Case ($second -and [IO.File]::Exists((Join-Path $folder 'Orchestrator.lock'))) 'Duplicate exit removed new lock.'
 }
 Case 'Uncertain incumbent PID preserves existing payload' {
  $folder=Join-Path $root 'incumbent';$m=New-CaseModule $folder;$payload='{"Pid":9999998}'
  [IO.File]::WriteAllText((Join-Path $folder 'Orchestrator.lock'),$payload)
  $acquired=& $m {$script:ProcessMode='denied';$script:LockOwned=Enter-OrchestratorLock;return $script:LockOwned}
  Assert-Case (-not $acquired -and [IO.File]::ReadAllText((Join-Path $folder 'Orchestrator.lock')) -eq $payload) 'Uncertain incumbent displaced.'
 }
}
finally {
 foreach($m in $modules){try {& $m {Exit-OrchestratorLock}}catch{};Remove-Module $m -Force}
 $resolved=[IO.Path]::GetFullPath($root);$parent=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
 if(-not $resolved.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolved) -notlike 'SmartInventory-Recovery-*'){throw 'Unsafe cleanup'}
 Remove-Item -LiteralPath $resolved -Recurse -Force
}
$failed=@($results|Where-Object {-not $_.Passed})
$report=[pscustomobject]@{PowerShell=$PSVersionTable.PSVersion.ToString();Total=$results.Count;Passed=$results.Count-$failed.Count;Failed=$failed.Count;Cases=$results.ToArray()}
if($ResultPath){$report|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $ResultPath -Encoding UTF8}
$report|Select-Object PowerShell,Total,Passed,Failed|Format-Table
$failed|Format-Table Name,Error -Wrap
if($failed.Count){exit 1}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC0COMPPv61/yNa
# dYGXdfCygupIKA3/J+8kY2KlOQyOXaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFSO171E+DeNilurGJhp/gaBWrSBYv4/eSXM5uh4GL5+MA0GCSqG
# SIb3DQEBAQUABIIBgHSsmj27E0CKJi3Y77iAOk13KQ43oZECrvYShABfsgrVUBdt
# LU9LKdqYg3kEP/KJg6z3mqGFhEfgBZh5AKxyuuHC4MOmoJXWK0azbOx54JxHGa04
# 1cB4y7S4fJlikU/Ro+GJSgRjAte2o8mXKSXCK5ffnRQMvCnZp9Ub1ZvIsJUVVnDx
# TgHw1F9XrPtAr+NU3FVWRK7FNbkdxYWXexYM5bExKEvXOje0YSSV9RWz2tu6p5cj
# 3hOz72lP3IU6llOC75UJeBbYTWTzs1TEKZnEl3h2sGKFg6AcwDKIABFfc7OUT6j5
# ycbC6jSzISztvsDdTUyjxovDygOMcsZYmec8pS+8aoR2Noi2Jt0ey3UNQLpOeOMG
# vC3cfka/D2nlAIssZAEVv5JHapPzZhNjDDjhiBGGGIllw/Pb5ox7cSIyZVUorr8m
# 8D9zT7iwXleEoSmxpXnsDCC5FQH+iM+eIt6kizzDt1oKBSHHz7c4NHgvmObfSxW+
# ZqkJDzJ0mBr6iERW3KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU3
# NThaMC8GCSqGSIb3DQEJBDEiBCBTRzTZPY1FQlM7x7ZkSHQLJmOrSfGUM7wARlsj
# 2xuksTANBgkqhkiG9w0BAQEFAASCAgCHpG4VHhKBXaZsHAaoIVXWTmcemPq4EIFG
# c8tvcYk4chJC2eCr3VQn+dow+SOY5ZxwU9skdWV+VjeRFa6g9p3Fs3o2E/QqatTG
# eUqtgGRyPQG+njZ+mbVAJvH8Br9lBDvjexnaEo5cY3hIGVK54j/icO6tDlRArtZ3
# u+qgY9vEwaxsoY71nq4RahSgLw6W4Bdz35GKsC+TTBReZSoAHmPLqXLhybVI83/b
# rBN9mGII5tM5EyeOGUWTMScR6vQ7XWEznaAvKp+QUtlJmxLHVW0GnPUDU94/7Bd6
# /TNC5mBrGAx0nh1qJHGXz9WQQoNdv1mzNW/fJmNFAUsyn/hZnKh6vtc/T3lZkosO
# ycMi4Py2kk9Zw8FG1AlBNTws0T/hgnrji39dD99LbkwWzvKCuMYmZggoz15Rg8wu
# hNw2bXtnOWJJwNmpMzLI9IPwaFnwggRZYGI0ZpxzYiiAuTDMTUjSMZ5kUXTbA9lk
# 6jjk8S8b8EZnZ3ZNFLS0tsMwXskjsTtX1/1aaMnlIJYiyy0j4WUsBqdCiRyTSWy3
# oBqX/3Fyfjg7JP2Smxx0YslELEkCa1NrzHgpLV0vzX3Elaomnn/Ivf2pbV0sFV76
# TlYqNmTkuhGsxmrvwwyEMopg1i0qWXBuZWD+sQmrrZwtUu6vvEWQOGgxW1UdgtpO
# LDT1S//ddw==
# SIG # End signature block
