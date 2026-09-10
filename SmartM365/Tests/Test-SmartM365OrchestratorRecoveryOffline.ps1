<#
.SYNOPSIS
Synthetic recovery and resident-lock tests; no scheduler entry point.
.VERSION
1.0.0
#>
[CmdletBinding()]
param([string]$SourceRoot,[string]$ResultPath)
$ErrorActionPreference='Stop'
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
  $script:State=@{Jobs=@{Synthetic=@{Running=@{Pid=9999999;StartTime=$script:MockProcess.Started.ToString('o');ScheduledOccurrence=$script:MockProcess.Started.ToString('o');LogPath='synthetic.log';Attempt=0;TimeoutMinutes=60;ClaimPath='synthetic-claim';ConcurrencyLeasePath='synthetic-lease';ConcurrencyLeaseId='synthetic-id'};PendingRetry=$null}}}
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
  function script:Enter-SmartM365OrchestratorConcurrencyLease {param($LeasesRootPath,$ConcurrencyKey,$JobName,$Occurrence,$OwnerServer,$SafeMinutes,$HeartbeatRootPath,$HeartbeatStaleMinutes) return @{Acquired=$true;LeasePath='synthetic-lease';Lease=@{LeaseId='synthetic-id'}}}
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDaF6e+8T0uDrxE
# Y4G6PRfF/JwXm1EDYw4n3tWhVDPncqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDzF97qY2OpI42KLf7Hl4qh
# iEUILliF3nDx71RzWhYR5zANBgkqhkiG9w0BAQEFAASCAYBIOptzv1cd5DDe79z/
# jdYVCq3mh8BgfZj5/GYcYDG7HPq+7dpZG042FRji1lkPj0/B+dgI7s83d3WWGi3O
# ohvZxtx1TeVl6BBhWK13pfy9VMy2L7sCK4xvU2L2omSvsTNL+3Qb7217cRMPmQfi
# IUOqrhFB0sTw/osJTzxZvtaPgqMQMzGzNwMtjD3JKxoUuqoMbx2AofM+tNo6WKfx
# uTziQvT6dOIRfFTt8/ebwQOLcieMfTgtTnws8gDl9qnxeWAmZ9AGL4tychAtCzFl
# enqn7FAvfxWCUMT/r4iSPsU/zs9ngMNOPdUzesaAwZ7CF5OlhxupKkcCwZEHm4CZ
# vTllbwFJAjPBih8BepdiPUwgSCY4LJRIQU0vmuTQT5S9qNMzoDpM/2Bh2RXWnkjx
# YREAm6woz43LIJTMcCPfQ9NhVhy8L0hqxnzuVzZXiREw9+Ggk53wUDz4h+b60t6T
# aYEINpHO5gkc4BLWZdhAOwPKZTulyLyS0X5z3SUPR6/fG18=
# SIG # End signature block
