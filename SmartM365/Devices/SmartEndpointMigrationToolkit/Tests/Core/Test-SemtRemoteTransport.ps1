#requires -Version 5.1
<#
.SYNOPSIS
Synthetic Task Scheduler tests. No scheduled task or remote computer is touched.
.VERSION
1.0.0
#>
$ErrorActionPreference = 'Stop'
class SemtFakeAction { [string]$Path; [string]$Arguments }
class SemtFakeActions {
    [int]$Count = 1
    [SemtFakeAction]$Action = [SemtFakeAction]::new()
    [object] Create([int]$type) { return $this.Action }
    [object] Item([int]$index) { return $this.Action }
}
class SemtFakePrincipal { [string]$UserId; [int]$LogonType; [int]$RunLevel }
class SemtFakeSettings { [bool]$AllowDemandStart; [bool]$DisallowStartIfOnBatteries; [bool]$StopIfGoingOnBatteries; [string]$ExecutionTimeLimit }
class SemtFakeRegistration { [string]$Description }
class SemtFakeDefinition {
    [SemtFakeActions]$Actions = [SemtFakeActions]::new()
    [SemtFakePrincipal]$Principal = [SemtFakePrincipal]::new()
    [SemtFakeSettings]$Settings = [SemtFakeSettings]::new()
    [SemtFakeRegistration]$RegistrationInfo = [SemtFakeRegistration]::new()
}
class SemtFakeTask {
    [SemtFakeDefinition]$Definition
    [int]$State = 3
    [datetime]$LastRunTime = [datetime]::Now
    [int]$LastTaskResult = 0
    [object] Run([object]$arg) { return $null }
}
class SemtFakeFolder {
    [SemtFakeTask]$Task
    [bool]$Deleted = $false
    [int]$Flags
    [bool]$Collide = $false
    [bool]$ChangeAction = $false
    [string]$ReplacementPrincipal = ''
    [object] RegisterTaskDefinition([string]$name,[object]$definition,[int]$flags,[object]$user,[object]$password,[int]$logon,[object]$sddl) {
        if ($this.Collide) { throw 'task already exists' }
        $this.Flags = $flags
        $this.Task = [SemtFakeTask]::new(); $this.Task.Definition = $definition
        return $this.Task
    }
    [object] GetTask([string]$name) { if ($this.ChangeAction) { $this.Task.Definition.Actions.Action.Arguments = 'changed' }; if ($this.ReplacementPrincipal) { $this.Task.Definition.Principal.UserId = $this.ReplacementPrincipal }; return $this.Task }
    [void] DeleteTask([string]$name,[int]$flags) { $this.Deleted = $true }
}
class SemtFakeService {
    [SemtFakeFolder]$Folder = [SemtFakeFolder]::new()
    [void] Connect() { }
    [void] Connect([string]$target) { }
    [object] GetFolder([string]$path) { return $this.Folder }
    [object] NewTask([int]$flags) { return [SemtFakeDefinition]::new() }
}
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$remote = Join-Path $root 'Modules\SmartEndpointMigration.Core\Functions\Semt.Remote.ps1'
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($remote,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw $errors[0].Message }
$fn=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-SemtSystemTaskEndpoint'},$true)
$assignment=$fn.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$operation'},$true)
$expression=$assignment.Right.Extent.Text
$operation=[scriptblock]::Create($expression.Substring(1,$expression.Length-2))
function New-Object { param([string]$ComObject) if ($ComObject -ne 'Schedule.Service') { throw 'Unexpected COM object' }; return $script:fake }
function Start-Sleep { param([int]$Seconds) }
$passed=0
foreach ($mode in @('ScheduledTask','WinRM')) {
    $script:fake=[SemtFakeService]::new()
    $dir=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-test-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($dir)
    $file=Join-Path $dir 'endpoint.ps1'
    [IO.File]::WriteAllText($file,'synthetic')
    $hash=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
    $request=@{Target='synthetic.invalid'; Name='WCH-SEMT-11111111111111111111111111111111'; Arguments='-File "C:\Program Files\test.ps1"'; Script=$file; ShareScript=$file; Sha256=$hash; Mode=$mode}
    $result=& $operation $request
    if (-not $result.Cleaned -or -not $script:fake.Folder.Deleted -or (Test-Path -LiteralPath $dir)) { throw "$mode cleanup failed" }
    if ($script:fake.Folder.Flags -ne 2 -or $script:fake.Folder.Task.Definition.Principal.UserId -ne 'SYSTEM' -or $script:fake.Folder.Task.Definition.Principal.LogonType -ne 5) { throw "$mode SYSTEM or collision guard failed" }
    $passed++
}
$systemSid = [System.Security.Principal.SecurityIdentifier]::new('S-1-5-18')
$localizedName = $systemSid.Translate([System.Security.Principal.NTAccount]).Value
$shortName = ($localizedName -split '\\')[-1]
foreach ($principal in @('S-1-5-18',$localizedName,$shortName)) {
    $script:fake=[SemtFakeService]::new()
    $script:fake.Folder.ReplacementPrincipal=$principal
    $dir=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-test-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($dir)
    $file=Join-Path $dir 'endpoint.ps1'; [IO.File]::WriteAllText($file,'synthetic')
    $hash=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
    $result=& $operation @{Target='synthetic.invalid';Name='WCH-SEMT-11111111111111111111111111111111';Arguments='test';Script=$file;ShareScript=$file;Sha256=$hash;Mode='ScheduledTask'}
    if (-not $result.Cleaned -or -not $script:fake.Folder.Deleted -or (Test-Path -LiteralPath $dir)) { throw "SYSTEM principal '$principal' was not accepted" }
    $passed++
}
$script:fake=[SemtFakeService]::new()
$script:fake.Folder.ReplacementPrincipal='S-1-5-19'
try { & $operation @{Target='x';Name='x';Arguments='x';Mode='ScheduledTask'}; throw 'Expected rejection of non-SYSTEM principal' }
catch { if ($_.Exception.Message -notmatch 'identity changed') { throw } }
if ($script:fake.Folder.Deleted) { throw 'Non-SYSTEM task was deleted' }
$passed++
$script:fake=[SemtFakeService]::new()
$script:fake.Folder.Collide=$true
try { & $operation @{Target='x';Name='x';Arguments='x';Mode='ScheduledTask'}; throw 'Expected a collision' }
catch { if ($_.Exception.Message -notmatch 'already exists') { throw } }
if ($script:fake.Folder.Deleted) { throw 'Collision deleted a task' }
$passed++
$script:fake=[SemtFakeService]::new()
$script:fake.Folder.ChangeAction=$true
try { & $operation @{Target='x';Name='x';Arguments='x';Mode='ScheduledTask'}; throw 'Expected an identity mismatch' }
catch { if ($_.Exception.Message -notmatch 'identity changed') { throw } }
if ($script:fake.Folder.Deleted) { throw 'Changed task was deleted' }
$passed++
$script:fake=[SemtFakeService]::new()
$dir=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($dir)
$file=Join-Path $dir 'endpoint.ps1'; [IO.File]::WriteAllText($file,'synthetic')
try {
    try { & $operation @{Target='x';Name='x';Arguments='x';Mode='ScheduledTask';Script=$file;ShareScript=$file;Sha256='wrong'}; throw 'Expected a hash mismatch' }
    catch { if ($_.Exception.Message -notmatch 'hash changed') { throw } }
    if (-not (Test-Path -LiteralPath $file)) { throw 'Changed endpoint was deleted' }
    if ($script:fake.Folder.Deleted) { throw 'Task deleted despite changed endpoint' }
    $passed++
} finally { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue; Remove-Item -LiteralPath $dir -Force -ErrorAction SilentlyContinue }
Remove-Item Function:\New-Object
Remove-Item Function:\Start-Sleep
Import-Module (Join-Path $root 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
$catalog=Import-SemtOptionCatalog -PackageRoot $root
foreach ($campaign in @('HybridJoin','Windows11Upgrade')) {
    $resolved=Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $campaign
    if ($resolved['RemoteExecutionMode'].Value -ne 'ScheduledTask') { throw "$campaign default transport differs" }
    $passed++
    foreach ($choice in @('ScheduledTask','WinRM','PsExec')) {
        $resolved=Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $campaign -CliLayer @{RemoteExecutionMode=$choice}
        if ($resolved['RemoteExecutionMode'].Value -ne $choice) { throw "$campaign transport choice $choice was not retained" }
        $passed++
    }
    try { [void](Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $campaign -CliLayer @{RemoteExecutionMode='Unexpected'}); throw 'Invalid choice accepted' }
    catch { if ($_.Exception.Message -eq 'Invalid choice accepted') { throw } }
    $passed++
}
$module=Get-Module SmartEndpointMigration.Core
& $module {
    function script:Start-Job { param($ArgumentList,$ScriptBlock); $script:CapturedRequest=$ArgumentList[0]; return [pscustomobject]@{State='Completed'} }
    function script:Wait-Job { param($Job,$Timeout); return $Job }
    function script:Receive-Job { param($Job); return [pscustomobject]@{Cleaned=$true; ExitCode='0'} }
    function script:Remove-Job { param($Job,[switch]$Force) }
}
$log=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-' + [guid]::NewGuid().ToString('N') + '.log')
try {
    foreach ($mode in @('ScheduledTask','WinRM')) {
        $result=Invoke-SemtSystemTaskEndpoint -Mode $mode -ConnectionTarget 'synthetic.invalid' -CorrelationId '11111111111111111111111111111111' -RemoteScriptPath 'C:\Program Files\test.ps1' -RemoteSharePath '\\synthetic.invalid\C$\test.ps1' -ExpectedSha256 'test' -Arguments @('-AuditOnly','-Value','has space') -LogPath $log
        $captured=& $module { $script:CapturedRequest }
        if (-not $result.Completed -or $captured.Mode -ne $mode -or $captured.Arguments -notmatch '"C:\\Program Files\\test.ps1"' -or $captured.Arguments -notmatch '"has space"') { throw "$mode dispatch failed" }
        $passed++
    }
    & $module {
        $script:FakeClock = 0
        function script:Wait-Job { param($Job,$Timeout); return $null }
        function script:Get-Date { param($Format); $script:FakeClock++; $value=([datetime]'2026-01-01').AddMinutes($script:FakeClock * 2); if ($Format) { return $value.ToString($Format) }; return $value }
        function script:Update-SemtWorkerLease { param($LeasePath) }
    }
    $timeout=Invoke-SemtSystemTaskEndpoint -Mode ScheduledTask -ConnectionTarget 'synthetic.invalid' -CorrelationId '11111111111111111111111111111111' -RemoteScriptPath 'C:\Program Files\test.ps1' -RemoteSharePath '\\synthetic.invalid\C$\test.ps1' -ExpectedSha256 'test' -Arguments @() -LogPath $log -TimeoutMinutes 1
    if (-not $timeout.TimedOut -or $timeout.Completed -or $timeout.Detail -notmatch 'verify before retrying') { throw 'Timeout did not leave the remote state unconfirmed' }
    $passed++
} finally { Remove-Item -LiteralPath $log -Force -ErrorAction SilentlyContinue }
$worker = Join-Path $root 'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1'
$tokens=$null; $errors=$null
$workerAst=[Management.Automation.Language.Parser]::ParseFile($worker,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw $errors[0].Message }
$statusBranch=$workerAst.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith('if ($attemptResult.Completed)') -and $n.Extent.Text.Contains('REMOTE_EXECUTION_ACCESS_DENIED')},$true)
if (-not $statusBranch) { throw 'Transport status branch not found' }
$classify=[scriptblock]::Create($statusBranch.Extent.Text)
function Set-WorkerStatus { param($Status,$Detail,$NextAction); $script:observedStatus=$Status; $script:observedNextAction=$NextAction }
foreach ($exitCode in @('-1073741790','3221225506')) {
    $attemptResult=[pscustomobject]@{Completed=$false;Detail='no result'}
    $transport=[pscustomobject]@{Completed=$true;ExitCode=$exitCode;Detail=''}
    $result=[ordered]@{ExitCode='';JobErrorMessage=''}
    $remoteMode='ScheduledTask'
    & $classify
    if ($script:observedStatus -ne 'REMOTE_EXECUTION_ACCESS_DENIED' -or $script:observedNextAction -ne 'CHECK_TARGET_EXECUTION_ACCESS' -or $result.ExitCode -ne $exitCode) { throw "Incorrect status for exit $exitCode" }
    $passed++
}
Write-Host "Summary: Passed=$passed; Failed=0"

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD8Q9I2MvB62lGT
# fh7kA1OejgPrNRtDWirRvt2pJJOehKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB68wZazcqzHyit/s38DAOR
# E/Bj3DsRK58UuGYakvXA0zANBgkqhkiG9w0BAQEFAASCAYCOGneVh8mpC32Y9ZEv
# biC8o0N2nWU8X8KWbnMqurWD8/CqmHQkBkzzUQ0KwuU3wx2rhQsXw5fD2YEuGlYn
# O/hCI/CuXZ9fZP8Z/88ZQuZS2iubD7HX/2IAlSBUEd1fXPHvoiv21/I0AcFIdYGQ
# iDN4nuV9PSPyITO28jYas/xDgTNaSgqiDfK4ok6PjuOiVHVpfozvM/S3LpKZtoyl
# c8V/5gzGbOtYg7pbBXEz2OmdxusO9fRKeqA3TrLRz/ykJjeLKsi/GyA6D7nHyioK
# gegkr8E8OygxP7S1qWvJVN/8QNnPaUcVbMHMAWCiPPEQZc0mEYFS7ptIIHX00FrE
# DH01IOYLrV6YN/DeYDa80r8QRJEwmIe8a0je2ErHGPnCJCctCft28fuZctOr984w
# OdQVh+YUP7/xBZieiiuolgKHMP0YX+qIhgMpUXi0nL8767xS8LZAUEynYlmBL/2V
# 92OxcQ3rx3kU1azQXOkyZpcObHPMzuLkaw1iLLtQls7a3pY=
# SIG # End signature block
