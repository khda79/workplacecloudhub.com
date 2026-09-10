<#
.SYNOPSIS
Synthetic lifecycle regression tests; no orchestrator entry point is executed.
.VERSION
1.0.0
#>
[CmdletBinding()]
param([string]$SourceRoot, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$source = Join-Path $SourceRoot 'SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'Orchestrator source failed parsing.' }
$names = @('Update-RunningJobs','Restore-RunningJobs','Sync-RunningJobConcurrencyLease',
    'Get-RunningJobTimeoutWindow','Complete-JobRun','Get-JobState','ConvertTo-StateTime',
    'ConvertFrom-StateTime','Invoke-LaunchPhase')
$definitions = foreach ($name in $names) {
    $node = $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name}, $true)
    if (-not $node) { throw "Missing function: $name" }
    $node.Extent.Text
}
$module = New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
$results = New-Object 'System.Collections.Generic.List[object]'
& $module {
    param($OrchestratorSource)
    $script:OrchestratorSource=$OrchestratorSource
    function script:Assert-Case { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message } }
    function script:Reset-Case {
        $script:Now = Get-Date
        $script:Events = New-Object 'System.Collections.Generic.List[string]'
        $script:Rows = New-Object 'System.Collections.Generic.List[object]'
        $script:LeaseUntil = [datetime]::MinValue
        $script:Released = 0; $script:Kills = 0; $script:Launched = 0
        $script:KillMode = 'survive'; $script:Adopt = $true
        $script:Settings = @{ElectionClaimGraceMinutes=5;JobMailMode='Never';MaxConcurrency=4}
        $job = [pscustomobject]@{Name='Synthetic'; TimeoutMinutes=1;MaxRetries=1;RetryDelaySeconds=60;ConcurrencyKey='synthetic-shared';AssignmentMode='Manual';Enabled=$false;DependsOn=@()}
        $script:Manifest = @{JobsByName=@{Synthetic=$job}; OrderedJobs=@($job)}
        $script:Process = [pscustomobject]@{Id=987654;HasExited=$false;ExitCode=0;ExitTime=$script:Now;RefreshFails=$false;WaitFails=$false}
        $script:Process | Add-Member ScriptMethod Refresh { if ($this.RefreshFails) { throw 'Synthetic process access failure' } }
        $script:Process | Add-Member ScriptMethod WaitForExit { param($Milliseconds) if($this.WaitFails){throw 'Synthetic wait failure'}; return $this.HasExited }
        $start = $script:Now.AddMinutes(-10)
        $info = @{Process=$script:Process;StartTime=$start;Occurrence=$start;LogPath='synthetic.log';Attempt=0;TimeoutMinutes=1;ClaimPath='synthetic-claim';ConcurrencyLeasePath='synthetic-lease';ConcurrencyLeaseId='synthetic-id'}
        $script:RunningJobs = @{Synthetic=$info}
        $script:State = @{Jobs=@{Synthetic=@{Running=@{Pid=987654;StartTime=$start.ToString('o');ScheduledOccurrence=$start.ToString('o');LogPath='synthetic.log';Attempt=0;TimeoutMinutes=1;ClaimPath='synthetic-claim';ConcurrencyLeasePath='synthetic-lease';ConcurrencyLeaseId='synthetic-id'};PendingRetry=$null;LastStatus='Running';LastScheduledOccurrence=$start.ToString('o')}}}
        $script:ForcedPending = @()
        $script:StatePersistenceHealthy = $true
    }
    function script:Write-OrchestratorLog { param($Message,$Level) }
    function script:Write-OrchestratorRuntimeUpdateWarning { param($Key,$Message,$Now) }
    function script:Save-OrchestratorState {
        $script:Events.Add('save')
        $script:SavedState = [Management.Automation.PSSerializer]::Serialize($script:State)
        $script:SavedJson = $script:State | ConvertTo-Json -Depth 10
    }
    function script:Stop-ProcessTree {
        param($TargetPid)
        Assert-Case ($TargetPid -eq 987654) 'Unexpected process identity.'
        $script:Kills++; $script:Events.Add('kill')
        if($script:KillMode -eq 'exit'){$script:Process.HasExited=$true}
        if($script:KillMode -eq 'probe-error'){$script:Process.RefreshFails=$true}
    }
    function script:Set-SmartM365OrchestratorConcurrencyLease {
        param($LeasePath,$LeaseId,$OwnerServer,$SafeUntilUtc)
        $script:LeaseUntil=$SafeUntilUtc; $script:Events.Add('renew'); return $true
    }
    function script:Exit-SmartM365OrchestratorConcurrencyLease {
        param($LeasePath,$LeaseId,$OwnerServer) $script:Released++; return $true
    }
    function script:Set-SmartM365OrchestratorOccurrenceClaim {
        param($ClaimPath,$OwnerServer,$Status,$Attempt,$SafeMinutes,$Detail) $script:ClaimStatus=$Status
    }
    function script:Add-JobRunCsvRow {
        param($JobName,$ScheduledTime,$StartTime,$EndTime,$DurationSec,$ExitCode,$Status,$RetryCount,$LogPath)
        $script:Rows.Add([pscustomobject]@{Status=$Status;ExitCode=$ExitCode})
    }
    function script:Invoke-OrchestratorSharePointUpload { param($LocalFilePath,$Reason,[switch]$Force) return $false }
    function script:Get-JobRunsCsvPath { 'synthetic-runs.csv' }
    function script:Test-JobRunSuccessEvidence { param($JobName,$LogPath,$DurationSec,$ManifestJob) return @{IsValid=$true} }
    function script:Send-JobResultEmail { throw 'Unexpected mail path.' }
    function script:Test-ProcessMatchesRecord { param($RecordedPid,$ExpectedStartTime,$ExpectedProcessName) if($script:Adopt){return $script:Process}; return $null }
    function script:Test-Path { param($LiteralPath,$PathType) return $false }
    function script:Enter-SmartM365OrchestratorConcurrencyLease {
        param($LeasesRootPath,$ConcurrencyKey,$JobName,$Occurrence,$OwnerServer,$SafeMinutes,$HeartbeatRootPath,$HeartbeatStaleMinutes)
        return @{Acquired=$true;LeasePath='synthetic-lease';Lease=@{LeaseId='synthetic-id'}}
    }
    function script:Test-JobSelected { param($JobName) return $true }
    function script:Test-JobAllowedOnServer { param($Job,[switch]$AllowManual) return $true }
    function script:Get-DueOccurrence { param($Job,$LastOccurrence,$Now) return $null }
    function script:Clear-DependencyWaitLog { param($JobName) }
    function script:Start-InventoryJob { param($Job,$Occurrence,$Attempt,$ClaimPath,$ConcurrencyLeasePath,$ConcurrencyLeaseId) $script:Launched++ }
    function script:Test-OrchestratorStatePersistenceReady { return $script:StatePersistenceHealthy }
    function script:Test-OrchestratorAuthenticodeFile { throw 'Unexpected launch preparation.' }
    function script:Resume-Case {
        $script:State=[Management.Automation.PSSerializer]::Deserialize($script:SavedState)
        $script:RunningJobs=@{}
        Restore-RunningJobs
    }
} $source
function Test-LifecycleCase {
    param([string]$Name,[scriptblock]$Body)
    try { & $module {Reset-Case}; & $module $Body; $results.Add([pscustomobject]@{Name=$Name;Passed=$true;Error=''}) }
    catch { $results.Add([pscustomobject]@{Name=$Name;Passed=$false;Error=$_.Exception.Message}) }
}
try {
    Test-LifecycleCase 'Failed kill retains running state, lease and retry gate' {
        Update-RunningJobs $script:Now
        Assert-Case ($script:RunningJobs.ContainsKey('Synthetic') -and $null -ne $script:State.Jobs.Synthetic.Running) 'Live job was completed.'
        Assert-Case ($script:Released -eq 0 -and $null -eq $script:State.Jobs.Synthetic.PendingRetry -and $script:Rows.Count -eq 0) 'Lease/retry/result was finalized before exit.'
    }
    Test-LifecycleCase 'Timeout intent saved before process termination' {
        Update-RunningJobs $script:Now
        Assert-Case ($script:Events.IndexOf('save') -ge 0 -and $script:Events.IndexOf('save') -lt $script:Events.IndexOf('kill')) 'Timeout intent was not persisted before kill.'
        Assert-Case ($script:SavedJson -match '"TimeoutRequested":\s*true') 'Durable timeout intent missing.'
    }
    Test-LifecycleCase 'Expired lease renewed while termination pending' {
        Update-RunningJobs $script:Now
        Assert-Case ($script:LeaseUntil -ge $script:Now.ToUniversalTime().AddMinutes(5)) 'Active overdue lease was not extended.'
        Assert-Case ($script:Events.IndexOf('renew') -lt $script:Events.IndexOf('kill')) 'Lease was not renewed before the wait.'
    }
    Test-LifecycleCase 'Failed kill blocks forced relaunch of same job' {
        Update-RunningJobs $script:Now
        $script:ForcedPending=@('Synthetic')
        Invoke-LaunchPhase $script:Now
        Assert-Case ($script:Launched -eq 0 -and $script:ForcedPending.Count -eq 0) 'Forced overlap was not blocked.'
    }
    Test-LifecycleCase 'Failed kill blocks sibling sharing ConcurrencyKey' {
        Update-RunningJobs $script:Now
        $sibling=[pscustomobject]@{Name='Sibling';TimeoutMinutes=1;MaxRetries=0;RetryDelaySeconds=60;ConcurrencyKey='synthetic-shared';AssignmentMode='Manual';Enabled=$false;DependsOn=@()}
        $script:Manifest.JobsByName.Sibling=$sibling
        $script:Manifest.OrderedJobs=@($sibling)
        $script:ForcedPending=@('Sibling')
        Invoke-LaunchPhase $script:Now
        Assert-Case ($script:Launched -eq 0) 'Shared-output sibling launched during failed termination.'
    }
    Test-LifecycleCase 'Later exit zero remains TimedOut and completes once' {
        Update-RunningJobs $script:Now
        $script:Process.HasExited=$true
        Update-RunningJobs ($script:Now.AddSeconds(60))
        Update-RunningJobs ($script:Now.AddSeconds(120))
        Assert-Case ($script:State.Jobs.Synthetic.LastStatus -eq 'TimedOut') 'Delayed exit zero became Success.'
        Assert-Case ($script:Released -eq 1 -and $script:Rows.Count -eq 1 -and $script:State.Jobs.Synthetic.PendingRetry.Attempt -eq 1) 'Completion/retry was duplicated or lost.'
    }
    Test-LifecycleCase 'Confirmed immediate kill finalizes existing timeout contract' {
        $script:KillMode='exit'
        Update-RunningJobs $script:Now
        Assert-Case ($script:State.Jobs.Synthetic.LastStatus -eq 'TimedOut' -and $script:Released -eq 1 -and $script:Rows[0].Status -eq 'Retried') 'Confirmed timeout completion changed.'
    }
    Test-LifecycleCase 'Process probe exception keeps supervision and lease' {
        $script:Process.RefreshFails=$true
        Update-RunningJobs $script:Now
        Assert-Case ($script:RunningJobs.ContainsKey('Synthetic') -and $script:Released -eq 0 -and $script:Kills -eq 0) 'Probe failure was treated as a confirmed exit.'
        Assert-Case ($script:LeaseUntil -gt $script:Now.ToUniversalTime()) 'Unknown process state lost lease protection.'
    }
    Test-LifecycleCase 'Post-kill probe exception does not confirm exit' {
        $script:KillMode='probe-error'
        Update-RunningJobs $script:Now
        Assert-Case ($script:RunningJobs.ContainsKey('Synthetic') -and $script:Released -eq 0) 'Post-kill probe error released the job.'
    }
    Test-LifecycleCase 'Wait exception with live process retains ownership' {
        $script:Process.WaitFails=$true
        Update-RunningJobs $script:Now
        Assert-Case ($script:RunningJobs.ContainsKey('Synthetic') -and $script:Released -eq 0) 'Wait exception released a live job.'
    }
    Test-LifecycleCase 'Restart preserves pending timeout and future lease' {
        Update-RunningJobs $script:Now
        Resume-Case
        Assert-Case ($script:RunningJobs.ContainsKey('Synthetic') -and $script:RunningJobs.Synthetic.TimeoutRequested) 'Re-adoption lost pending timeout intent.'
        Assert-Case ($script:LeaseUntil -gt $script:Now.ToUniversalTime()) 'Re-adoption wrote an expired lease.'
        $script:Manifest.JobsByName.Synthetic.TimeoutMinutes=1440
        $script:Process.HasExited=$true
        Update-RunningJobs ($script:Now.AddSeconds(60))
        Assert-Case ($script:State.Jobs.Synthetic.LastStatus -eq 'TimedOut') 'Restart/config reload erased an already requested timeout.'
    }
    Test-LifecycleCase 'Restart after timed-out process disappeared keeps TimedOut' {
        Update-RunningJobs $script:Now
        $script:Adopt=$false
        Resume-Case
        Assert-Case ($script:State.Jobs.Synthetic.LastStatus -eq 'TimedOut') 'Recovered timeout was relabelled Interrupted.'
    }
    Test-LifecycleCase 'Legacy running record is re-adopted without timeout intent' {
        $script:Manifest.JobsByName.Synthetic.TimeoutMinutes=60
        Save-OrchestratorState
        Resume-Case
        Assert-Case ($script:RunningJobs.ContainsKey('Synthetic') -and -not $script:RunningJobs.Synthetic.TimeoutRequested) 'Legacy state compatibility changed.'
    }
    Test-LifecycleCase 'Legacy missing process remains Interrupted' {
        Save-OrchestratorState
        $script:Adopt=$false
        Resume-Case
        Assert-Case ($script:State.Jobs.Synthetic.LastStatus -eq 'Interrupted') 'Legacy missing-process contract changed.'
    }
    Test-LifecycleCase 'Normal exit zero succeeds without timeout' {
        $script:Process.HasExited=$true
        Update-RunningJobs $script:Now
        Assert-Case ($script:State.Jobs.Synthetic.LastStatus -eq 'Success' -and $script:Released -eq 1 -and $null -eq $script:State.Jobs.Synthetic.PendingRetry) 'Normal success contract changed.'
    }
    Test-LifecycleCase 'Normal nonzero exit keeps failure retry contract' {
        $script:Process.HasExited=$true; $script:Process.ExitCode=7
        Update-RunningJobs $script:Now
        Assert-Case ($script:State.Jobs.Synthetic.LastStatus -eq 'Failed' -and $script:State.Jobs.Synthetic.LastExitCode -eq 7 -and $script:State.Jobs.Synthetic.PendingRetry.Attempt -eq 1) 'Normal failure contract changed.'
    }
    Test-LifecycleCase 'Unexpired job keeps original deadline and receives no kill' {
        $script:Manifest.JobsByName.Synthetic.TimeoutMinutes=60
        Update-RunningJobs $script:Now
        $expected=$script:RunningJobs.Synthetic.StartTime.ToUniversalTime().AddMinutes(65)
        Assert-Case ($script:Kills -eq 0 -and $script:LeaseUntil -eq $expected -and $script:Rows.Count -eq 0) 'Unexpired timeout or lease deadline changed.'
    }
    Test-LifecycleCase 'Timeout remains pending after configured deadline increases' {
        Update-RunningJobs $script:Now
        $script:Manifest.JobsByName.Synthetic.TimeoutMinutes=1440
        Update-RunningJobs ($script:Now.AddSeconds(60))
        Assert-Case ($script:Kills -eq 2 -and $script:Released -eq 0) 'Hot reload cancelled a termination already requested.'
    }
    Test-LifecycleCase 'Unavailable state persistence defers kill and preserves supervision' {
        $script:StatePersistenceHealthy=$false
        Update-RunningJobs $script:Now
        Assert-Case ($script:Kills -eq 0 -and $script:Released -eq 0 -and $script:RunningJobs.ContainsKey('Synthetic')) 'Termination proceeded without durable timeout state.'
        Assert-Case ($script:LeaseUntil -gt $script:Now.ToUniversalTime()) 'Persistence failure lost the active lease.'
        $script:StatePersistenceHealthy=$true
        Update-RunningJobs ($script:Now.AddSeconds(60))
        Assert-Case ($script:Kills -eq 1) 'Termination did not resume after persistence recovery.'
    }
    Test-LifecycleCase 'Synthetic file lease rejects a competing writer until confirmed exit' {
        $distributedPath=Join-Path (Split-Path $script:OrchestratorSource -Parent) 'SmartM365.Orchestrator.Distributed.psm1'
        $dt=$null; $de=$null
        $distributedAst=[Management.Automation.Language.Parser]::ParseFile($distributedPath,[ref]$dt,[ref]$de)
        if($de.Count){throw 'Distributed module parse failed.'}
        $leaseDefinitions=foreach($functionName in @('ConvertTo-SafeFileName','Write-JsonAtomically','Enter-SmartM365OrchestratorConcurrencyLease','Set-SmartM365OrchestratorConcurrencyLease','Exit-SmartM365OrchestratorConcurrencyLease')){
            $fn=$distributedAst.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $functionName},$true)
            if(-not $fn){throw "Missing lease function $functionName"}
            $fn.Extent.Text
        }
        $script:LeaseModule=New-Module -ScriptBlock ([scriptblock]::Create($leaseDefinitions -join "`n"))
        $fixtureRoot=Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-Lifecycle-'+[guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($fixtureRoot)
        try {
            $lease=& $script:LeaseModule {param($root,$now) Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $root -ConcurrencyKey synthetic-shared -JobName Synthetic -Occurrence $now -OwnerServer SYNTHETIC-SERVER} $fixtureRoot $script:Now
            Assert-Case $lease.Acquired 'Synthetic fixture lease acquisition failed.'
            & $script:LeaseModule {param($p,$id,$now) Set-SmartM365OrchestratorConcurrencyLease -LeasePath $p -LeaseId $id -OwnerServer SYNTHETIC-SERVER -SafeUntilUtc $now.AddMinutes(-20)} $lease.LeasePath $lease.Lease.LeaseId $script:Now | Out-Null
            $script:RunningJobs.Synthetic.ConcurrencyLeasePath=$lease.LeasePath
            $script:RunningJobs.Synthetic.ConcurrencyLeaseId=$lease.Lease.LeaseId
            function script:Set-SmartM365OrchestratorConcurrencyLease {
                param($LeasePath,$LeaseId,$OwnerServer,$SafeUntilUtc)
                & $script:LeaseModule {param($p,$id,$until) Set-SmartM365OrchestratorConcurrencyLease -LeasePath $p -LeaseId $id -OwnerServer SYNTHETIC-SERVER -SafeUntilUtc $until} $LeasePath $LeaseId $SafeUntilUtc
            }
            function script:Exit-SmartM365OrchestratorConcurrencyLease {
                param($LeasePath,$LeaseId,$OwnerServer)
                & $script:LeaseModule {param($p,$id) Exit-SmartM365OrchestratorConcurrencyLease -LeasePath $p -LeaseId $id -OwnerServer SYNTHETIC-SERVER} $LeasePath $LeaseId
            }
            Update-RunningJobs $script:Now
            $competitor=& $script:LeaseModule {param($root,$now) Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $root -ConcurrencyKey synthetic-shared -JobName Sibling -Occurrence $now -OwnerServer SYNTHETIC-SERVER} $fixtureRoot $script:Now
            Assert-Case (-not $competitor.Acquired) 'A competing writer acquired the overdue live job lease.'
            $script:Process.HasExited=$true
            Update-RunningJobs ($script:Now.AddSeconds(60))
            $afterExit=& $script:LeaseModule {param($root,$now) Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $root -ConcurrencyKey synthetic-shared -JobName Sibling -Occurrence $now -OwnerServer SYNTHETIC-SERVER} $fixtureRoot $script:Now
            Assert-Case $afterExit.Acquired 'Lease was not released after confirmed exit.'
        }
        finally {
            Remove-Module $script:LeaseModule -Force
            $resolvedFixture=[IO.Path]::GetFullPath($fixtureRoot)
            $expectedParent=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
            if(-not $resolvedFixture.StartsWith($expectedParent,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedFixture) -notlike 'SmartInventory-Lifecycle-*'){throw 'Unsafe fixture cleanup path.'}
            Remove-Item -LiteralPath $resolvedFixture -Recurse -Force
        }
    }
}
finally { Remove-Module $module -Force }
$failed=@($results | Where-Object {-not $_.Passed})
$report=[pscustomobject]@{PowerShell=$PSVersionTable.PSVersion.ToString();Source=$source;Total=$results.Count;Passed=$results.Count-$failed.Count;Failed=$failed.Count;Cases=$results.ToArray()}
if($ResultPath){$report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ResultPath -Encoding UTF8}
$report | Select-Object PowerShell,Total,Passed,Failed | Format-Table -AutoSize
$failed | Format-Table Name,Error -Wrap
if($failed.Count){exit 1}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBVo5vezQgLhhp4
# GN5V2lR2Q2F3KP4nGtoPu8DuQQ583qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDRXZD2Ra1A0LLvTrxaevVN
# JyzRVrMRks1bE5h5F4pjjTANBgkqhkiG9w0BAQEFAASCAYCX2/HBA1hcKIHP7xgE
# RIhG97auOnPKIcSQubM7ACjQ1fj9z8lGCfMkaN2x9q/f7Po/4TwiWgAgJO5H+KcA
# 0P1fr8+e9eebqFQMB5LljGJTRCU3MBElrYObHMG4sZP3p2d6/iLoS3HQko+sGliC
# ghwgeNVACKscKzE7d98olgP+7zMn3LpK9SgmbVULTD649pbfEQManxY1+BBYAO1j
# xXlYzBu4cQv43Bwk9DnKCOwIoIf40SPDpfQJ1HORZyaU4JNX15luVxsvnG/1b4ns
# OFHoJMiC6lO7xV12mutI3AYs/kPUZHfI29PG8WN3/T3B9ZskXDVcxINmp1G9EXQ8
# ESH/G8P7LuSYJLemV2dtDN3jIYQrG3gWLEfV8v04uL4ZYTxtRTeKjzbRAauWkCkO
# yViYAspR0knn6qr4gZa1AZA41QreThpJzEtzM3u676GMrhlk01WmaT0sa5xqt/Dd
# an46HGP2gtwH41yezhI0CqLmiipw1SPnhgN8UKxzaV/poGs=
# SIG # End signature block
