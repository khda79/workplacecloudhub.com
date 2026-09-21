<#
.SYNOPSIS
Synthetic lifecycle regression tests; no orchestrator entry point is executed.
.VERSION
1.0.1
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
    Test-LifecycleCase 'Exit code three completes with warnings without retry' {
        $script:Process.HasExited=$true; $script:Process.ExitCode=3
        Update-RunningJobs $script:Now
        Assert-Case ($script:State.Jobs.Synthetic.LastStatus -eq 'CompletedWithWarnings' -and $script:State.Jobs.Synthetic.LastExitCode -eq 3 -and $null -eq $script:State.Jobs.Synthetic.PendingRetry) 'Exit code three was not retained as a non-retrying warning result.'
    }
    Test-LifecycleCase 'Launcher preserves native success warning and failure exit codes' {
        $orchestratorText=[IO.File]::ReadAllText($script:OrchestratorSource)
        $launcherContract='`$commandSucceeded = `$?; `$nativeExitCode = `$LASTEXITCODE; '
        Assert-Case ($orchestratorText.Contains($launcherContract)) 'Launcher child command does not capture native exit state before classification.'
        foreach($expectedExitCode in @(0,3,1)){
            & $env:ComSpec /d /c ("exit /b {0}" -f $expectedExitCode)
            $commandSucceeded=$?
            $nativeExitCode=$LASTEXITCODE
            $forwardedExitCode=if($null -ne $nativeExitCode){[int]$nativeExitCode}elseif(-not $commandSucceeded){1}else{0}
            Assert-Case ($forwardedExitCode -eq $expectedExitCode) ("Launcher exit code {0} was converted to {1}." -f $expectedExitCode,$forwardedExitCode)
        }
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBvfv3JON3kiiNS
# wlOcZ+aeWHy4Fgt1hnMPyXJ7eYOvWqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFtPdYIBM8CiUhJOhx5OpGq5TajZ07tGw1O02LsRIHOnMA0GCSqG
# SIb3DQEBAQUABIIBgI5hYbx+zzC4XYtlMW0ISfJYVHBr1ZFAu2VGYOGvtSrbEpRi
# leG6XFgYP2z1yRmoOqTcllUsxTBNMqnJM/XFTFsveI1Y1ZHxR6+/JUed8QvM3htd
# WSfJMTcUoEVeKqX6ZOeTFqA6vXo5p/K/lTCM2VSUDIPlNhxyByXgqVHUxdDKOgVP
# RLfHFTbJAyO9ZDJ2/OU9oIzSBlA450+hBJc7Of5kJ7yOu+fJAXTeA1fyBHXZpVNV
# 82RRdNqtf5BnsFkkqmXrLH3hvfIWEZ2jzWu11mgy0CJ1bMtae1PkNQM0tXZe7moi
# e9Nf4v4OiFQPpjqLkTbJNOsFrhlR93O1bV8e/Nqd3IF6UbaghUVURTBc4gl0bC2m
# tMOHxI3TzszuH1tyorvSE+W2wiDzSxXDCRiUW1GunS4UedMSywLEZTLebxcnY9g2
# oSLB/U5jbMiWUXHEKT6EGbCq4+wMCi4558up9Q43XwhPYpM+sHmSjPjFw2ChI/rj
# r0CiuZokf5eMlR0xQ6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjExOTE1
# NTlaMC8GCSqGSIb3DQEJBDEiBCDTU1NO8IA8yCAg3v8PeAXSfvrrtIlF3fVz57LF
# eNZYDTANBgkqhkiG9w0BAQEFAASCAgBUcOwjaJEIHG9H4HPs0rBKYqVE6M7/Almf
# 6K/8br3LKD45wg0X2jQ6AitvyBKaPvQyBCjCbV7IVlMqP/7McjmF9KBSb/0OAORw
# 8i6Gc5aRpECQPp2ydbEU8AsWFpme9BXdzMnN6Pc4YlyIt9HmSkeSQuylOVazmB6O
# DFNeypwbn+FvGyrp49B1qJ7L/hf+Xz4pHs4gJU0jmGxCe0+SLXqO2DYFmZ+5A1Q4
# OdBB5A1mcbn0lfp6JthWmOUuRFaKyLniTy64Bze5lDHKE9txGwR4xx4jawNK8eqh
# Aw6E9oFMVeDCst116TPLNQGkzKKM1K6Wog/IFPZrlm+PejRtYiPQg9CywAkB1Rld
# wECS4WfUgbDV6Nf6/rzpcn51Ws2SjjaPIpKLTiUHp8vaMk9TqP+KmaWSA5p0SJMB
# RIhq3uT+73VlubMNu+lEnA6ugvvqpIrP4GXX5nkFCt7IXVgdBtBaqk8K6iMCDwi1
# 6D72c72nggAW2uPIvoXKWURL9Ojs0delm5sr3gVIJoSez4w4os4W6+VCAIHfSDGS
# OSBtVwNUXtx36UUaMRs8iz+BSGMV4BY01M10YLT1sXnVD03K8FtFvEivT5SJ+jPk
# teSFOcWGWwHOrM4pm6bIzmcIKxuYSYkwQBOd6UySYfkUrQhkYh8hejn2HF4VtTqv
# ejwgi/KW1g==
# SIG # End signature block
