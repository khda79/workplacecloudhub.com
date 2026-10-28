#Requires -Version 7.0
<#
.SYNOPSIS
Offline maintenance, scheduler admission and GUI contract tests using synthetic shared data.
.VERSION
1.0.8
#>
[CmdletBinding()]
param([string]$Tenant = 'test')
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Maintenance-Test-' + [guid]::NewGuid().ToString('N'))
$orchFolder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $orchFolder 'SmartM365.Orchestrator.Maintenance.psm1') -Force
$script:Cases = 0
$runtime = $null
function Assert-Case { param([bool]$Condition,[string]$Message) $script:Cases++; if (-not $Condition) { throw $Message } }
function Assert-Throws { param([scriptblock]$Body,[string]$Pattern) $failed=$false; try { & $Body | Out-Null } catch { if ($_.Exception.Message -notlike $Pattern) { throw }; $failed=$true }; Assert-Case $failed "Expected refusal: $Pattern" }
function Write-Fixture { param([string]$Path,$Data) $null=New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force; $null=Write-SmartM365JsonBytesAtomically -Path $Path -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($Data | ConvertTo-Json -Depth 20))) -Validate { param($d) if ($d -isnot [pscustomobject]) { throw 'Fixture must be an object.' } } }
function Write-Heartbeat {
    param([string]$Server,$Control,[int]$Protocol=1,[int]$AgeMinutes=0,[bool]$Healthy=$true,[string]$Lifecycle='Running',[bool]$CatchUpSupported=$false)
    Write-Fixture (Join-Path $root "$Server/Orchestrator-Heartbeat.json.txt") ([pscustomobject]@{
        ScriptVersion='1.5.39'; Timestamp=[datetime]::UtcNow.AddMinutes(-$AgeMinutes).ToString('o')
        Lifecycle=$Lifecycle; MaintenanceProtocol=$Protocol; MaintenanceCatchUpProtocol=if($CatchUpSupported){1}else{0}; MaintenanceHealthy=$Healthy
        MaintenanceRevision=$Control.Revision; MaintenanceEnabled=$Control.Enabled
        RunningJobs=@([pscustomobject]@{Name='Synthetic';Pid=999999})
    })
}
function Load-Functions { param([string]$Path,[string[]]$Names) $t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$t,[ref]$e); if($e.Count){throw ($e.Message -join '; ')}; foreach($name in $Names){$n=$ast.Find({param($x) $x -is [Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $name},$true); if(-not $n){throw "Missing $name"}; $n.Extent.Text} }
try {
    $null=New-Item -ItemType Directory -Path (Join-Path $root 'Config') -Force
    $cluster=[pscustomobject]@{ExpectedOrchestratorServers=@('SERVER-A','SERVER-B');PeerHeartbeatStaleMinutes=5}
    Write-Fixture (Join-Path $root 'Config/Orchestrator-Cluster.json.txt') $cluster
    Write-Fixture (Join-Path $root 'Config/Orchestrator-Jobs.json.txt') ([pscustomobject]@{Jobs=@(
        [pscustomobject]@{Name='Included';Enabled=$true;AssignmentMode='Elected';Schedule=[pscustomobject]@{Type='Daily';Times=@('00:05','00:10');DaysOfWeek=@()}},
        [pscustomobject]@{Name='Excluded';Enabled=$false;AssignmentMode='Elected';Schedule=[pscustomobject]@{Type='Daily';Times=@('00:05');DaysOfWeek=@()}}
    )})
    $clusterHash=(Get-FileHash (Join-Path $root 'Config/Orchestrator-Cluster.json.txt')).Hash
    $zero=Get-SmartM365OrchestratorMaintenanceState $root
    Assert-Case (-not $zero.Enabled -and $zero.Revision -eq 0) 'A new deployment is not initially inactive.'
    Assert-Case (-not (Test-Path (Join-Path $root 'Config/Orchestrator-Maintenance.guard'))) 'Read-only getter created a gate.'
    $zero=Initialize-SmartM365OrchestratorMaintenance $root
    $paths=Get-SmartM365OrchestratorMaintenancePaths $root
    Assert-Case ($paths.State.EndsWith('.json.txt') -and (Test-Path $paths.State)) 'State is not OneDrive-compatible.'
    Assert-Case ((Get-Item $paths.Gate).Length -gt 0) 'Initialization marker is missing.'
    Assert-Case ((Initialize-SmartM365OrchestratorMaintenance $root).Revision -eq 0) 'Initialization is not idempotent.'
    $legacy=Join-Path $root 'Legacy'; $null=New-Item -ItemType Directory -Path $legacy
    Assert-Case ((Initialize-SmartM365OrchestratorMaintenance $legacy).Revision -eq 0) 'Non-central deployment could not initialize its Config folder.'
    Assert-Throws { Set-SmartM365OrchestratorMaintenance $root $true 0 ' ' } '*reason*'
    Assert-Throws { Set-SmartM365OrchestratorMaintenance $root $true 0 'offline qualification' } '*Every expected server*'
    Write-Heartbeat 'SERVER-A' $zero
    Write-Heartbeat 'SERVER-B' $zero -Protocol 0
    Assert-Case (@(Get-SmartM365OrchestratorMaintenanceReadiness $root $cluster $zero | Where-Object Status -eq 'Unsupported version').Count -eq 1) 'Old protocol is not visible.'
    Assert-Throws { Set-SmartM365OrchestratorMaintenance $root $true 0 'offline qualification' } '*Every expected server*'
    Write-Heartbeat 'SERVER-B' $zero -AgeMinutes 10 -Protocol 0
    Assert-Throws { Set-SmartM365OrchestratorMaintenance $root $true 0 'offline qualification' } '*Every expected server*'
    # An offline peer whose last heartbeat advertises protocol 1 applies maintenance when it returns.
    Write-Heartbeat 'SERVER-B' $zero -AgeMinutes 10
    $offlineReadiness=@(Get-SmartM365OrchestratorMaintenanceReadiness $root $cluster $zero)
    Assert-Case (@($offlineReadiness | Where-Object { $_.Server -eq 'SERVER-B' -and $_.Status -eq 'Offline (applies on return)' }).Count -eq 1) 'A compatible offline peer is not reported as applying on return.'
    Assert-Case (Test-SmartM365OrchestratorMaintenanceActivationReady -Servers $offlineReadiness) 'A compatible offline peer blocks activation.'
    Assert-Case (-not (Test-SmartM365OrchestratorMaintenanceActivationReady -Servers @([pscustomobject]@{Status='Offline (applies on return)'},[pscustomobject]@{Status='Offline (applies on return)'}))) 'Activation was allowed with no online server.'
    Assert-Case (-not (Test-SmartM365OrchestratorMaintenanceActivationReady -Servers @([pscustomobject]@{Status='Applied'},[pscustomobject]@{Status='Offline'}))) 'Activation was allowed with an unreadable offline peer.'
    Assert-Case (-not (Test-SmartM365OrchestratorMaintenanceActivationReady -Servers @())) 'Activation was allowed without expected servers.'
    Write-Heartbeat 'SERVER-B' $zero -AgeMinutes -10
    Assert-Case (@(Get-SmartM365OrchestratorMaintenanceReadiness $root $cluster $zero | Where-Object { $_.Server -eq 'SERVER-B' -and $_.Status -eq 'Offline' }).Count -eq 1) 'A future heartbeat was accepted as offline-compatible.'
    # Real activation with a compatible offline peer, in an isolated copy of the shared root.
    $mainRoot=$root; $root=Join-Path $mainRoot 'OfflinePeer'
    try {
        $null=New-Item -ItemType Directory -Path (Join-Path $root 'Config') -Force
        Write-Fixture (Join-Path $root 'Config/Orchestrator-Cluster.json.txt') $cluster
        $offlineZero=Initialize-SmartM365OrchestratorMaintenance $root
        Write-Heartbeat 'SERVER-A' $offlineZero; Write-Heartbeat 'SERVER-B' $offlineZero -AgeMinutes 30
        $offlineActive=Set-SmartM365OrchestratorMaintenance $root $true 0 'Synthetic pause with an offline peer'
        Assert-Case ($offlineActive.Enabled -and $offlineActive.Revision -eq 1) 'Activation with a compatible offline peer was not published.'
        Write-Heartbeat 'SERVER-A' $offlineActive -CatchUpSupported $true
        Assert-Throws { Set-SmartM365OrchestratorMaintenance $root $false 1 'Catch-up with an offline peer' -CatchUpMissedOccurrences } '*Every expected server*support catch-up*'
    }
    finally { $root=$mainRoot }
    Write-Heartbeat 'SERVER-B' $zero -Healthy $false
    Assert-Throws { Set-SmartM365OrchestratorMaintenance $root $true 0 'offline qualification' } '*Every expected server*'
    Write-Heartbeat 'SERVER-B' $zero -Lifecycle Starting
    Assert-Throws { Set-SmartM365OrchestratorMaintenance $root $true 0 'offline qualification' } '*Every expected server*'
    Write-Heartbeat 'SERVER-B' $zero
    $active=Set-SmartM365OrchestratorMaintenance $root $true 0 'Synthetic maintenance'
    Assert-Case ($active.Enabled -and $active.Revision -eq 1) 'Activation was not published.'
    Assert-Case ((Get-SmartM365OrchestratorMaintenanceState $root).Enabled) 'Activation does not survive a fresh read.'
    Assert-Case (@(Get-SmartM365OrchestratorMaintenanceReadiness $root $cluster $active | Where-Object Status -eq 'Pending').Count -eq 2) 'Unacknowledged activation was marked applied.'
    Assert-Throws { Set-SmartM365OrchestratorMaintenance $root $false 0 'Stale GUI' } '*another session*'
    foreach($origin in @('due','retry')) { Assert-Case (-not (Test-SmartM365OrchestratorMaintenanceLaunch $active $origin (Get-Date))) "$origin bypassed maintenance." }
    foreach($origin in @('pipeline','pipeline-retry','forced')) { Assert-Case (Test-SmartM365OrchestratorMaintenanceLaunch $active $origin (Get-Date)) "$origin was blocked despite explicit request." }
    Write-Heartbeat 'SERVER-A' $active; Write-Heartbeat 'SERVER-B' $active
    Assert-Case (@(Get-SmartM365OrchestratorMaintenanceReadiness $root $cluster $active | Where-Object Status -eq 'Applied').Count -eq 2) 'Both acknowledgements were not recognized.'
    Assert-Case ((Initialize-SmartM365OrchestratorMaintenance $root).Enabled) 'Restart initialization cleared maintenance.'
    $gate=Enter-SmartM365OrchestratorMaintenanceGate $root
    try { Assert-Throws { Enter-SmartM365OrchestratorMaintenanceGate $root -TimeoutSeconds 0 } '*' }
    finally { $gate.Dispose() }
    # Disable must remain possible when a peer goes offline; acknowledgement stays pending.
    Write-Heartbeat 'SERVER-B' $active -AgeMinutes 10
    $resumed=Set-SmartM365OrchestratorMaintenance $root $false 1 'Synthetic resume'
    Assert-Case (-not $resumed.Enabled -and $resumed.Revision -eq 2 -and $resumed.ResumeAfterUtc) 'Resume has no durable cutoff.'
    $cutoff=([datetimeoffset]::Parse($resumed.ResumeAfterUtc)).LocalDateTime
    Assert-Case (-not (Test-SmartM365OrchestratorMaintenanceLaunch $resumed 'due' $cutoff)) 'Exact cutoff was caught up.'
    Assert-Case (-not (Test-SmartM365OrchestratorMaintenanceLaunch $resumed 'retry' $cutoff.AddDays(-1))) 'Old retry was caught up.'
    Assert-Case (Test-SmartM365OrchestratorMaintenanceLaunch $resumed 'due' $cutoff.AddSeconds(1)) 'Future schedule was blocked.'
    Assert-Case (Test-SmartM365OrchestratorMaintenanceLaunch $resumed 'pipeline' $cutoff.AddDays(-1)) 'Pending manual request was discarded.'
    $audit=@(Import-Csv $paths.Audit)
    Assert-Case ($audit.Count -eq 4 -and @($audit | Where-Object Outcome -eq Published).Count -eq 2) 'Transitions were not audited.'
    Assert-Case ((Get-FileHash (Join-Path $root 'Config/Orchestrator-Cluster.json.txt')).Hash -eq $clusterHash) 'Maintenance modified cluster configuration.'
    # Real runtime functions, isolated state and no process/tenant connections.
    $definitions=Load-Functions (Join-Path $orchFolder 'SmartM365-Inventory-Orchestrator.ps1') @('Update-OrchestratorMaintenance','Get-OrchestratorPipelineWaitStartUtc','Get-OrchestratorDependencyActivityText','Get-OrchestratorPipelineDependencyWaitText','Set-OrchestratorPipelineDependencyWait','ConvertTo-StateTime','ConvertFrom-StateTime','Get-JobState','Get-LatestPastOccurrence','Get-JobOccurrencesInWindow','Get-DueOccurrence','Get-MaintenanceCatchUpOccurrence','Complete-MaintenanceCatchUpOccurrence','Invoke-LaunchPhase','Test-OrchestratorLocalConcurrencyBlocked','Get-OrchestratorPendingJobSnapshot')
    $runtime=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
    $runtimeResult=& $runtime {
        param($Root,$Cutoff)
        $script:Settings=@{SharedDataFolderPath=$Root;MaxConcurrency=4}
        $script:State=@{Jobs=@{}}; $script:RunningJobs=@{Synthetic=@{Pid=999999}}
        $job=[pscustomobject]@{Name='Synthetic';Enabled=$true;AssignmentMode='Manual';Schedule=[pscustomobject]@{Type='Daily';Times=@('00:00');DaysOfWeek=@();MissedRunPolicy='RunOnce'};DependsOn=@();ConcurrencyKey='Synthetic'}
        $job2=$job.PSObject.Copy(); $job2.Name='ManualRetry'
        $script:Manifest=@{OrderedJobs=@($job,$job2)}
        $script:MaintenanceControl=$null;$script:MaintenanceHealthy=$false;$script:MaintenanceError='';$script:StatePersistenceHealthy=$true
        $script:Saved=0;$script:Starts=0;$script:ForcedPending=@();$script:PipelinePending=@{};$script:PipelineBatchJobNames=@{};$script:DependencyWaitReason=@{};$script:PipelineWaitState=@{}
        function script:Write-OrchestratorLog { param($Message,$Level) }
        function script:Invoke-SmartM365OrchestratorMaintenanceNotification { param($SharedDataFolderPath,$Now,$SendAction,[switch]$Initialize) [pscustomobject]@{Status='Completed';Sent=0} }
        function script:Write-OrchestratorRuntimeUpdateWarning { param($Key,$Message,$Now,$Level) }
        function script:Save-OrchestratorState { $script:Saved++ }
        function script:Test-JobSelected {param($JobName) $true}
        function script:Test-JobAllowedOnServer {param($Job,[switch]$AllowManual) $true}
        # Launch prerequisites have their own synthetic suite; this one assumes they are met.
        function script:Test-OrchestratorLaunchPrerequisiteGate { param($Job,$Occurrence,$Reason,$Attempt,[switch]$IsForced,$PipelineInfo,$Now) $true }
        function script:Start-InventoryJob {param($Job,$Occurrence,$Attempt,$ClaimPath,$ConcurrencyLeasePath,$ConcurrencyLeaseId,$PipelineBatchId,$PipelineStatusPath) $script:Starts++}
        $s=Get-JobState Synthetic; $s.LastStatus='Failed';$s.LastScheduledOccurrence=$Cutoff.AddDays(-3).ToString('o');$s.Running=@{Pid=999999};$s.PendingRetry=@{ScheduledOccurrence=$Cutoff.AddDays(-2).ToString('o');Attempt=1;NotBefore=$Cutoff.AddDays(-1).ToString('o')}
        $m=Get-JobState ManualRetry;$m.PendingRetry=@{PipelineBatchId='synthetic-batch';ScheduledOccurrence=$Cutoff.AddDays(-1).ToString('o');Attempt=1;NotBefore=$Cutoff.AddDays(-1).ToString('o')}
        Update-OrchestratorMaintenance -Now $Cutoff.AddMinutes(10)
        [pscustomobject]@{Healthy=$script:MaintenanceHealthy;Saved=$script:Saved;LastStatus=$s.LastStatus;Running=$s.Running.Pid;Retry=$s.PendingRetry;ManualRetry=$m.PendingRetry;Cursor=$s.LastScheduledOccurrence}
    } $root $cutoff
    Assert-Case ($runtimeResult.Healthy -and $runtimeResult.Saved -gt 0) 'Resume was not persisted by runtime.'
    Assert-Case ($runtimeResult.LastStatus -eq 'Failed' -and $runtimeResult.Running -eq 999999) 'Runtime erased failure or running process.'
    Assert-Case ($null -eq $runtimeResult.Retry -and $null -ne $runtimeResult.ManualRetry) 'Retry origins were not distinguished.'
    Assert-Case ([datetime]::Parse($runtimeResult.Cursor) -le $cutoff -and [datetime]::Parse($runtimeResult.Cursor) -gt $cutoff.AddDays(-3)) 'Resume cursor is not bounded by cutoff.'
    Write-Heartbeat 'SERVER-A' $resumed; Write-Heartbeat 'SERVER-B' $resumed
    $again=Set-SmartM365OrchestratorMaintenance $root $true 2 'Synthetic second pause'
    Import-Module (Join-Path $orchFolder 'SmartM365.Orchestrator.Insights.psm1') -Force
    $operations=Get-SmartM365OrchestratorOperations -SharedDataFolderPath $root -ClusterDocument $cluster
    Assert-Case ($operations.Maintenance.Enabled -and -not $operations.MaintenanceError) 'Operational view lost the control state.'
    Assert-Case (@($operations.Servers | Where-Object { $_.Maintenance -eq 'ON' -and $_.MaintenanceAcknowledgement -eq 'Pending' }).Count -eq 2) 'Operational view claimed unacknowledged activation.'
    Write-Heartbeat 'SERVER-A' $again; Write-Heartbeat 'SERVER-B' $again
    $operations=Get-SmartM365OrchestratorOperations -SharedDataFolderPath $root -ClusterDocument $cluster
    Assert-Case (@($operations.Servers | Where-Object { $_.Maintenance -eq 'ON' -and $_.MaintenanceAcknowledgement -eq 'Applied' }).Count -eq 2) 'Operational view does not show acknowledgements.'
    # Mirror selects the state/audit, never the live launch gate.
    $mirrorDefinitions=Load-Functions (Join-Path $orchFolder 'SmartM365-Inventory-Orchestrator.ps1') @('Get-OrchestratorSharePointMirrorRelativePath','Test-OrchestratorSharePointMirrorFile','Get-OrchestratorSharePointMirrorSnapshot')
    $mirror=New-Module -ScriptBlock ([scriptblock]::Create(($mirrorDefinitions -join "`n") + "`nfunction Update-OrchestratorHeartbeatDuringLongOperation {}"))
    try {
        $snapshot=& $mirror {param($Root) Get-OrchestratorSharePointMirrorSnapshot $Root} $root
        Assert-Case (@($snapshot.Files | Where-Object LocalFilePath -eq $paths.State).Count -eq 1) 'Maintenance state is not mirrored.'
        Assert-Case (@($snapshot.Files | Where-Object LocalFilePath -eq $paths.Gate).Count -eq 0) 'Live maintenance gate would be mirrored.'
        Assert-Case (@($snapshot.Files | Where-Object LocalFilePath -eq $paths.Audit).Count -eq 1) 'Maintenance audit is not mirrored.'
    } finally { Remove-Module $mirror }
    $launchResults=& $runtime {
        param($Active,$OldControl)
        $script:Settings.ConcurrencyLeasesPath='synthetic';$script:Settings.ElectionClaimGraceMinutes=5;$script:Settings.PeerHeartbeatStaleMinutes=5
        $script:Settings.DependencyWaitTimeoutMinutes=60
        $job=$script:Manifest.OrderedJobs[1];$job.AssignmentMode='Legacy'
        $job | Add-Member NoteProperty TimeoutMinutes 60
        $job | Add-Member NoteProperty DependencyWaitTimeoutMinutes 60
        $script:Manifest=@{OrderedJobs=@($job);JobsByName=@{ManualRetry=$job}}
        $script:RunningJobs=@{};$script:Starts=0;$script:LeaseCalls=0;$script:MaintenanceHealthy=$true
        $script:PipelinePending=@{};$script:MaintenanceControl=$Active
        function script:Clear-DependencyWaitLog {param($JobName)}
        # This suite isolates maintenance; cancellation admission has its own synthetic suite.
        function script:Test-SmartM365OrchestratorPipelineCancellation {param($SharedDataFolderPath,$BatchId) $false}
        function script:Enter-SmartM365OrchestratorPipelineLaunchGuard {param($SharedDataFolderPath,$BatchId,$JobName) [pscustomobject]@{Allowed=$true;Lock=$null;LaunchAttempted=$false}}
        function script:Exit-SmartM365OrchestratorPipelineLaunchGuard {param($Guard)}
        function script:Write-DependencyWaitLog {param($JobName,$BlockingDependencies,$Now)}
        function script:Enter-SmartM365OrchestratorConcurrencyLease {param($LeasesRootPath,$ConcurrencyKey,$JobName,$Occurrence,$OwnerServer,$SafeMinutes,$HeartbeatRootPath,$HeartbeatStaleMinutes) $script:LeaseCalls++; @{Acquired=$true;LeasePath='synthetic';Lease=@{LeaseId='synthetic'}}}
        function script:Get-DueOccurrence {param($Job,$LastOccurrence,$Now) $Now}
        # Launch prerequisites have their own synthetic suite; this one assumes they are met.
        function script:Test-OrchestratorLaunchPrerequisiteGate { param($Job,$Occurrence,$Reason,$Attempt,[switch]$IsForced,$PipelineInfo,$Now) $true }
        function script:Start-InventoryJob {param($Job,$Occurrence,$Attempt,$ClaimPath,$ConcurrencyLeasePath,$ConcurrencyLeaseId,$PipelineBatchId,$PipelineStatusPath) $script:Starts++;$script:LastBatch=$PipelineBatchId;$script:State.Jobs[$Job.Name].PendingRetry=$null; if($script:ThrowStart){throw 'Synthetic launch failure'} }
        $script:ThrowStart=$false
        $s=Get-JobState ManualRetry; $s.PendingRetry=$null
        Invoke-LaunchPhase (Get-Date); $automatic=$script:Starts
        $s.PendingRetry=@{ScheduledOccurrence=(Get-Date).AddDays(-1).ToString('o');NotBefore=(Get-Date).AddMinutes(-1).ToString('o');Attempt=1}
        Invoke-LaunchPhase (Get-Date); $retry=$script:Starts
        # A manual request must not be trapped behind a suspended automatic retry.
        $script:PipelinePending=@{ManualRetry=@{Occurrence=(Get-Date);CreatedAtUtc=[datetime]::UtcNow;Attempt=0;BatchId='synthetic';StatusPath='synthetic'}}
        $script:PipelineBatchJobNames=@{synthetic=[Collections.Generic.HashSet[string]]::new()}
        Invoke-LaunchPhase (Get-Date); $manual=$script:Starts; $batch=$script:LastBatch
        $script:PipelinePending=@{}
        $s.PendingRetry=@{ScheduledOccurrence=(Get-Date).AddDays(-1).ToString('o');NotBefore=(Get-Date).AddMinutes(-1).ToString('o');Attempt=1;PipelineBatchId='synthetic';PipelineStatusPath='synthetic'}
        Invoke-LaunchPhase (Get-Date); $pipelineRetry=$script:Starts
        $job.DependsOn=@('Parent');$script:ParentStatus='Pending'
        $script:Manifest.JobsByName.Parent=[pscustomobject]@{ContinueOnError=$false}
        $script:PipelineBatchJobNames.synthetic.Add('Parent') | Out-Null
        function script:Get-OrchestratorPipelineDependencyStatus {param($BatchId,$JobName) $script:ParentStatus}
        function script:Set-OrchestratorPipelineJobStatus {param($BatchId,$JobName,$Status,$Attempt,$Detail) $script:WaitDetail=$Detail}
        $script:PipelinePending=@{ManualRetry=@{Occurrence=(Get-Date);CreatedAtUtc=[datetime]::UtcNow;Attempt=0;BatchId='synthetic';StatusPath='synthetic'}}
        Invoke-LaunchPhase (Get-Date); $waiting=$script:Starts
        Assert-Case ($script:WaitDetail -eq 'Waiting for 1 dependency: Parent (not finished).') "A waiting request job must record its reason: $script:WaitDetail"
        $script:ParentStatus='Success';Invoke-LaunchPhase (Get-Date);$ready=$script:Starts
        $job.DependsOn=@();$script:PipelinePending=@{};$s.PendingRetry=$null
        # Activation after the tick read must stop the due launch before any shared claim.
        $script:MaintenanceControl=$OldControl; $before=$script:LeaseCalls
        Invoke-LaunchPhase (Get-Date); $raced=$script:Starts; $raceLeases=$script:LeaseCalls - $before
        $script:MaintenanceControl=$Active
        $snapshot=@(Get-OrchestratorPendingJobSnapshot (Get-Date))
        $script:ForcedPending=@('ManualRetry');$script:ThrowStart=$true;$startError=''
        try { Invoke-LaunchPhase (Get-Date) } catch { $startError=$_.Exception.Message }
        [pscustomobject]@{Automatic=$automatic;Retry=$retry;Manual=$manual;Batch=$batch;PipelineRetry=$pipelineRetry;Waiting=$waiting;Ready=$ready;Raced=$raced;RaceLeases=$raceLeases;Snapshot=$snapshot;StartError=$startError}
    } $again $resumed
    Assert-Case ($launchResults.Automatic -eq 0 -and $launchResults.Retry -eq 0) 'Real launch phase started an automatic run/retry.'
    Assert-Case ($launchResults.Manual -eq 1 -and $launchResults.Batch -eq 'synthetic') 'Manual request was trapped behind scheduled retry.'
    Assert-Case ($launchResults.PipelineRetry -eq 2) 'Real launch phase blocked pipeline retry.'
    Assert-Case ($launchResults.Waiting -eq 2 -and $launchResults.Ready -eq 3) 'Maintenance bypassed dependencies or blocked their success.'
    Assert-Case ($launchResults.Raced -eq 3 -and $launchResults.RaceLeases -eq 0) 'Activation race acquired a lease or started a process.'
    Assert-Case ($launchResults.Snapshot.Count -eq 1 -and $launchResults.Snapshot[0].Reason -eq 'Maintenance') 'Snapshot reports overdue instead of deliberate suspension.'
    Assert-Case ($launchResults.StartError -eq 'Synthetic launch failure') 'Launch exception was swallowed by maintenance wrapper.'
    $gate=Enter-SmartM365OrchestratorMaintenanceGate $root -TimeoutSeconds 0; $gate.Dispose()
    Assert-Case $true 'Gate was not released after launch exception.'
    Assert-Throws { Set-SmartM365OrchestratorMaintenance $root $false 3 'Catch-up without upgraded peers' -CatchUpMissedOccurrences } '*Every expected server*support catch-up*'
    Write-Heartbeat 'SERVER-A' $again -CatchUpSupported $true
    Write-Heartbeat 'SERVER-B' $again -CatchUpSupported $true
    $catchUpControl=Set-SmartM365OrchestratorMaintenance $root $false 3 'Catch up all missed occurrences' -CatchUpMissedOccurrences
    Assert-Case ($catchUpControl.CatchUpAll -and $catchUpControl.CatchUpFromUtc -eq $again.ChangedAtUtc -and $catchUpControl.ResumeAfterUtc) 'Catch-up choice or window was not persisted.'
    Assert-Case (@($catchUpControl.CatchUpJobs).Count -eq 1 -and $catchUpControl.CatchUpJobs[0].Name -eq 'Included' -and @($catchUpControl.CatchUpJobs[0].Schedule.Times).Count -eq 2) 'Catch-up did not freeze enabled jobs and their schedule.'
    $catchUpStart=([datetimeoffset]::Parse($catchUpControl.CatchUpFromUtc)).LocalDateTime
    $catchUpEnd=([datetimeoffset]::Parse($catchUpControl.ResumeAfterUtc)).LocalDateTime
    Assert-Case (Test-SmartM365OrchestratorMaintenanceLaunch $catchUpControl 'maintenance-catch-up' $catchUpStart) 'Confirmed catch-up was blocked.'
    Assert-Case (-not (Test-SmartM365OrchestratorMaintenanceLaunch $catchUpControl 'maintenance-catch-up' $catchUpStart.AddSeconds(-1))) 'Occurrence before maintenance was admitted.'
    Assert-Case (-not (Test-SmartM365OrchestratorMaintenanceLaunch $catchUpControl 'maintenance-catch-up' $catchUpEnd.AddSeconds(1))) 'Occurrence after resume was admitted as catch-up.'
    Assert-Case (-not (Test-SmartM365OrchestratorMaintenanceLaunch $resumed 'maintenance-catch-up' $cutoff)) 'Unconfirmed catch-up was admitted.'
    Write-Heartbeat 'SERVER-A' $catchUpControl; Write-Heartbeat 'SERVER-B' $catchUpControl
    $pausedAgain=Set-SmartM365OrchestratorMaintenance $root $true 4 'Synthetic pause before catch-up completed'
    Write-Heartbeat 'SERVER-A' $pausedAgain; Write-Heartbeat 'SERVER-B' $pausedAgain
    $resumedWithoutNewQueue=Set-SmartM365OrchestratorMaintenance $root $false 5 'Preserve earlier catch-up only'
    Assert-Case (-not $resumedWithoutNewQueue.CatchUpAll -and @($resumedWithoutNewQueue.CatchUpWindows).Count -eq 1) 'A later maintenance pause or No choice erased earlier catch-up work.'
    Write-Heartbeat 'SERVER-A' $resumedWithoutNewQueue; Write-Heartbeat 'SERVER-B' $resumedWithoutNewQueue
    $nextPause=Set-SmartM365OrchestratorMaintenance $root $true 6 'Synthetic third pause'
    Write-Heartbeat 'SERVER-A' $nextPause -CatchUpSupported $true; Write-Heartbeat 'SERVER-B' $nextPause -CatchUpSupported $true
    $nextCatchUp=Set-SmartM365OrchestratorMaintenance $root $false 7 'Append another confirmed catch-up' -CatchUpMissedOccurrences
    Assert-Case (@($nextCatchUp.CatchUpWindows).Count -eq 2 -and $nextCatchUp.CatchUpWindows[0].Revision -eq 4 -and $nextCatchUp.CatchUpWindows[1].Revision -eq 8) 'A subsequent Yes did not append a distinct durable catch-up window.'

    $catchUpDefinitions=Load-Functions (Join-Path $orchFolder 'SmartM365-Inventory-Orchestrator.ps1') @('ConvertTo-StateTime','ConvertFrom-StateTime','Get-JobOccurrencesInWindow','Get-MaintenanceCatchUpOccurrence','Complete-MaintenanceCatchUpOccurrence')
    $catchUpRuntime=New-Module -ScriptBlock ([scriptblock]::Create($catchUpDefinitions -join "`n"))
    try {
        $queue=& $catchUpRuntime {
            $script:Saved=0
            function script:Save-OrchestratorState { $script:Saved++ }
            $from=[datetime]::Today.AddMinutes(1)
            $through=$from.AddMinutes(18)
            $snapshot=[pscustomobject]@{Name='Repeated';Schedule=[pscustomobject]@{Type='Daily';Times=@('00:05','00:10','00:15');DaysOfWeek=@()}}
            $control=[pscustomobject]@{Enabled=$false;CatchUpAll=$true;CatchUpFromUtc=$from.ToUniversalTime().ToString('o');ResumeAfterUtc=$through.ToUniversalTime().ToString('o');Revision=20;CatchUpJobs=@($snapshot);CatchUpWindows=@([pscustomobject]@{Revision=20;FromUtc=$from.ToUniversalTime().ToString('o');ThroughUtc=$through.ToUniversalTime().ToString('o');Jobs=@($snapshot)})}
            $job=[pscustomobject]@{Name='Repeated';Enabled=$true;Schedule=[pscustomobject]@{Type='Daily';Times=@('00:05','00:10','00:15');DaysOfWeek=@();MissedRunPolicy='Skip'}}
            $laterEnabled=[pscustomobject]@{Name='Later';Enabled=$true;Schedule=$job.Schedule}
            $unplanned=Get-MaintenanceCatchUpOccurrence $laterEnabled @{} $control -ReadOnly
            $job.Schedule.Times=@('00:30')
            $state=@{}
            $readOnly=Get-MaintenanceCatchUpOccurrence $job $state $control -ReadOnly
            $first=Get-MaintenanceCatchUpOccurrence $job $state $control
            Complete-MaintenanceCatchUpOccurrence $state $first
            $second=Get-MaintenanceCatchUpOccurrence $job $state $control
            Complete-MaintenanceCatchUpOccurrence $state $second
            $third=Get-MaintenanceCatchUpOccurrence $job $state $control
            Complete-MaintenanceCatchUpOccurrence $state $third
            $done=Get-MaintenanceCatchUpOccurrence $job $state $control
            $again=Get-MaintenanceCatchUpOccurrence $job $state $control
            $nextFrom=$from.AddDays(1); $nextThrough=$through.AddDays(1)
            $control.CatchUpWindows+=@([pscustomobject]@{Revision=22;FromUtc=$nextFrom.ToUniversalTime().ToString('o');ThroughUtc=$nextThrough.ToUniversalTime().ToString('o');Jobs=@($snapshot)})
            $next=Get-MaintenanceCatchUpOccurrence $job $state $control
            [pscustomobject]@{ReadOnly=$readOnly;First=$first;Second=$second;Third=$third;Done=$done;Again=$again;Next=$next;Unplanned=$unplanned;Saved=$script:Saved;Cursor=$state.MaintenanceCatchUpLastOccurrence;Revision=$state.MaintenanceCatchUpRevision}
        }
        Assert-Case ($queue.ReadOnly -eq [datetime]::Today.AddMinutes(5) -and $queue.First -eq [datetime]::Today.AddMinutes(5) -and $queue.Second -eq [datetime]::Today.AddMinutes(10) -and $queue.Third -eq [datetime]::Today.AddMinutes(15)) 'Catch-up did not preserve every occurrence in order.'
        Assert-Case ($null -eq $queue.Unplanned) 'A job enabled after confirmation entered an old catch-up queue.'
        Assert-Case ($queue.Next -eq [datetime]::Today.AddDays(1).AddMinutes(5) -and $queue.Revision -eq 22) 'A new confirmed maintenance window did not follow the earlier completed window.'
        Assert-Case ($null -eq $queue.Done -and $null -eq $queue.Again -and $queue.Saved -ge 4) 'Catch-up was not durably completed once.'
    } finally { Remove-Module $catchUpRuntime }
    # Invalid state cannot be interpreted as disabled, and deletion cannot reset it.
    $goodBytes=[IO.File]::ReadAllBytes($paths.State)
    $invalid=$again.PSObject.Copy();$invalid.Enabled='false'
    Write-Fixture $paths.State $invalid
    Assert-Throws { Get-SmartM365OrchestratorMaintenanceState $root } '*invalid*'
    $invalid=$again.PSObject.Copy();$invalid.Revision=-1
    Write-Fixture $paths.State $invalid
    Assert-Throws { Get-SmartM365OrchestratorMaintenanceState $root } '*invalid*'
    [IO.File]::WriteAllText($paths.State,'{malformed')
    Assert-Throws { Get-SmartM365OrchestratorMaintenanceState $root } '*'
    Write-Fixture $paths.State ([pscustomobject]@{SchemaVersion=1;Revision=3;Enabled='false'})
    Assert-Throws { Get-SmartM365OrchestratorMaintenanceState $root } '*missing*'
    $failure=& $runtime { $script:Starts=0; Update-OrchestratorMaintenance; Invoke-LaunchPhase (Get-Date); [pscustomobject]@{Healthy=$script:MaintenanceHealthy;Starts=$script:Starts} }
    Assert-Case (-not $failure.Healthy -and $failure.Starts -eq 0) 'Invalid state did not pause new launches.'
    [IO.File]::WriteAllBytes($paths.State,$goodBytes)
    Remove-Item -LiteralPath $paths.State
    Assert-Throws { Get-SmartM365OrchestratorMaintenanceState $root } '*missing*'
    Assert-Throws { Initialize-SmartM365OrchestratorMaintenance $root } '*disappeared*'
    [IO.File]::WriteAllBytes($paths.State,$goodBytes)
    Assert-Throws { Get-SmartM365OrchestratorMaintenanceState (Join-Path $root 'Unavailable') } '*'
    # GUI AST and XAML remain independently headless-testable; no live GUI is opened here.
    $guiPath=Join-Path $orchFolder 'SmartM365-Inventory-Orchestrator-GUI.ps1'
    $gui=[IO.File]::ReadAllText($guiPath)
    foreach($name in @('MaintenanceBannerText','MaintenanceReasonBox','EnableMaintenanceButton','DisableMaintenanceButton')) { Assert-Case ($gui.Contains('x:Name="'+$name+'"')) "Missing GUI control $name" }
    $guiDefinitions=Load-Functions $guiPath @('Set-GuiMaintenance','Refresh-MaintenanceView')
    Assert-Case (($guiDefinitions -join "`n") -notmatch 'Publish-SmartM365OrchestratorConfiguration|Start-Process|Stop-Process') 'Maintenance GUI publishes draft or controls processes.'
    Assert-Case (($guiDefinitions -join "`n") -match 'ExpectedRevision') 'GUI does not guard stale writes.'
    $runtimeSource=[IO.File]::ReadAllText((Join-Path $orchFolder 'SmartM365-Inventory-Orchestrator.ps1'))
    Assert-Case ($runtimeSource.Contains('MaintenanceModuleFingerprint = $maintenanceModuleFingerprint')) 'Maintenance module is not monitored for runtime updates.'
    Assert-Case ($runtimeSource.Contains("-Role 'Maintenance orchestrator runtime module'")) 'Runtime updates omit maintenance signature validation.'
    [pscustomobject]@{Status='Passed';TestCount=$script:Cases;ProductionActions=0;PowerShell=[string]$PSVersionTable.PSVersion}
}
finally {
    if ($runtime) { Remove-Module $runtime -ErrorAction SilentlyContinue }
    $resolved=[IO.Path]::GetFullPath($root)
    if ((Split-Path $resolved -Parent) -ne ([IO.Path]::GetTempPath()).TrimEnd('\') -or (Split-Path $resolved -Leaf) -notlike 'SmartM365-Maintenance-Test-*') { throw 'Unsafe synthetic cleanup target.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB9lsuS3jr6QREf
# yjA5DaLjBOJV9rpRksoQOFHU1mJ7LaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIN5lp1Q6Buq93ybBsz8UkOhzoZV+KXruyDRSz8WfaAqJMA0GCSqG
# SIb3DQEBAQUABIIBgF/XgnPUGYS6vMxrhBF160B9QV2FvXUyzMDTX7bD5pKU3CDk
# 1KhWAv8CGzu0Ev25zIagsbQ1aPIKNzM4TV4XrbwqHY80dIIrGYSJ6mTtNUZGx9Go
# 6/DrUH0XFbGniTzEEPx0CZ3rG+gUz0YBAeKeI3wuMN1JMstGjLtqrqV8QbsmTP/1
# c9p16ExOX/qSE68R6BbI6uKHYPgvXoh3N0LAOcDMOKEfAVk5kgpt+YlCXgVeCRTH
# xlbo7bBBwiQOuDN3BJorezez1L9ITZ14x1D078YAxS+MzyxdZxUq8hU/KVVxolad
# Fl2+7mfYJNUfyVjj0RrBeaUs24+nKB57ccg835khh8wpHuVilBgOxdSE4FDP/afw
# vFsKFwOHgNEO06fBb0ZayQ6AOvrUQbqAirh0fF5Or5LJ+ZRCxhtkiidIPn4Bfibe
# qmOmLXQckg/n9V5pYsBt1x8/wIGwQhd5D/Y3fA4eQouToZpIhamLsU52vRPvF10r
# vd0jwY+aVBkXjhMeeqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgyMDQz
# MzhaMC8GCSqGSIb3DQEJBDEiBCB7j/N9PO93VPd1hcdxHmlxE4JEZanUy97bUOaN
# 2zrtNzANBgkqhkiG9w0BAQEFAASCAgCcu2hEKxsAQbSKZJFjyblhNVGqjreGOIa4
# cbbzOW517qtjZWFhrAD2saHSog3OXeJr7bstxprI3myDmhJZQrK+ex1MtpgGpRvP
# bd7u9grSN2KxAzTlUPci8x5WOcKf6axsDNIhgKMRBIbRtiRQbYuew5AOo1Klszgn
# K/uDpvUDXtS7htNIY4vgzeqVzfmdVb1KOY28V/QN68RTQXXugoRtBJdvOC9WFAO0
# eCW5UvfJHDUF5k1yXizAi3p89iEYqhUih/3Ci66tRYnc9s52H5DhClQjyGrOjgN2
# 8r0BEmPWWjR3QcPCGBwmv2u7od5y92912SHEsUPHTypBsDHJgzllK3tcqeL+Bndb
# bvljyX7wdN05CxagOqYX12PyYz5d5pkIwGoVsLEDnsKc3n49+E96QlWMRQVhnuPJ
# 05WSDwoNd2unm9CGkq/ophQubSTkAP/7/PxtbfJH6pB+vtSLfBjqcaLYYBCq5FXq
# p4QAz9DhWbWtYAQg4mFU8npR7PZ+S7sVSHFjAo5BO+bhNyACykaqprVEe9mCDbc3
# h+X3sDZozsIqRdUywty0a12C81W6mLGHuahLF8KuopsqK5aFPqk2c5tjgc0G2NXT
# Ab03LAda5N3AOnlwkxPWsafEE9CvJuHdXGkBalElgvRUf28wwOItCNvxd7V+XwXw
# Dq/0ysjKTw==
# SIG # End signature block
