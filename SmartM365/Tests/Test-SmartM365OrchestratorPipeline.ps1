<#
.SYNOPSIS
Runs offline pipeline and production-manifest contract tests for the SmartM365 orchestrator.
.VERSION
1.2.0
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$orchestratorRoot = Join-Path -Path $smartM365Root -ChildPath 'SmartInventory\Orchestrator'
$smartInventoryRoot = Split-Path -Path $orchestratorRoot -Parent
$pipelineModuleFile = Join-Path -Path $orchestratorRoot -ChildPath 'SmartM365.Orchestrator.Pipeline.psm1'
$orchestratorPath = Join-Path -Path $orchestratorRoot -ChildPath 'SmartM365-Inventory-Orchestrator.ps1'
$jobsTemplatePath = Join-Path -Path $orchestratorRoot -ChildPath 'Orchestrator-Jobs.json.template'
$pipelineModule = Import-Module -Name $pipelineModuleFile -Force -PassThru -ErrorAction Stop

function Assert-True {
    param([Parameter(Mandatory)][bool]$Condition, [Parameter(Mandatory)][string]$Message)
    if (-not $Condition) { throw $Message }
}

function New-SyntheticJob {
    param([string]$Name, [string]$Group, [bool]$Enabled = $true, [string]$Mode = 'Elected', [string[]]$DependsOn = @())
    [pscustomobject][ordered]@{
        Name = $Name
        Group = $Group
        Enabled = $Enabled
        AssignmentMode = $Mode
        DependsOn = @($DependsOn)
        RequiredCapabilities = @('SharedRuntime')
    }
}

$temporaryRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ('SmartM365-OrchestratorPipeline-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $temporaryRoot -Force | Out-Null
    $document = [pscustomobject]@{
        Jobs = @(
            (New-SyntheticJob -Name 'Parent' -Group 'M365'),
            (New-SyntheticJob -Name 'Child' -Group 'Intune' -DependsOn @('Parent')),
            (New-SyntheticJob -Name 'Disabled' -Group 'Intune' -Enabled $false),
            (New-SyntheticJob -Name 'Manual' -Group 'Intune' -Mode 'Manual')
        )
    }

    $selection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -Pipeline Intune
    Assert-True ($selection.SelectedCount -eq 2) 'Dependency closure did not include the enabled parent.'
    Assert-True ($selection.AddedDependencyCount -eq 1) 'Dependency closure count is incorrect.'
    Assert-True (@($selection.SelectedJobs.Name) -contains 'Parent') 'The dependency parent is absent.'
    Assert-True (@($selection.ExcludedJobs | Where-Object Reason -EQ 'Disabled').Count -eq 1) 'Disabled exclusion was not reported.'
    Assert-True (@($selection.ExcludedJobs | Where-Object Reason -EQ 'Manual').Count -eq 1) 'Manual exclusion was not reported.'

    # -JobName selects only the named jobs; dependencies are external unless -IncludeDependencies.
    $jobSelection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -JobName Child
    Assert-True ($jobSelection.Pipeline -eq 'Jobs' -and $jobSelection.SelectedCount -eq 1 -and $jobSelection.AddedDependencyCount -eq 0) 'Job selection added a dependency without -IncludeDependencies.'
    Assert-True ((@($jobSelection.SelectedJobs[0].ExternalDependencies) -join ',') -eq 'Parent') 'Job selection did not report the external dependency.'
    Assert-True (@($jobSelection.ExcludedJobs | Where-Object Reason -EQ 'NotRequested').Count -eq 1) 'Unrequested runnable job was not reported as NotRequested.'
    $jobSelectionWithDependencies = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -JobName Child -IncludeDependencies
    Assert-True ($jobSelectionWithDependencies.SelectedCount -eq 2 -and @($jobSelectionWithDependencies.SelectedJobs[1].ExternalDependencies).Count -eq 0) '-IncludeDependencies did not add the dependency.'
    foreach ($refused in @('Disabled', 'Manual', 'Unknown')) {
        $wasRefused = $false
        try { Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -JobName $refused | Out-Null } catch { $wasRefused = $true }
        Assert-True $wasRefused "Job request '$refused' was accepted."
    }
    # Disabled or manual dependencies are ignored, as in the scheduled dependency gate.
    $ignoredDocument = [pscustomobject]@{ Jobs = @((New-SyntheticJob -Name 'Off' -Group 'M365' -Enabled $false), (New-SyntheticJob -Name 'Consumer' -Group 'M365' -DependsOn @('Off'))) }
    $ignoredSelection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $ignoredDocument -Pipeline Full
    Assert-True ($ignoredSelection.SelectedCount -eq 1 -and @($ignoredSelection.IgnoredDependencies).Count -eq 1 -and $ignoredSelection.IgnoredDependencies[0].Dependency -eq 'Off') 'A disabled dependency blocked the Full selection instead of being ignored.'

    $selectionAgain = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -Pipeline Intune
    Assert-True ((@($selection.SelectedJobs.Name) -join '|') -eq (@($selectionAgain.SelectedJobs.Name) -join '|')) 'Selection is not idempotent.'

    $request = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $temporaryRoot -Tenant test -Pipeline Intune -Selection $selection -ManifestHash 'ABC123' -RequestedBy 'offline-test' -RequestedFrom 'offline-test'
    Assert-True ($request.OverallStatus -eq 'Running' -and $request.PendingCount -eq 2) 'New request status is incorrect.'
    Assert-True (Test-Path -LiteralPath $request.RequestPath -PathType Leaf) 'Atomic request document is missing.'

    $retrySource = Join-Path -Path $temporaryRoot -ChildPath 'retry-source.tmp'
    $retryDestination = Join-Path -Path $temporaryRoot -ChildPath 'retry-destination.json'
    [IO.File]::WriteAllText($retrySource, '{"retry":true}', [Text.UTF8Encoding]::new($false))
    $retryAttempts = & $pipelineModule {
        param($SourcePath, $DestinationPath)
        $state = [pscustomobject]@{ Attempts = 0 }
        $operation = {
            param($Source, $Destination)
            $state.Attempts++
            if ($state.Attempts -lt 3) { throw [IO.IOException]::new('Synthetic transient move failure.') }
            [IO.File]::Move($Source, $Destination, $true)
        }.GetNewClosure()
        Move-SmartM365OrchestratorPipelineFileWithRetry `
            -SourcePath $SourcePath `
            -DestinationPath $DestinationPath `
            -MaximumAttempts 5 `
            -RetryDelayMilliseconds 1 `
            -MoveOperation $operation
        $state.Attempts
    } $retrySource $retryDestination
    Assert-True ($retryAttempts -eq 3 -and (Test-Path -LiteralPath $retryDestination -PathType Leaf)) 'Transient pipeline status replacement was not retried successfully.'

    $boundedRetryAttempts = & $pipelineModule {
        param($SourcePath, $DestinationPath)
        $state = [pscustomobject]@{ Attempts = 0 }
        $operation = {
            param($Source, $Destination)
            $state.Attempts++
            throw [UnauthorizedAccessException]::new('Synthetic persistent access denial.')
        }.GetNewClosure()
        try {
            Move-SmartM365OrchestratorPipelineFileWithRetry `
                -SourcePath $SourcePath `
                -DestinationPath $DestinationPath `
                -MaximumAttempts 2 `
                -RetryDelayMilliseconds 1 `
                -MoveOperation $operation
        }
        catch [UnauthorizedAccessException] { return $state.Attempts }
        throw 'Persistent access denial was unexpectedly swallowed.'
    } $retryDestination (Join-Path -Path $temporaryRoot -ChildPath 'never-created.json')
    Assert-True ($boundedRetryAttempts -eq 2) 'Persistent pipeline replacement failure did not stop at the configured retry bound.'

    $concurrentRejected = $false
    try {
        New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $temporaryRoot -Tenant test -Pipeline Intune -Selection $selection -ManifestHash 'ABC123' | Out-Null
    }
    catch { $concurrentRejected = $_.Exception.Message -like '*active pipeline request already exists*Queueing is unavailable*' }
    Assert-True $concurrentRejected 'A second request was accepted although no resident advertises the queue protocol.'

    # FIFO queue: accepted only when every expected resident advertises PipelineQueueProtocol 1.
    $queueRoot = Join-Path -Path $temporaryRoot -ChildPath 'Queue'
    $writeFixture = {
        param([string]$Path, $Fixture)
        New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force | Out-Null
        $null = Write-SmartM365JsonBytesAtomically -Path $Path -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($Fixture | ConvertTo-Json -Depth 20))) -Validate { param($value) $null = $value }
    }
    & $writeFixture (Join-Path $queueRoot 'Config/Orchestrator-Cluster.json.txt') ([pscustomobject]@{ ExpectedOrchestratorServers = @('WorkerA', 'WorkerB') })
    & $writeFixture (Join-Path $queueRoot 'WorkerA/Orchestrator-Heartbeat.json.txt') ([pscustomobject]@{ Lifecycle = 'Running'; ScriptVersion = '1.5.66'; PipelineQueueProtocol = 1 })
    & $writeFixture (Join-Path $queueRoot 'WorkerB/Orchestrator-Heartbeat.json.txt') ([pscustomobject]@{ Lifecycle = 'Running'; ScriptVersion = '1.5.65' })
    $parentSelection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -JobName Parent
    $childSelection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -JobName Child
    $queueHead = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $queueRoot -Tenant test -Pipeline Jobs -Selection $parentSelection -ManifestHash 'ABC123'
    Assert-True ($queueHead.QueuePosition -eq 0 -and $queueHead.OverallStatus -eq 'Running') 'The first request was not the running head of the queue.'
    $oldResidentRefused = $false
    try { New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $queueRoot -Tenant test -Pipeline Jobs -Selection $childSelection -ManifestHash 'ABC123' | Out-Null }
    catch { $oldResidentRefused = $_.Exception.Message -like '*Queueing is unavailable*WorkerB runs v1.5.65*' }
    Assert-True $oldResidentRefused 'A request was queued although a resident lacks the queue protocol.'
    Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $queueRoot 'PipelineRuns') -Directory).Count -eq 1) 'A refused queued request left a batch folder.'
    & $writeFixture (Join-Path $queueRoot 'WorkerB/Orchestrator-Heartbeat.json.txt') ([pscustomobject]@{ Lifecycle = 'Running'; ScriptVersion = '1.5.66'; PipelineQueueProtocol = 1 })
    $queued = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $queueRoot -Tenant test -Pipeline Jobs -Selection $childSelection -ManifestHash 'ABC123'
    Assert-True ($queued.QueuePosition -eq 1 -and $queued.OverallStatus -eq 'Queued' -and (@($queued.QueuedBehind) -join ',') -eq $queueHead.BatchId) 'The second request was not queued behind the first one.'
    Assert-True ([string]$queued.CreatedAtUtc -gt [string]$queueHead.CreatedAtUtc) 'Queue order does not follow creation time.'
    foreach ($duplicate in @($parentSelection, $childSelection)) {
        $duplicateRefused = $false
        try { New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $queueRoot -Tenant test -Pipeline Jobs -Selection $duplicate -ManifestHash 'ABC123' | Out-Null }
        catch { $duplicateRefused = $_.Exception.Message -like '*identical pipeline request is already*' }
        Assert-True $duplicateRefused 'An identical running or queued request was accepted.'
    }
    $fullSelection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -Pipeline Full
    $limitRefused = $false
    try { New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $queueRoot -Tenant test -Pipeline Full -Selection $fullSelection -ManifestHash 'ABC123' -MaxActiveRequests 2 | Out-Null }
    catch { $limitRefused = $_.Exception.Message -like '*queue is full*limit 2*' }
    Assert-True $limitRefused 'The queue limit was not enforced.'
    $activeQueue = @(Get-SmartM365OrchestratorActivePipelineRuns -SharedDataFolderPath $queueRoot)
    Assert-True ($activeQueue.Count -eq 2 -and $activeQueue[0].BatchId -eq $queueHead.BatchId -and $activeQueue[0].OverallStatus -eq 'Running' -and $activeQueue[1].OverallStatus -eq 'Queued' -and $activeQueue[1].QueuePosition -eq 1) 'Active runs do not report the FIFO queue.'

    # Activation is recorded once, by the first resident, then kept.
    $activated = Set-SmartM365OrchestratorPipelineActivation -SharedDataFolderPath $queueRoot -BatchId $queueHead.BatchId -ActivatedBy WorkerA
    $activatedAgain = Set-SmartM365OrchestratorPipelineActivation -SharedDataFolderPath $queueRoot -BatchId $queueHead.BatchId -ActivatedBy WorkerB
    $headStatus = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $queueRoot -BatchId $queueHead.BatchId
    Assert-True (-not [string]::IsNullOrWhiteSpace($activated) -and $activatedAgain -eq $activated -and $headStatus.ActivatedAtUtc -eq $activated -and [string]$headStatus.Request.ActivatedBy -eq 'WorkerA') 'Activation was not recorded exactly once.'
    Assert-True ([string]::IsNullOrWhiteSpace((Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $queueRoot -BatchId $queued.BatchId).ActivatedAtUtc)) 'A queued request was activated before it reached the head.'

    # Cancelling a queued request finishes it at once; finishing the head promotes the next one.
    $cancelQueued = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $queueRoot -Tenant test -Pipeline Full -Selection $fullSelection -ManifestHash 'ABC123'
    Assert-True ($cancelQueued.QueuePosition -eq 2) 'The third request was not queued at position 2.'
    foreach ($queuedJob in @($cancelQueued.Jobs)) {
        Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $queueRoot -BatchId $cancelQueued.BatchId -JobName $queuedJob.JobName -Status Cancelled -OwnerServer '' | Out-Null
    }
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $queueRoot -BatchId $queueHead.BatchId -JobName Parent -Status Success -OwnerServer WorkerA | Out-Null
    $promoted = @(Get-SmartM365OrchestratorActivePipelineRuns -SharedDataFolderPath $queueRoot)
    Assert-True ($promoted.Count -eq 1 -and $promoted[0].BatchId -eq $queued.BatchId -and $promoted[0].OverallStatus -eq 'Running' -and $promoted[0].QueuePosition -eq 0) 'The queued request was not promoted when the head finished.'

    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId -JobName Parent -Status Success -Attempt 0 -OwnerServer SERVER-A | Out-Null
    $retryNotBefore = [datetime]::UtcNow.AddMinutes(5).ToString('o')
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId -JobName Child -Status RetryScheduled -Attempt 1 -OwnerServer SERVER-B -NotBeforeUtc $retryNotBefore | Out-Null
    $retrying = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId
    $retryingChild = @($retrying.Jobs | Where-Object JobName -EQ Child)[0]
    Assert-True (-not $retrying.IsTerminal -and $retryingChild.Attempt -eq 1 -and $retryingChild.NotBeforeUtc -eq $retryNotBefore) 'Retry timing/attempt metadata was not preserved.'
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId -JobName Child -Status CompletedWithWarnings -Attempt 0 -OwnerServer SERVER-B | Out-Null
    $completed = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId
    Assert-True ($completed.IsTerminal -and $completed.OverallStatus -eq 'CompletedWithWarnings') 'Warning completion was not aggregated correctly.'

    $secondRequest = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $temporaryRoot -Tenant test -Pipeline Intune -Selection $selection -ManifestHash 'ABC123'
    Assert-True ($secondRequest.BatchId -ne $request.BatchId) 'A completed request prevented the next unique batch.'
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $secondRequest.BatchId -JobName Parent -Status Failed -Attempt 1 -OwnerServer SERVER-A -Detail 'Synthetic failure' | Out-Null
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $secondRequest.BatchId -JobName Child -Status BlockedDependencyFailed -OwnerServer SERVER-B | Out-Null
    $failed = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $temporaryRoot -BatchId $secondRequest.BatchId
    Assert-True ($failed.IsTerminal -and $failed.OverallStatus -eq 'Failed' -and $failed.FailedCount -eq 2) 'Failure aggregation is incorrect.'

    $source = Get-Content -LiteralPath $orchestratorPath -Raw
    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$tokens, [ref]$parseErrors)
    Assert-True (@($parseErrors).Count -eq 0) 'Orchestrator source failed parser validation.'
    $scanDefinitions = foreach ($functionName in @('Update-OrchestratorPipelineRequests', 'Get-OrchestratorPipelineDependencyStatus', 'Get-OrchestratorExternalDependencyStatus', 'Repair-OrchestratorPipelineOrphanStatus', 'ConvertFrom-StateTime', 'Get-OrchestratorPipelineWaitStartUtc', 'Get-OrchestratorEffectiveDependencyMode')) {
        $node = $ast.Find({ param($candidate) $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq $functionName }, $true)
        if ($null -eq $node) { throw "Missing orchestrator function: $functionName" }
        $node.Extent.Text
    }

    $scanRoot = Join-Path -Path $temporaryRoot -ChildPath 'Scan'
    New-Item -ItemType Directory -Path $scanRoot -Force | Out-Null
    $scanManifestPath = Join-Path -Path $scanRoot -ChildPath 'Orchestrator-Jobs.json'
    $document | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $scanManifestPath -Encoding utf8
    $scanSelection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -Pipeline Full
    $scanRequest = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $scanRoot -Tenant test -Pipeline Full -Selection $scanSelection -ManifestHash ((Get-FileHash -LiteralPath $scanManifestPath -Algorithm SHA256).Hash)
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $scanRoot -BatchId $scanRequest.BatchId -JobName Child -Status RetryScheduled -Attempt 1 -OwnerServer SERVER-B -NotBeforeUtc ([datetime]::UtcNow.AddMinutes(10).ToString('o')) | Out-Null
    $scanModule = New-Module -ScriptBlock ([scriptblock]::Create($scanDefinitions -join [Environment]::NewLine))
    $scanResult = & $scanModule {
        param($Root, $ManifestPath, $JobsDocument, $BatchId)
        $script:Settings = @{ SharedDataFolderPath = $Root; JobsManifestPath = $ManifestPath }
        $script:Manifest = @{ JobsByName = @{ Parent = $JobsDocument.Jobs[0]; Child = $JobsDocument.Jobs[1] } }
        $script:PipelinePending = @{}
        $script:Tenant = 'test'
        function script:Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message) }
        function script:Set-OrchestratorPipelineJobStatus { param($BatchId, $JobName, $Status, $Detail) throw "Unexpected rejection: $JobName/$Status/$Detail" }
        Update-OrchestratorPipelineRequests
        [pscustomobject]@{
            PendingNames = @($script:PipelinePending.Keys)
            ParentStatus = Get-OrchestratorPipelineDependencyStatus -BatchId $BatchId -JobName Parent
            BatchJobNames = @($script:PipelineBatchJobNames[$BatchId])
            ParentActivatedAtUtc = $script:PipelinePending['Parent'].ActivatedAtUtc
            ParentWaitStartUtc = Get-OrchestratorPipelineWaitStartUtc -PipelineInfo $script:PipelinePending['Parent']
        }
    } $scanRoot $scanManifestPath $document $scanRequest.BatchId
    Assert-True ($scanResult.PendingNames.Count -eq 1 -and $scanResult.PendingNames[0] -eq 'Parent') 'A retry was made runnable before NotBeforeUtc.'
    Assert-True ($scanResult.ParentStatus -eq 'Pending') 'Shared dependency status was not read from the batch.'
    Assert-True ((@($scanResult.BatchJobNames | Sort-Object) -join ',') -eq 'Child,Parent') 'The request job names were not recorded for the dependency gate.'
    $scanActivated = (Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $scanRoot -BatchId $scanRequest.BatchId).ActivatedAtUtc
    Assert-True (-not [string]::IsNullOrWhiteSpace($scanActivated) -and $scanResult.ParentWaitStartUtc -eq $scanResult.ParentActivatedAtUtc -and $scanResult.ParentActivatedAtUtc -eq ([datetime]::Parse($scanActivated, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime())) 'The resident did not record the activation or does not count waits from it.'

    # Waits count from activation; retries without it keep their creation time.
    $waitStarts = & $scanModule {
        $created = [datetime]'2026-10-28T08:00:00Z'
        $activatedAt = [datetime]'2026-10-28T11:00:00Z'
        @((Get-OrchestratorPipelineWaitStartUtc -PipelineInfo @{ CreatedAtUtc = $created; ActivatedAtUtc = $activatedAt }), (Get-OrchestratorPipelineWaitStartUtc -PipelineInfo @{ CreatedAtUtc = $created }))
    }
    Assert-True ($waitStarts[0] -eq [datetime]'2026-10-28T11:00:00Z' -and $waitStarts[1] -eq [datetime]'2026-10-28T08:00:00Z') 'Pipeline wait start does not prefer the activation time.'

    # Two active requests: the resident processes and activates only the oldest one.
    $fifoRoot = Join-Path -Path $temporaryRoot -ChildPath 'Fifo'
    & $writeFixture (Join-Path $fifoRoot 'Config/Orchestrator-Cluster.json.txt') ([pscustomobject]@{ ExpectedOrchestratorServers = @('WorkerA') })
    & $writeFixture (Join-Path $fifoRoot 'WorkerA/Orchestrator-Heartbeat.json.txt') ([pscustomobject]@{ Lifecycle = 'Running'; ScriptVersion = '1.5.66'; PipelineQueueProtocol = 1 })
    $scanHash = (Get-FileHash -LiteralPath $scanManifestPath -Algorithm SHA256).Hash
    $fifoHead = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $fifoRoot -Tenant test -Pipeline Jobs -Selection $parentSelection -ManifestHash $scanHash
    $fifoNext = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $fifoRoot -Tenant test -Pipeline Jobs -Selection $childSelection -ManifestHash $scanHash
    $fifoResult = & $scanModule {
        param($Root, $ManifestPath, $JobsDocument)
        $script:Settings = @{ SharedDataFolderPath = $Root; JobsManifestPath = $ManifestPath }
        $script:Manifest = @{ JobsByName = @{ Parent = $JobsDocument.Jobs[0]; Child = $JobsDocument.Jobs[1] } }
        $script:PipelinePending = @{}
        $script:Logged = [Collections.Generic.List[string]]::new()
        function script:Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now, $Level) $null = $Now; $script:Logged.Add("$Level|$Message") }
        Update-OrchestratorPipelineRequests
        [pscustomobject]@{ PendingNames = @($script:PipelinePending.Keys); BatchIds = @($script:PipelinePending.Values | ForEach-Object { $_.BatchId } | Sort-Object -Unique); Logged = $script:Logged.ToArray() }
    } $fifoRoot $scanManifestPath $document
    Assert-True ((@($fifoResult.PendingNames) -join ',') -eq 'Parent' -and (@($fifoResult.BatchIds) -join ',') -eq $fifoHead.BatchId) 'The resident did not process only the oldest active request.'
    Assert-True (@($fifoResult.Logged | Where-Object { $_ -like 'INFO|*1 request(s) queued behind it*' }).Count -eq 1) 'The queued request was not logged at INFO.'
    Assert-True ([string]::IsNullOrWhiteSpace((Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $fifoRoot -BatchId $fifoNext.BatchId).ActivatedAtUtc)) 'The queued request was activated by the resident.'

    # A Starting/Running status whose final shared write failed is reconciled by its own server only
    # once the job is no longer supervised, from the recorded result of the request occurrence.
    $orphanResult = & $scanModule {
        $me = $env:COMPUTERNAME
        $now = [datetime]'2026-09-30T12:00:00'
        $occurrence = [datetime]'2026-09-21T17:21:08'
        $script:Settings = @{ ElectionClaimsPath = 'claims'; ElectionClaimGraceMinutes = 15 }
        $job = { param($Name, $Mode, $Timeout = 60) [pscustomobject]@{ Name = $Name; AssignmentMode = $Mode; TimeoutMinutes = $Timeout } }
        $script:Manifest = @{ JobsByName = @{
            ClaimDone = & $job ClaimDone Elected; CsvDone = & $job CsvDone Pinned; Supervised = & $job Supervised Elected; Peer = & $job Peer Elected
            Retry = & $job Retry Pinned; Recent = & $job Recent Pinned 1440; Lost = & $job Lost Pinned; StateRunning = & $job StateRunning Pinned } }
        $script:RunningJobs = @{ Supervised = @{} }
        $states = @{
            Retry = @{ Running = $null; PendingRetry = @{ PipelineBatchId = 'B1'; Attempt = 2; NotBefore = '2026-09-30T12:05:00.0000000+02:00' } }
            StateRunning = @{ Running = @{ Pid = 1 }; PendingRetry = $null }
        }
        function script:Get-JobState { param($JobName) if ($states.ContainsKey($JobName)) { $states[$JobName] } else { @{ Running = $null; PendingRetry = $null } } }
        function script:Get-SmartM365OrchestratorOccurrenceClaim {
            param($ClaimsRootPath, $JobName, $Occurrence)
            $null = $ClaimsRootPath
            if ($JobName -eq 'ClaimDone' -and $Occurrence -eq [datetime]'2026-09-21T17:21:08') { return [pscustomobject]@{ Claim = [pscustomobject]@{ Status = 'CompletedWithWarnings'; OwnerServer = $env:COMPUTERNAME } } }
            return $null
        }
        function script:Get-OrchestratorJobRunsForWindow {
            param($Now, $Hours)
            $null = $Now; $script:WindowHours = $Hours
            @(
                [pscustomobject]@{ Server = $env:COMPUTERNAME; JobName = 'CsvDone'; ScheduledTime = '2026-09-21 17:21:08'; EndTime = '2026-09-21 17:31:29'; Status = 'Success' }
                [pscustomobject]@{ Server = $env:COMPUTERNAME; JobName = 'CsvDone'; ScheduledTime = '2026-09-22 00:30:00'; EndTime = '2026-09-22 00:40:00'; Status = 'Failed' }
                [pscustomobject]@{ Server = 'OTHER'; JobName = 'Lost'; ScheduledTime = '2026-09-21 17:21:08'; EndTime = '2026-09-21 17:40:00'; Status = 'Success' }
            )
        }
        $script:Calls = [Collections.Generic.List[object]]::new()
        function script:Set-OrchestratorPipelineJobStatus { param($BatchId, $JobName, $Status, $Attempt, $Detail, $NotBeforeUtc) $script:Calls.Add([pscustomobject]@{ BatchId = $BatchId; JobName = $JobName; Status = $Status; Attempt = $Attempt; Detail = $Detail; NotBeforeUtc = $NotBeforeUtc }) }
        function script:Write-OrchestratorLog { param($Message, $Level) $null = $Message, $Level }
        function script:Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now, $Level) $null = $Key, $Message, $Now, $Level }
        $status = { param($Name, $Owner, $Updated) [pscustomobject]@{ JobName = $Name; Status = 'Running'; OwnerServer = $Owner; Attempt = 0; UpdatedAtUtc = $Updated } }
        $run = [pscustomobject]@{ BatchId = 'B1'; Jobs = @(
            (& $status ClaimDone $me '2026-09-21T15:21:40Z'), (& $status CsvDone $me '2026-09-21T15:21:51Z'), (& $status Supervised $me '2026-09-21T15:21:51Z'),
            (& $status Peer 'OTHER' '2026-09-21T15:21:51Z'), (& $status Retry $me '2026-09-30T09:59:00Z'), (& $status Recent $me '2026-09-30T09:30:00Z'),
            (& $status Lost $me '2026-09-21T15:21:51Z'), (& $status StateRunning $me '2026-09-21T15:21:51Z'),
            [pscustomobject]@{ JobName = 'Done'; Status = 'Success'; OwnerServer = $me; Attempt = 0; UpdatedAtUtc = '2026-09-21T15:30:00Z' }) }
        Repair-OrchestratorPipelineOrphanStatus -Run $run -Occurrence $occurrence -Now $now
        [pscustomobject]@{ Calls = $script:Calls.ToArray(); WindowHours = $script:WindowHours }
    }
    $orphanCalls = @{}
    foreach ($call in @($orphanResult.Calls)) { $orphanCalls[[string]$call.JobName] = $call }
    Assert-True ((@($orphanCalls.Keys | Sort-Object) -join ',') -eq 'ClaimDone,CsvDone,Lost,Retry') ("Orphan reconciliation touched the wrong jobs: {0}" -f (@($orphanCalls.Keys | Sort-Object) -join ','))
    Assert-True ($orphanCalls.ClaimDone.Status -eq 'CompletedWithWarnings' -and $orphanCalls.ClaimDone.Detail -like '*shared occurrence claim*') 'An orphan elected status was not reconciled from its terminal occurrence claim.'
    Assert-True ($orphanCalls.CsvDone.Status -eq 'Success' -and $orphanCalls.CsvDone.Detail -like '*job-run history*') 'An orphan status was not reconciled from the job-run row of the request occurrence.'
    Assert-True ($orphanCalls.Retry.Status -eq 'RetryScheduled' -and $orphanCalls.Retry.Attempt -eq 2 -and -not [string]::IsNullOrWhiteSpace([string]$orphanCalls.Retry.NotBeforeUtc)) 'A pending local retry was not reconciled as RetryScheduled.'
    Assert-True ($orphanCalls.Lost.Status -eq 'Interrupted') 'An orphan without any recorded result was not marked Interrupted after its timeout window (a peer job-run row must not be used).'
    Assert-True ($orphanResult.WindowHours -ge 212) 'The job-run history window does not reach the request occurrence.'

    # Dependencies outside a -Job request follow the requested job's scheduled dependency rule.
    $externalResults = & $scanModule {
        $script:Settings = @{ DistributedSchedulingEnabled = $true }
        $script:RunningJobs = @{}
        $script:Now = Get-Date
        function script:Get-OrchestratorDependencyMaxAge { param($DependencyJob, $OverrideHours, $Now) $null = $DependencyJob, $OverrideHours, $Now; 26 }
        function script:Get-OrchestratorDependencyFreshStatus { param($DependencyJob, $MaxAgeHours, $Now) $null = $DependencyJob, $MaxAgeHours, $Now; $script:FreshStatus }
        function script:Get-OrchestratorSharedDependencyStatus { param($Job, $Now) $null = $Job, $Now; $script:SharedStatus }
        function script:Get-OrchestratorDependencyActivity { param($DependencyJob, $DownstreamOccurrence, $Now) $null = $DependencyJob, $DownstreamOccurrence, $Now; [pscustomobject]@{ State = $script:ActivityState; Detail = 'synthetic' } }
        $script:ActivityState = 'Clear'
        function script:Test-JobAllowedOnServer { param($Job) $null = $Job; $true }
        function script:Get-JobState { param($JobName) $null = $JobName; @{ LastStatus = $script:LegacyStatus; PendingRetry = $null } }
        function script:Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now) $null = $Key, $Message, $Now }
        $elected = [pscustomobject]@{ Name = 'Elected'; Enabled = $true; AssignmentMode = 'Elected'; ContinueOnError = $false }
        $legacy = [pscustomobject]@{ Name = 'Legacy'; Enabled = $true; AssignmentMode = 'Legacy'; ContinueOnError = $false }
        $off = [pscustomobject]@{ Name = 'Off'; Enabled = $false; AssignmentMode = 'Elected'; ContinueOnError = $false }
        $script:Manifest = @{ JobsByName = @{ Elected = $elected; Legacy = $legacy; Off = $off } }
        $fresh = [pscustomobject]@{ Name = 'Consumer'; DependencyMode = 'FreshSuccess'; DependencyMaxAgeHours = 48 }
        $latest = [pscustomobject]@{ Name = 'Consumer'; DependencyMode = 'LatestOccurrence'; DependencyMaxAgeHours = 0 }
        $r = [ordered]@{}
        $script:FreshStatus = 'Ready'; $r.FreshReady = Get-OrchestratorExternalDependencyStatus -Job $fresh -DependencyName Elected -Now $script:Now
        $script:ActivityState = 'Waiting'; $r.FreshActive = Get-OrchestratorExternalDependencyStatus -Job $fresh -DependencyName Elected -Now $script:Now; $script:ActivityState = 'Clear'
        $script:FreshStatus = 'Waiting'; $r.FreshStale = Get-OrchestratorExternalDependencyStatus -Job $fresh -DependencyName Elected -Now $script:Now
        $script:SharedStatus = 'Ready'; $r.SharedReady = Get-OrchestratorExternalDependencyStatus -Job $latest -DependencyName Elected -Now $script:Now
        $script:SharedStatus = 'Failed'; $r.SharedFailed = Get-OrchestratorExternalDependencyStatus -Job $latest -DependencyName Elected -Now $script:Now
        $script:SharedStatus = 'Running'; $r.SharedRunning = Get-OrchestratorExternalDependencyStatus -Job $latest -DependencyName Elected -Now $script:Now
        $script:LegacyStatus = 'Success'; $r.LegacySuccess = Get-OrchestratorExternalDependencyStatus -Job $latest -DependencyName Legacy -Now $script:Now
        $script:LegacyStatus = 'Failed'; $r.LegacyFailed = Get-OrchestratorExternalDependencyStatus -Job $latest -DependencyName Legacy -Now $script:Now
        $r.Disabled = Get-OrchestratorExternalDependencyStatus -Job $latest -DependencyName Off -Now $script:Now
        $r.Unknown = Get-OrchestratorExternalDependencyStatus -Job $latest -DependencyName Missing -Now $script:Now
        [pscustomobject]$r
    }
    $expectedExternal = [ordered]@{ FreshReady = 'Ready'; FreshActive = 'Waiting'; FreshStale = 'Waiting'; SharedReady = 'Ready'; SharedFailed = 'Failed'; SharedRunning = 'Waiting'; LegacySuccess = 'Ready'; LegacyFailed = 'Failed'; Disabled = 'Ready'; Unknown = 'Failed' }
    foreach ($case in $expectedExternal.Keys) {
        Assert-True ($externalResults.$case -eq $expectedExternal[$case]) "External dependency case $case returned '$($externalResults.$case)', expected '$($expectedExternal[$case])'."
    }

    $mailNode = $ast.Find({ param($candidate) $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq 'Send-OrchestratorRuntimeUpdateEmail' }, $true)
    Assert-True ($null -ne $mailNode) 'Runtime update email function is missing.'
    $mailModule = New-Module -ScriptBlock ([scriptblock]::Create($mailNode.Extent.Text))
    $mailResult = & $mailModule {
        $script:RuntimeUpdateBaseline = [pscustomobject]@{ ScriptVersion = [version]'1.5.13'; CoreVersion = [version]'1.0.49' }
        $script:RunningJobs = @{ JobA = @{}; JobB = @{} }
        $script:LastRuntimeUpdateMailIdentity = ''
        $script:Tenant = 'test'
        $script:ScriptName = 'SmartM365-Inventory-Orchestrator'
        $script:ScriptVersion = '1.5.14'
        $script:MailCalls = New-Object 'System.Collections.Generic.List[object]'
        function script:ConvertTo-HtmlText { param($Text) [Net.WebUtility]::HtmlEncode([string]$Text) }
        function script:Send-OrchestratorMail {
            param($Subject, $HtmlBody)
            $script:MailCalls.Add([pscustomobject]@{ Subject = $Subject; Body = $HtmlBody })
            return $true
        }
        $candidate = [pscustomobject]@{ ScriptVersion = [version]'1.5.14'; CoreVersion = [version]'1.0.50'; Identity = 'offline-runtime-1.5.14' }
        $first = Send-OrchestratorRuntimeUpdateEmail -Candidate $candidate -DetectedAt ([datetime]'2026-09-21T13:00:00Z')
        $second = Send-OrchestratorRuntimeUpdateEmail -Candidate $candidate -DetectedAt ([datetime]'2026-09-21T13:01:00Z')
        [pscustomobject]@{ First = $first; Second = $second; Calls = $script:MailCalls.ToArray() }
    }
    Assert-True ($mailResult.First -and $mailResult.Second -and $mailResult.Calls.Count -eq 1) 'Runtime update mail anti-duplication failed.'
    Assert-True ($mailResult.Calls[0].Subject -like '*1.5.13 -> 1.5.14*') 'Runtime update mail subject does not contain the version transition.'
    Assert-True ($mailResult.Calls[0].Body -like '*Stable, parser-valid and Authenticode-approved*' -and $mailResult.Calls[0].Body -like '*Detached jobs still supervised*') 'Runtime update mail body is incomplete.'

    $productionDocument = Get-Content -LiteralPath $jobsTemplatePath -Raw | ConvertFrom-Json -Depth 100
    $missingExternalActionOptIns = foreach ($job in @($productionDocument.Jobs)) {
        if ([string]::IsNullOrWhiteSpace([string]$job.ScriptPath)) { continue }
        $jobScriptPath = Join-Path -Path $smartInventoryRoot -ChildPath ([string]$job.ScriptPath)
        if (-not (Test-Path -LiteralPath $jobScriptPath -PathType Leaf)) { continue }
        $jobScriptSource = Get-Content -LiteralPath $jobScriptPath -Raw
        $supportsExternalActionOptIn = $jobScriptSource -match '\[switch\]\s*\$EnableConfiguredExternalActions'
        $enablesExternalActions = [string]$job.Arguments -match '(^|\s)-EnableConfiguredExternalActions(?=\s|$)'
        if ($supportsExternalActionOptIn -and -not $enablesExternalActions) { [string]$job.Name }
    }
    Assert-True (@($missingExternalActionOptIns).Count -eq 0) ("Jobs declaring EnableConfiguredExternalActions must opt in explicitly: {0}" -f (@($missingExternalActionOptIns) -join ', '))
    $productionSelection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $productionDocument -Pipeline Full
    Assert-True ($productionSelection.SelectedCount -eq 49) 'Template Full selection must contain the 49 enabled non-manual jobs.'
    Assert-True ($productionSelection.ExcludedCount -eq 6) 'Full selection must exclude the four disabled jobs and the two manual SharePoint on-premises jobs.'
    Assert-True (@($productionSelection.ExcludedJobs | Where-Object { $_.Name -like 'SharePointOnPrem-*' -and $_.Reason -eq 'Manual' }).Count -eq 2) 'Manual SharePoint on-premises jobs must be excluded from Full pipelines.'
    Assert-True (@($productionDocument.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare').Count -eq 0) 'Retired CMDB preparation must not remain in the template.'
    Assert-True (@($productionSelection.IgnoredDependencies).Count -eq 3) 'The 3 disabled Backup dependencies of WorkplaceEvidence-Prepare must be ignored, not block Full.'

    $facadeRoot = Join-Path -Path $temporaryRoot -ChildPath 'Facade'
    $facadeElectionRoot = Join-Path -Path $facadeRoot -ChildPath 'Election'
    New-Item -ItemType Directory -Path $facadeElectionRoot -Force | Out-Null
    $facadeManifestPath = Join-Path -Path $facadeRoot -ChildPath 'Orchestrator-Jobs.json'
    Copy-Item -LiteralPath $jobsTemplatePath -Destination $facadeManifestPath
    $facadePlan = [pscustomobject]@{
        PlanId = 'offline-plan'
        Assignments = @($productionSelection.SelectedJobs | ForEach-Object { [pscustomobject]@{ JobName = $_.Name; OwnerServer = $env:COMPUTERNAME } })
    }
    $facadePlan | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $facadeElectionRoot 'Orchestrator-ElectionPlan.json') -Encoding utf8
    $facadePath = Join-Path -Path $orchestratorRoot -ChildPath 'SmartM365-Inventory-Pipeline.ps1'
    $facadeOutput = & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -ExecutionPolicy Bypass -File $facadePath -Tenant test -Pipeline Full -Collect -NoWait -JobsManifestPath $facadeManifestPath -SharedDataFolderPath $facadeRoot 2>&1
    Assert-True ($LASTEXITCODE -eq 0) ("The offline Collect/NoWait façade failed: {0}" -f (@($facadeOutput) -join [Environment]::NewLine))
    Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $facadeRoot 'PipelineRuns') -Directory).Count -eq 1) 'The façade did not publish exactly one batch.'

    $launcherCheck = & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -ExecutionPolicy Bypass -File (Join-Path -Path $orchestratorRoot -ChildPath 'New-SmartM365-OrchestratorJobLaunchers.ps1') -Check 2>&1
    Assert-True ($LASTEXITCODE -eq 0) ("LaunchersByOrchestrator or Launchers\Jobs is not aligned with the jobs template: {0}" -f (@($launcherCheck) -join [Environment]::NewLine))

    foreach ($requiredHook in @(
        'Update-OrchestratorPipelineRequests',
        'Get-OrchestratorPipelineDependencyStatus',
        'Get-OrchestratorExternalDependencyStatus -Job $job -DependencyName',
        'PipelineBatchId',
        'Set-OrchestratorPipelineJobStatus',
        'Send-OrchestratorRuntimeUpdateEmail -Candidate $candidate',
        "if (`$reason -in @('due', 'forced'))"
    )) {
        Assert-True ($source.Contains($requiredHook)) "Orchestrator pipeline hook is missing: $requiredHook"
    }
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) { Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ("[{0}] SmartM365 Orchestrator pipeline tests passed." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB3lV0n7zwKGnuf
# 7Rc5hY3Sjpp8cK2SPBa/hT1+lNWiVaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIILrgZLA6924Iy6GY2sWSSe11gRn258PxXvuOl7MnE24MA0GCSqG
# SIb3DQEBAQUABIIBgH6KS9lU86bQOtojjssZhmvKrgyHp7SEJwUx4HnGsOEcLBmB
# WC9S3ELYUrg0nTAvUazXi4mZo9sSiyCcSG9MPEDPsG0V8/6K1ekfYBdZWS47mgQ/
# vAJMhRFCAAQxFUu21qSMO3ahod748AbYlil6GA0hCPLnoLPFtYltv28UI0GOM2gQ
# HTKoPESI8t4RyuqMk9dMrRB8f/XWBac57+2rCo163IOWj3g93kv7W6bes3i+u3d5
# HmpKU+4Joelq2Yt51Et10/lMzbx4mm0nwpcrZ7CoM4Y6sVc3ziY10P1LN0HlMptv
# 0TvKuJ2O9VEDVQBlpsy0dyE7cwIi8Ff2MJZS0Nsd3g5ZWwvdzKbpUDSLWLeDhY2B
# xBP9UdO+zsIm9a6r6CVuIPQiLIduNrQ6H7L7xda+qwpFKA4ZJ7VS9fFXBvK6lQWE
# rJHCFGnfV1RiW61SP6SUAQD7HVEvSvmYzzRRhsaqevYJ3jYt2BoDRCkWBAxuiLrI
# xez7pckDfpOJ/aVWuaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxMDA2
# NDVaMC8GCSqGSIb3DQEJBDEiBCAScEcFqvk9zDKXSS8xfsBIfg3ZrxmdaDPN/QlP
# 7GL7VDANBgkqhkiG9w0BAQEFAASCAgCdiBSYl6ITqV+GHNYkxk+O9wz6RxC7cvI1
# k7faKpK2pDGvi4ROimNiWMR7LrkjL8vbphpimYLEvcT3UflDj2g7Axb7D+cT7H+a
# Y1jnn6bdCE8dEh1IGhv/+1vzFbYfiBsZkFjntUAj2xpusnmBQUbqdk7P1qVfYMU+
# zjWlFw6RFMcKpuHRhoeVNSOOCCl6EM/hWAwNRR5yg9mI5+CoJFcC4WRf9hnUCG7Z
# tv2A5SUuQaFReHlaWtTS+AUp6LJOXizfosw8+/sK0cC/AN3FyPC9BkGrHxcjRIuw
# k4Km/gkbGN/2TFElmt9lrv2x2PN8toJOUgoROlx1BmGVbYU1Yci959Y+yrsLAmrg
# 68X/RUZOJ7un0XLOXfnfNZOLtT956uEeV0qSy0grNdlWKh2pRcpfieMOh6LN6EDH
# 8QMyCH53/NFN8PKv8dI4uEB/BKmQ0rxSxvJWucQyp0F/NqBg3pGaMKEoqTfpl00v
# +t8zCi3vqbMO3rMSk3aU4ij8jD2Ia+f99rN7kPsCGCaM0z01/BFsnMcDjLGi51+p
# KCTJqt1HxJSyjad/fhNkM2LQiSpYASXeCZPDgkRqw6Bq7VxuSXMt8H5EG4g1OHHw
# AKzf+g18y76RV/k7jx/roYqwLziB7HkUWfTrLjDqzAu3rWKVb+m6ZcaraOKrMfLa
# 9Falnd8XvQ==
# SIG # End signature block
