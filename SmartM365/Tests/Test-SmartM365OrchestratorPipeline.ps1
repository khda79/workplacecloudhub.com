<#
.SYNOPSIS
Runs offline pipeline and production-manifest contract tests for the SmartM365 orchestrator.
.VERSION
1.1.3
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
    catch { $concurrentRejected = $_.Exception.Message -like '*active pipeline request already exists*' }
    Assert-True $concurrentRejected 'A concurrent active pipeline request was accepted.'

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
    $scanDefinitions = foreach ($functionName in @('Update-OrchestratorPipelineRequests', 'Get-OrchestratorPipelineDependencyStatus', 'Get-OrchestratorExternalDependencyStatus', 'Repair-OrchestratorPipelineOrphanStatus', 'ConvertFrom-StateTime')) {
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
        }
    } $scanRoot $scanManifestPath $document $scanRequest.BatchId
    Assert-True ($scanResult.PendingNames.Count -eq 1 -and $scanResult.PendingNames[0] -eq 'Parent') 'A retry was made runnable before NotBeforeUtc.'
    Assert-True ($scanResult.ParentStatus -eq 'Pending') 'Shared dependency status was not read from the batch.'
    Assert-True ((@($scanResult.BatchJobNames | Sort-Object) -join ',') -eq 'Child,Parent') 'The request job names were not recorded for the dependency gate.'

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
    $expectedExternal = [ordered]@{ FreshReady = 'Ready'; FreshStale = 'Waiting'; SharedReady = 'Ready'; SharedFailed = 'Failed'; SharedRunning = 'Waiting'; LegacySuccess = 'Ready'; LegacyFailed = 'Failed'; Disabled = 'Ready'; Unknown = 'Failed' }
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
    Assert-True ($productionSelection.SelectedCount -eq 44) 'Published Full selection must contain the 44 enabled non-manual jobs.'
    Assert-True ($productionSelection.ExcludedCount -eq 4) 'Published Full selection must exclude the 4 disabled/manual jobs.'
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
    Assert-True ($LASTEXITCODE -eq 0) ("LaunchersByOrchestrator is not aligned with the jobs template: {0}" -f (@($launcherCheck) -join [Environment]::NewLine))

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
