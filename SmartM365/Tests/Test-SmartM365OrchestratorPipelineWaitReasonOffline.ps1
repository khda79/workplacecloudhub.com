<#
.SYNOPSIS
Offline tests of the wait reason of a pipeline request job that waits for its dependencies: the
reason names each blocking dependency with its server and start time, is written to the request
job status only when it changes, and is published in the heartbeat pending jobs.
.VERSION
1.0
.NOTES
Synthetic jobs, claims and request status files in a temporary folder. No tenant initialization,
collectors, scheduling, authentication or publication.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
function Get-FunctionText([string]$Name) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing orchestrator function: $Name" }
    $definition.Extent.Text
}

# The pipeline gate records the wait and forgets it when the job proceeds or is blocked.
$launchPhase = Get-FunctionText 'Invoke-LaunchPhase'
Check ($launchPhase -match 'Set-OrchestratorPipelineDependencyWait -JobName \$name -PipelineInfo \$pipelineInfo') 'The pipeline gate no longer records the dependency wait.'
Check (([regex]::Matches($launchPhase, '\$script:PipelineWaitState\.Remove\(\$name\)')).Count -ge 3) 'The pipeline wait is not cleared on launch, failed dependency and timeout.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-PipelineWait-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root | Out-Null
try {
    $functions = @('ConvertTo-StateTime', 'Write-OrchestratorRuntimeUpdateWarning', 'Get-OrchestratorExternalDependencyStatus', 'Get-OrchestratorDependencyActivityText',
        'Get-OrchestratorPipelineDependencyWaitText', 'Set-OrchestratorPipelineDependencyWait', 'Get-OrchestratorPipelineDependencyStatus', 'Get-OrchestratorPendingJobSnapshot')
    $stubs = @'
function Write-OrchestratorLog { param([string]$Message, [string]$Level = 'INFO') }
function Get-OrchestratorEffectiveDependencyMode { param($Job, $DependencyName) $script:Mode }
function Get-OrchestratorDependencyMaxAge { param($DependencyJob, $OverrideHours, $Now) 24 }
function Get-OrchestratorDependencyActivity { param($DependencyJob, $DownstreamOccurrence, $Now) $script:Activity }
function Get-OrchestratorDependencyFreshStatus { param($DependencyJob, $MaxAgeHours, $Now) $script:Fresh }
function Get-OrchestratorSharedDependencyStatus { param($Job, $Now) 'Waiting' }
function Test-JobAllowedOnServer { param($Job, [switch]$AllowManual) $true }
function Get-JobState { param($JobName) [pscustomobject]@{ PendingRetry = $null; LastScheduledOccurrence = '' } }
function Get-DueOccurrence { param($Job, $LastOccurrence, $Now) $null }
function Get-SmartM365OrchestratorPipelineJobStatusPath { param($SharedDataFolderPath, $BatchId, $JobName) Join-Path $SharedDataFolderPath "$BatchId-$JobName.json" }
function Set-OrchestratorPipelineJobStatus { param($BatchId, $JobName, $Status, [int]$Attempt = 0, $Detail = '') $script:StatusWrites.Add([pscustomobject]@{ BatchId = $BatchId; JobName = $JobName; Status = $Status; Detail = $Detail }) }
'@
    $text = (@($functions | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs
    $module = New-Module -ScriptBlock ([scriptblock]::Create($text))
    & $module {
        param($Root)
        $script:Settings = [pscustomobject]@{ SharedDataFolderPath = $Root; DistributedSchedulingEnabled = $true; MaxConcurrency = 4 }
        $script:StatusWrites = [Collections.Generic.List[object]]::new()
        $script:RuntimeUpdateWarnings = @{}
        $script:DependencyWaitReason = @{}
        $script:PipelineWaitState = @{}
        $script:DependencyWaitLogState = @{}
        $script:RunningJobs = @{}
        $script:MaintenanceHealthy = $true
        $script:MaintenanceControl = [pscustomobject]@{ Enabled = $false }
        $scope = [pscustomobject]@{ Name = 'Dep-Scope'; Enabled = $true; AssignmentMode = 'Elected'; ContinueOnError = $false }
        $inventory = [pscustomobject]@{ Name = 'Dep-Inventory'; Enabled = $true; AssignmentMode = 'Elected'; ContinueOnError = $false }
        $script:PrepareJob = [pscustomobject]@{ Name = 'Prepare'; Enabled = $true; AssignmentMode = 'Elected'; DependencyMaxAgeHours = 0; ConcurrencyKey = 'Prepare'; DependsOn = @('Dep-Scope', 'Dep-Inventory') }
        $script:Manifest = [pscustomobject]@{ JobsByName = @{ 'Dep-Scope' = $scope; 'Dep-Inventory' = $inventory; 'Prepare' = $script:PrepareJob }; OrderedJobs = @($script:PrepareJob) }
    } $root

    $now = [datetime]'2026-10-28 21:00'
    $claimedAt = '2026-10-28T17:16:00.0000000Z'
    $claimedLocal = ([datetimeoffset]::Parse($claimedAt)).LocalDateTime.ToString('yyyy-MM-dd HH:mm')

    # Dependency outside the request, running on another server: server and start time.
    $status = & $module {
        param($n, $c)
        $script:Mode = 'FreshSuccess'
        $script:Activity = [pscustomobject]@{ State = 'Waiting'; Detail = 'Running for occurrence 2026-10-28 18:15, may rewrite DATA-LAST during the run'; OwnerServer = 'SRV-B'; ClaimedAtUtc = $c; Status = 'Running' }
        Get-OrchestratorExternalDependencyStatus -Job $script:PrepareJob -DependencyName 'Dep-Scope' -Now $n
    } $now $claimedAt
    Check ($status -eq 'Waiting') 'A running external dependency must keep the request job waiting.'
    $reason = & $module { $script:DependencyWaitReason['Dep-Scope'] }
    Check ($reason -eq "Running on SRV-B since $claimedLocal") "The external reason must name the server and the start time: $reason"

    # Not running but not fresh: the freshness rule is named.
    & $module { param($n) $script:Activity = [pscustomobject]@{ State = 'Clear'; Detail = 'No active execution' }; $script:Fresh = 'Waiting'; [void](Get-OrchestratorExternalDependencyStatus -Job $script:PrepareJob -DependencyName 'Dep-Scope' -Now $n) } $now
    Check ((& $module { $script:DependencyWaitReason['Dep-Scope'] }) -eq 'no successful run in the last 24 h') 'A stale external dependency must name the freshness rule.'
    & $module { param($n) $script:Fresh = 'Ready'; [void](Get-OrchestratorExternalDependencyStatus -Job $script:PrepareJob -DependencyName 'Dep-Scope' -Now $n) } $now
    Check (-not (& $module { $script:DependencyWaitReason.ContainsKey('Dep-Scope') })) 'A ready dependency must not keep an old wait reason.'

    # Dependency of the same request: its request status and server.
    [IO.File]::WriteAllText((Join-Path $root 'B1-Dep-Inventory.json'), '{"Status":"Running","OwnerServer":"SRV-C"}')
    Check ((& $module { Get-OrchestratorPipelineDependencyStatus -BatchId 'B1' -JobName 'Dep-Inventory' }) -eq 'Running') 'The request dependency status must still be returned.'
    Check ((& $module { $script:DependencyWaitReason['Dep-Inventory'] }) -eq 'in request: Running on SRV-C') 'The request dependency reason must name its status and server.'

    # Wait text: each blocking dependency with its reason; unknown ones stay readable.
    & $module { param($n, $c) $script:Activity = [pscustomobject]@{ State = 'Waiting'; Detail = 'x'; OwnerServer = 'SRV-B'; ClaimedAtUtc = $c; Status = 'Running' }; [void](Get-OrchestratorExternalDependencyStatus -Job $script:PrepareJob -DependencyName 'Dep-Scope' -Now $n) } $now $claimedAt
    $waitText = & $module { Get-OrchestratorPipelineDependencyWaitText -BlockingDependencies @('Dep-Scope (outside request)', 'Dep-Inventory') }
    Check ($waitText -eq "Waiting for 2 dependencies: Dep-Scope (outside request: Running on SRV-B since $claimedLocal); Dep-Inventory (in request: Running on SRV-C).") "Unexpected wait text: $waitText"
    $single = & $module { Get-OrchestratorPipelineDependencyWaitText -BlockingDependencies @('Unknown-Job') }
    Check ($single -eq 'Waiting for 1 dependency: Unknown-Job (not finished).') "Unexpected single wait text: $single"

    # Request job status: written as Pending with the reason, once per change; a retry keeps its status.
    $info = @{ BatchId = 'B1' }
    & $module { param($i, $t, $n) Set-OrchestratorPipelineDependencyWait -JobName 'Prepare' -PipelineInfo $i -Text $t -Now $n; Set-OrchestratorPipelineDependencyWait -JobName 'Prepare' -PipelineInfo $i -Text $t -Now $n.AddMinutes(1) } $info $waitText $now
    $writes = & $module { @($script:StatusWrites) }
    Check ($writes.Count -eq 1 -and $writes[0].Status -eq 'Pending' -and $writes[0].Detail -eq $waitText -and $writes[0].BatchId -eq 'B1') 'The request job must receive the wait reason once.'
    & $module { param($i, $n) Set-OrchestratorPipelineDependencyWait -JobName 'Prepare' -PipelineInfo $i -Text 'Waiting for 1 dependency: Dep-Scope (x).' -Now $n.AddMinutes(5) } $info $now
    $state = & $module { $script:PipelineWaitState['Prepare'] }
    Check ((& $module { $script:StatusWrites.Count }) -eq 2 -and $state.FirstSeen -eq $now) 'A changed reason must be rewritten and keep the first wait time.'
    & $module { param($i, $n) Set-OrchestratorPipelineDependencyWait -JobName 'Prepare' -PipelineInfo $i -Text 'retry text' -IsRetry -Now $n } $info $now
    Check ((& $module { $script:StatusWrites.Count }) -eq 2 -and (& $module { $script:PipelineWaitState['Prepare'].Text }) -eq 'retry text') 'A scheduled retry must not be rewritten as Pending.'

    # Heartbeat: the waiting request job is listed with its request and reason.
    & $module { param($t) $script:PipelineWaitState['Prepare'].Text = $t; $script:PipelinePending = @{ 'Prepare' = @{ BatchId = 'B1' } } } $waitText
    $rows = @(& $module { param($n) Get-OrchestratorPendingJobSnapshot -Now $n } $now)
    Check ($rows.Count -eq 1 -and $rows[0].Name -eq 'Prepare' -and $rows[0].Reason -eq 'WaitingDependencies' -and $rows[0].Details -eq "Request B1: $waitText" -and $rows[0].FirstSeen -eq $now.ToString('o')) 'The heartbeat must list the waiting request job with its reason.'
    & $module { $script:MaintenanceControl = [pscustomobject]@{ Enabled = $true; Reason = 'Patching' } }
    $rows = @(& $module { param($n) Get-OrchestratorPendingJobSnapshot -Now $n } $now)
    Check ($rows.Count -eq 1 -and $rows[0].Reason -eq 'WaitingDependencies') 'Maintenance does not hide a request job that waits for its dependencies.'
    & $module { $script:PipelinePending = @{} }
    $rows = @(& $module { param($n) Get-OrchestratorPendingJobSnapshot -Now $n } $now)
    Check (@($rows | Where-Object Reason -eq 'WaitingDependencies').Count -eq 0) 'A finished or cancelled request must not stay in the heartbeat.'

    # Distributed module: the active claim gives its owner and claim time.
    Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
    $claims = Join-Path $root 'Claims'
    $occurrenceUtc = [datetime]::UtcNow.AddHours(-1)
    $folder = Join-Path $claims 'Dep-Scope'
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $claim = [ordered]@{ SchemaVersion = 1; JobName = 'Dep-Scope'; OccurrenceUtc = $occurrenceUtc.ToString('o'); OwnerServer = 'SRV-B'; Status = 'Running'; CreatedAtUtc = $claimedAt; UpdatedAtUtc = $claimedAt; SafeUntilUtc = [datetime]::UtcNow.AddHours(2).ToString('o') }
    [IO.File]::WriteAllText((Join-Path $folder ($occurrenceUtc.ToString('yyyyMMddTHHmmssfffZ') + '.json')), ($claim | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    $activity = Get-SmartM365OrchestratorDependencyActivity -ClaimsRootPath $claims -DependencyJob ([pscustomobject]@{ Name = 'Dep-Scope'; EstimatedDurationMinutes = 240 }) -DownstreamOccurrence (Get-Date) -Now (Get-Date)
    Check ($activity.State -eq 'Waiting' -and $activity.OwnerServer -eq 'SRV-B' -and ([datetimeoffset]::Parse($activity.ClaimedAtUtc, [Globalization.CultureInfo]::InvariantCulture)) -eq [datetimeoffset]::Parse($claimedAt, [Globalization.CultureInfo]::InvariantCulture) -and $activity.Status -eq 'Running') 'The dependency activity must return the owner and claim time of the active claim.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: pipeline wait reason, $($script:passed) offline checks."
