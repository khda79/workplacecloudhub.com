Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop

$script:PipelineTerminalStatuses = @(
    'Success',
    'CompletedWithWarnings',
    'Failed',
    'TimedOut',
    'Interrupted',
    'BlockedDependencyFailed',
    'BlockedDependencyTimeout',
    'Rejected',
    'Cancelled',
    'MissingStatus'
)

function Get-SmartM365OrchestratorPipelinePaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [string]$BatchId = ''
    )

    $root = Join-Path -Path $SharedDataFolderPath -ChildPath 'PipelineRuns'
    $result = [ordered]@{
        RootPath = $root
        SubmissionLockPath = Join-Path -Path $root -ChildPath 'PipelineRuns.lock'
    }
    if (-not [string]::IsNullOrWhiteSpace($BatchId)) {
        if ($BatchId -notmatch '^[A-Za-z0-9._-]+$' -or $BatchId -in @('.','..')) { throw "Invalid pipeline BatchId '$BatchId'." }
        $batchPath = Join-Path -Path $root -ChildPath $BatchId
        $result['BatchPath'] = $batchPath
        $result['ControlLockPath'] = Join-Path $batchPath 'Control.lock'
        $result['RequestPath'] = Resolve-SmartM365OwnedJsonPath -Path (Join-Path $batchPath 'request.json') -Owner 'Orchestrator Pipeline request' -Validate {
            param($document) if ($document.BatchId -ne $BatchId -or -not $document.PSObject.Properties['SelectedJobs']) { throw 'Pipeline request owner/batch mismatch.' }
        }
        $result['JobsFolderPath'] = Join-Path -Path $batchPath -ChildPath 'Jobs'
    }
    [pscustomobject]$result
}

function Get-SmartM365OrchestratorPipelineJobStatusPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][string]$BatchId,
        [Parameter(Mandatory)][string]$JobName
    )

    if ($JobName -notmatch '^[A-Za-z0-9._-]+$' -or $JobName -in @('.','..')) { throw "Invalid pipeline job name '$JobName'." }
    $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
    Resolve-SmartM365OwnedJsonPath -Path (Join-Path $paths.JobsFolderPath ($JobName + '.json')) -Owner 'Orchestrator Pipeline job' -Validate {
        param($document) if ($document.BatchId -ne $BatchId -or $document.JobName -ne $JobName) { throw 'Pipeline job owner/batch mismatch.' }
    }
}

function Enter-SmartM365OrchestratorPipelineLock {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSeconds = 15
    )

    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $deadline = [datetime]::UtcNow.AddSeconds([math]::Max(1, $TimeoutSeconds))
    do {
        try { return [IO.File]::Open($Path, 'CreateNew', 'ReadWrite', 'None') }
        catch [IO.IOException] { Start-Sleep -Milliseconds 200 }
    } while ([datetime]::UtcNow -lt $deadline)
    throw "Pipeline state is locked by another process: $Path"
}

function Move-SmartM365OrchestratorPipelineFileWithRetry {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestinationPath,
        [ValidateRange(1, 20)][int]$MaximumAttempts = 5,
        [ValidateRange(0, 5000)][int]$RetryDelayMilliseconds = 200,
        [scriptblock]$MoveOperation
    )

    if ($null -eq $MoveOperation) {
        $MoveOperation = { param($Source, $Destination) [IO.File]::Move($Source, $Destination, $true) }
    }

    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        try {
            & $MoveOperation $SourcePath $DestinationPath
            return
        }
        catch [IO.IOException] {
            if ($attempt -ge $MaximumAttempts) { throw }
        }
        catch [UnauthorizedAccessException] {
            if ($attempt -ge $MaximumAttempts) { throw }
        }
        if ($RetryDelayMilliseconds -gt 0) { Start-Sleep -Milliseconds $RetryDelayMilliseconds }
    }
}

function Write-SmartM365OrchestratorPipelineJsonAtomically {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Document,
        [switch]$CreateNew
    )
    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $content = ($Document | ConvertTo-Json -Depth 100) + [Environment]::NewLine
    $expected = if ($CreateNew) { 'ABSENT' } else { '' }
    $null = Write-SmartM365JsonBytesAtomically -Path $Path -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($content)) -ExpectedSHA256 $expected -Validate {
        param($value) if ($value -isnot [pscustomobject]) { throw 'Pipeline JSON must be an object.' }
    }
}

function Read-SmartM365OrchestratorPipelineJson {
    param([Parameter(Mandatory)][string]$Path)
    (Read-SmartM365JsonDocument -Path $Path).Document
}

function ConvertTo-SmartM365OrchestratorPipelineTimestampText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return ([datetime]$Value).ToUniversalTime().ToString('o') }
    [string]$Value
}

function Get-SmartM365OrchestratorPipelineSelection {
    [CmdletBinding(DefaultParameterSetName = 'Pipeline')]
    param(
        [Parameter(Mandatory)]$JobsDocument,
        [Parameter(ParameterSetName = 'Pipeline')]
        [ValidateSet('Full', 'AD', 'Exchange', 'Exchange2016', 'M365', 'Intune')]
        [string]$Pipeline = 'Full',
        [Parameter(Mandatory, ParameterSetName = 'Jobs')]
        [string[]]$JobName,
        [Parameter(ParameterSetName = 'Jobs')]
        [switch]$IncludeDependencies
    )

    if (-not $JobsDocument.PSObject.Properties['Jobs'] -or $null -eq $JobsDocument.Jobs) {
        throw "The jobs document must contain a 'Jobs' array."
    }

    $allJobs = @($JobsDocument.Jobs)
    $byName = @{}
    foreach ($job in $allJobs) {
        $name = ([string]$job.Name).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { throw 'A pipeline job has an empty Name.' }
        if ($byName.ContainsKey($name)) { throw "Duplicate pipeline job name '$name'." }
        $byName[$name] = $job
    }
    $isRunnable = {
        param($candidate)
        $enabled = $candidate.PSObject.Properties['Enabled'] -and [bool]$candidate.Enabled
        $mode = if ($candidate.PSObject.Properties['AssignmentMode'] -and $candidate.AssignmentMode) { [string]$candidate.AssignmentMode } else { 'Legacy' }
        $enabled -and $mode -ne 'Manual'
    }

    $selectionName = if ($PSCmdlet.ParameterSetName -eq 'Jobs') { 'Jobs' } else { $Pipeline }
    $selectedNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($PSCmdlet.ParameterSetName -eq 'Jobs') {
        foreach ($requestedName in @($JobName | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })) {
            if (-not $byName.ContainsKey($requestedName)) { throw "Job '$requestedName' is absent from the jobs manifest." }
            if (-not (& $isRunnable $byName[$requestedName])) { throw "Job '$requestedName' is disabled or manual and cannot be requested." }
            [void]$selectedNames.Add([string]$byName[$requestedName].Name)
        }
    }
    else {
        foreach ($job in $allJobs) {
            $groupMatches = $Pipeline -eq 'Full' -or ([string]$job.Group -eq $Pipeline)
            if ((& $isRunnable $job) -and $groupMatches) { [void]$selectedNames.Add([string]$job.Name) }
        }
    }

    if ($selectedNames.Count -eq 0) { throw "Pipeline '$selectionName' selected no enabled non-manual jobs." }

    # Group pipelines always add their dependencies; job requests only with -IncludeDependencies.
    # Disabled or manual dependencies are ignored, as the scheduled dependency gate ignores them.
    $expandDependencies = $PSCmdlet.ParameterSetName -eq 'Pipeline' -or $IncludeDependencies
    $addedDependencyNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $ignoredDependencies = [Collections.Generic.List[object]]::new()
    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($name in @($selectedNames)) {
            $job = $byName[$name]
            $dependencies = if ($job.PSObject.Properties['DependsOn']) { @($job.DependsOn | ForEach-Object { [string]$_ }) } else { @() }
            foreach ($dependencyName in $dependencies) {
                if (-not $byName.ContainsKey($dependencyName)) { throw "Job '$name' depends on unknown job '$dependencyName'." }
                if (-not (& $isRunnable $byName[$dependencyName])) {
                    if (-not @($ignoredDependencies | Where-Object { $_.Job -eq $name -and $_.Dependency -eq $dependencyName }).Count) {
                        $ignoredDependencies.Add([pscustomobject]@{ Job = $name; Dependency = $dependencyName; Reason = 'DisabledOrManual' })
                    }
                    continue
                }
                if (-not $expandDependencies) { continue }
                if ($selectedNames.Add($dependencyName)) {
                    [void]$addedDependencyNames.Add($dependencyName)
                    $changed = $true
                }
            }
        }
    }

    $selectedJobs = @(
        foreach ($job in $allJobs) {
            if (-not $selectedNames.Contains([string]$job.Name)) { continue }
            $dependsOn = if ($job.PSObject.Properties['DependsOn']) { @($job.DependsOn | ForEach-Object { [string]$_ }) } else { @() }
            [pscustomobject][ordered]@{
                Name = [string]$job.Name
                Group = [string]$job.Group
                AssignmentMode = if ($job.PSObject.Properties['AssignmentMode'] -and $job.AssignmentMode) { [string]$job.AssignmentMode } else { 'Legacy' }
                DependsOn = $dependsOn
                RequiredCapabilities = if ($job.PSObject.Properties['RequiredCapabilities']) { @($job.RequiredCapabilities | ForEach-Object { [string]$_ }) } else { @() }
                IncludedAsDependency = $addedDependencyNames.Contains([string]$job.Name)
                # Runnable dependencies outside the request: the orchestrator applies the job's scheduled rule.
                ExternalDependencies = @($dependsOn | Where-Object { -not $selectedNames.Contains($_) -and (& $isRunnable $byName[$_]) })
            }
        }
    )

    $excludedJobs = @(
        foreach ($job in $allJobs) {
            if ($selectedNames.Contains([string]$job.Name)) { continue }
            $enabled = $job.PSObject.Properties['Enabled'] -and [bool]$job.Enabled
            $assignmentMode = if ($job.PSObject.Properties['AssignmentMode'] -and $job.AssignmentMode) { [string]$job.AssignmentMode } else { 'Legacy' }
            $reason = if (-not $enabled) { 'Disabled' } elseif ($assignmentMode -eq 'Manual') { 'Manual' } elseif ($PSCmdlet.ParameterSetName -eq 'Jobs') { 'NotRequested' } else { 'OutsidePipelineGroup' }
            [pscustomobject]@{ Name = [string]$job.Name; Group = [string]$job.Group; Reason = $reason }
        }
    )

    [pscustomobject]@{
        Pipeline = $selectionName
        SelectedJobs = $selectedJobs
        ExcludedJobs = $excludedJobs
        IgnoredDependencies = @($ignoredDependencies)
        SelectedCount = $selectedJobs.Count
        ExcludedCount = $excludedJobs.Count
        AddedDependencyCount = $addedDependencyNames.Count
    }
}
function Get-SmartM365OrchestratorPipelineRunStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][string]$BatchId
    )

    $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
    if (-not (Test-Path -LiteralPath $paths.RequestPath -PathType Leaf)) { throw "Pipeline request not found: $($paths.RequestPath)" }
    $request = Read-SmartM365OrchestratorPipelineJson -Path $paths.RequestPath
    $jobs = @(
        foreach ($selectedJob in @($request.SelectedJobs)) {
            $name = if ($selectedJob -is [string]) { [string]$selectedJob } else { [string]$selectedJob.Name }
            $statusPath = Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId -JobName $name
            if (-not (Test-Path -LiteralPath $statusPath -PathType Leaf)) {
                [pscustomobject]@{ BatchId = $BatchId; JobName = $name; Status = 'MissingStatus'; OwnerServer = ''; UpdatedAtUtc = ''; Detail = ''; Path = $statusPath }
                continue
            }
            $status = Read-SmartM365OrchestratorPipelineJson -Path $statusPath
            [pscustomobject]@{
                BatchId = $BatchId
                JobName = $name
                Status = [string]$status.Status
                OwnerServer = if ($status.PSObject.Properties['OwnerServer']) { [string]$status.OwnerServer } else { '' }
                UpdatedAtUtc = if ($status.PSObject.Properties['UpdatedAtUtc']) { ConvertTo-SmartM365OrchestratorPipelineTimestampText $status.UpdatedAtUtc } else { '' }
                Detail = if ($status.PSObject.Properties['Detail']) { [string]$status.Detail } else { '' }
                Attempt = if ($status.PSObject.Properties['Attempt']) { [int]$status.Attempt } else { 0 }
                NotBeforeUtc = if ($status.PSObject.Properties['NotBeforeUtc']) { ConvertTo-SmartM365OrchestratorPipelineTimestampText $status.NotBeforeUtc } else { '' }
                Path = $statusPath
            }
        }
    )

    $pending = @($jobs | Where-Object { [string]$_.Status -notin $script:PipelineTerminalStatuses })
    $failed = @($jobs | Where-Object { [string]$_.Status -in @('Failed', 'TimedOut', 'Interrupted', 'BlockedDependencyFailed', 'BlockedDependencyTimeout', 'Rejected', 'MissingStatus') })
    $warnings = @($jobs | Where-Object { [string]$_.Status -eq 'CompletedWithWarnings' })
    $cancellationRequested = $null -ne $request.PSObject.Properties['Cancellation']
    $cancelled = @($jobs | Where-Object Status -eq 'Cancelled')
    $overallStatus = if ($pending.Count -gt 0) { if ($cancellationRequested) { 'Cancelling' } else { 'Running' } } elseif ($failed.Count -gt 0) { 'Failed' } elseif ($cancellationRequested) { 'Cancelled' } elseif ($cancelled.Count -gt 0) { 'Failed' } elseif ($warnings.Count -gt 0) { 'CompletedWithWarnings' } else { 'Success' }

    [pscustomobject]@{
        BatchId = $BatchId
        Pipeline = [string]$request.Pipeline
        Tenant = [string]$request.Tenant
        CreatedAtUtc = ConvertTo-SmartM365OrchestratorPipelineTimestampText $request.CreatedAtUtc
        ManifestHash = [string]$request.ManifestHash
        OverallStatus = $overallStatus
        IsTerminal = ($pending.Count -eq 0)
        TotalCount = $jobs.Count
        PendingCount = $pending.Count
        FailedCount = $failed.Count
        WarningCount = $warnings.Count
        CancelledCount = $cancelled.Count
        CancellationRequested = $cancellationRequested
        Jobs = $jobs
        Request = $request
        RequestPath = $paths.RequestPath
    }
}

function Get-SmartM365OrchestratorActivePipelineRuns {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)

    $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath
    if (-not (Test-Path -LiteralPath $paths.RootPath -PathType Container)) { return @() }
    @(
        foreach ($folder in @(Get-ChildItem -LiteralPath $paths.RootPath -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
            $requestPath = (Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath -BatchId $folder.Name).RequestPath
            if (-not (Test-Path -LiteralPath $requestPath -PathType Leaf)) { continue }
            try {
                $status = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $SharedDataFolderPath -BatchId $folder.Name
                if (-not $status.IsTerminal) { $status }
            }
            catch { throw "Invalid pipeline request '$($folder.Name)': $($_.Exception.Message)" }
        }
    )
}

function New-SmartM365OrchestratorPipelineRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][string]$Tenant,
        [Parameter(Mandatory)][ValidateSet('Full', 'AD', 'Exchange', 'Exchange2016', 'M365', 'Intune', 'Jobs')][string]$Pipeline,
        [Parameter(Mandatory)]$Selection,
        [Parameter(Mandatory)][string]$ManifestHash,
        [string]$RequestedBy = '',
        [string]$RequestedFrom = '',
        [int]$LockTimeoutSeconds = 15
    )

    if (@($Selection.SelectedJobs).Count -eq 0) { throw 'Cannot submit an empty pipeline request.' }
    $basePaths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath
    if (-not (Test-Path -LiteralPath $basePaths.RootPath)) { New-Item -ItemType Directory -Path $basePaths.RootPath -Force | Out-Null }
    $lock = Enter-SmartM365OrchestratorPipelineLock -Path $basePaths.SubmissionLockPath -TimeoutSeconds $LockTimeoutSeconds
    try {
        $active = @(Get-SmartM365OrchestratorActivePipelineRuns -SharedDataFolderPath $SharedDataFolderPath)
        if ($active.Count -gt 0) {
            throw "An active pipeline request already exists: $($active[0].BatchId) ($($active[0].OverallStatus))."
        }

        $createdUtc = [datetime]::UtcNow
        $batchId = '{0}-{1}' -f $createdUtc.ToString('yyyyMMddTHHmmssfffZ'), [guid]::NewGuid().ToString('N').Substring(0, 8)
        $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath -BatchId $batchId
        New-Item -ItemType Directory -Path $paths.JobsFolderPath -Force | Out-Null

        $request = [pscustomobject][ordered]@{
            SchemaVersion = 1
            BatchId = $batchId
            Tenant = $Tenant
            Pipeline = $Pipeline
            CreatedAtUtc = $createdUtc.ToString('o')
            OccurrenceUtc = $createdUtc.ToString('o')
            RequestedBy = $RequestedBy
            RequestedFrom = $RequestedFrom
            ManifestHash = $ManifestHash
            SelectedJobs = @($Selection.SelectedJobs)
        }

        foreach ($job in @($Selection.SelectedJobs)) {
            $statusPath = Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $SharedDataFolderPath -BatchId $batchId -JobName ([string]$job.Name)
            $jobStatus = [pscustomobject][ordered]@{
                SchemaVersion = 1
                BatchId = $batchId
                JobName = [string]$job.Name
                Status = 'Pending'
                Attempt = 0
                OwnerServer = ''
                CreatedAtUtc = $createdUtc.ToString('o')
                UpdatedAtUtc = $createdUtc.ToString('o')
                Detail = ''
                NotBeforeUtc = ''
            }
            Write-SmartM365OrchestratorPipelineJsonAtomically -Path $statusPath -Document $jobStatus -CreateNew
        }
        Write-SmartM365OrchestratorPipelineJsonAtomically -Path $paths.RequestPath -Document $request -CreateNew
        Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $SharedDataFolderPath -BatchId $batchId
    }
    finally {
        if ($null -ne $lock) { $lock.Dispose() }
        Remove-Item -LiteralPath $basePaths.SubmissionLockPath -Force -ErrorAction SilentlyContinue
    }
}

function Set-SmartM365OrchestratorPipelineJobStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][string]$BatchId,
        [Parameter(Mandatory)][string]$JobName,
        [Parameter(Mandatory)][ValidateSet('Pending', 'Starting', 'Running', 'RetryScheduled', 'Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted', 'BlockedDependencyFailed', 'BlockedDependencyTimeout', 'Rejected', 'Cancelled')][string]$Status,
        [int]$Attempt = 0,
        [string]$OwnerServer = '',
        [string]$Detail = '',
        [string]$NotBeforeUtc = '',
        [switch]$AllowRunningCancellation,
        [int]$LockTimeoutSeconds = 15
    )

    $statusPath = Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId -JobName $JobName
    if (-not (Test-Path -LiteralPath $statusPath -PathType Leaf)) { throw "Pipeline job status not found: $statusPath" }
    $lockPath = $statusPath + '.lock'
    $lock = Enter-SmartM365OrchestratorPipelineLock -Path $lockPath -TimeoutSeconds $LockTimeoutSeconds
    try {
        $document = Read-SmartM365OrchestratorPipelineJson -Path $statusPath
        # A cached scheduler view or late retry must never resurrect a cancelled job.
        if ([string]$document.Status -eq 'Cancelled') { return $document }
        if ($Status -eq 'Cancelled' -and [string]$document.Status -notin @('Pending', 'RetryScheduled') -and
            -not ($AllowRunningCancellation -and [string]$document.Status -in @('Starting', 'Running'))) { return $document }
        if ($Status -in @('Pending', 'Starting', 'RetryScheduled') -and
            (Test-SmartM365OrchestratorPipelineCancellation -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId)) {
            $Status = 'Cancelled'
            $Detail = 'Remaining attempt cancelled by the operator; previous attempt evidence is preserved.'
            $NotBeforeUtc = ''
        }
        $document.Status = $Status
        if ($document.PSObject.Properties['Attempt']) { $document.Attempt = $Attempt } else { $document | Add-Member -NotePropertyName Attempt -NotePropertyValue $Attempt }
        if ($document.PSObject.Properties['OwnerServer']) { $document.OwnerServer = $OwnerServer } else { $document | Add-Member -NotePropertyName OwnerServer -NotePropertyValue $OwnerServer }
        if ($document.PSObject.Properties['Detail']) { $document.Detail = $Detail } else { $document | Add-Member -NotePropertyName Detail -NotePropertyValue $Detail }
        if ($document.PSObject.Properties['NotBeforeUtc']) { $document.NotBeforeUtc = $NotBeforeUtc } else { $document | Add-Member -NotePropertyName NotBeforeUtc -NotePropertyValue $NotBeforeUtc }
        $document.UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
        Write-SmartM365OrchestratorPipelineJsonAtomically -Path $statusPath -Document $document
        $document
    }
    finally {
        if ($null -ne $lock) { $lock.Dispose() }
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
    }
}

function Test-SmartM365OrchestratorPipelineCancellation {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)][string]$BatchId)
    $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
    $request = Read-SmartM365OrchestratorPipelineJson -Path $paths.RequestPath
    return ($null -ne $request.PSObject.Properties['Cancellation'])
}

function Get-SmartM365OrchestratorPipelineCancellationReadiness {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [string]$Tenant = '')
    $cluster = (Read-SmartM365JsonDocument -Path (Join-Path $SharedDataFolderPath 'Config/Orchestrator-Cluster.json')).Document
    $servers = @($cluster.ExpectedOrchestratorServers)
    if ($servers.Count -eq 0) { throw 'Cancellation requires published expected orchestrator servers.' }
    $staleMinutes = if ($cluster.PSObject.Properties['PeerHeartbeatStaleMinutes']) { [int]$cluster.PeerHeartbeatStaleMinutes } else { 5 }
    if ($staleMinutes -lt 1) { throw 'Invalid heartbeat freshness threshold.' }
    foreach ($server in $servers) {
        if ([string]$server -notmatch '^[A-Za-z0-9_-]+$') { throw 'Invalid expected orchestrator server.' }
        $ready = $false
        $detail = ''
        $version = ''; $lifecycle = ''; $age = $null
        try {
            $heartbeat = (Read-SmartM365JsonDocument -Path (Join-Path $SharedDataFolderPath "$server/Orchestrator-Heartbeat.json")).Document
            $version = if ($heartbeat.PSObject.Properties['ScriptVersion']) { [string]$heartbeat.ScriptVersion } else { 'unknown' }
            $lifecycle = [string]$heartbeat.Lifecycle
            $timestamp = [datetimeoffset]::Parse((ConvertTo-SmartM365OrchestratorPipelineTimestampText $heartbeat.Timestamp), [Globalization.CultureInfo]::InvariantCulture)
            $age = ([datetimeoffset]::UtcNow - $timestamp.ToUniversalTime()).TotalMinutes
            $issues = @()
            if (-not $heartbeat.PSObject.Properties['PipelineCancellationProtocol'] -or [int]$heartbeat.PipelineCancellationProtocol -ne 1) {
                $issues += "Resident v$version does not advertise cancellation protocol 1; update/recycle safely."
            }
            if ($lifecycle -ne 'Running') { $issues += "Lifecycle=$lifecycle; wait for a Running heartbeat. Running collectors must not be stopped." }
            if ($age -lt -1) { $issues += 'Heartbeat is in the future; check clock synchronization.' }
            elseif ($age -gt $staleMinutes) { $issues += ("Heartbeat is stale ({0:F1} min; limit {1} min); wait for a fresh live heartbeat." -f $age, $staleMinutes) }
            if ($Tenant -and (-not $heartbeat.PSObject.Properties['Tenant'] -or [string]$heartbeat.Tenant -ne $Tenant)) {
                $issues += 'Heartbeat tenant does not match the cancellation tenant.'
            }
            $ready = $issues.Count -eq 0
            $detail = $issues -join ' '
        }
        catch { $detail = $_.Exception.Message }
        [pscustomobject]@{ Server = [string]$server; Ready = [bool]$ready; Detail = $detail; Version = $version; Lifecycle = $lifecycle; HeartbeatAgeMinutes = $age }
    }
}

function Enter-SmartM365OrchestratorPipelineLaunchGuard {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][string]$BatchId, [Parameter(Mandatory)][string]$JobName,
        [int]$LockTimeoutSeconds = 15)
    $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
    $lock = Enter-SmartM365OrchestratorPipelineLock -Path $paths.ControlLockPath -TimeoutSeconds $LockTimeoutSeconds
    try {
        $cancelled = Test-SmartM365OrchestratorPipelineCancellation -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
        $jobPath = Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId -JobName $JobName
        $state = Read-SmartM365OrchestratorPipelineJson -Path $jobPath
        if ($cancelled -or [string]$state.Status -notin @('Pending', 'RetryScheduled')) {
            $lock.Dispose(); $lock = $null
            Remove-Item -LiteralPath $paths.ControlLockPath -Force
            return [pscustomobject]@{ Allowed = $false; Lock = $null; LockPath = '' }
        }
        # Persist reservation before claims/process creation. A late status-write failure
        # cannot make a live collector look Pending and eligible for cancellation.
        $null = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId -JobName $JobName -Status Starting -Attempt ([int]$state.Attempt) -OwnerServer ([Environment]::MachineName)
        return [pscustomobject]@{
            Allowed = $true; Lock = $lock; LockPath = $paths.ControlLockPath
            SharedDataFolderPath = $SharedDataFolderPath; BatchId = $BatchId; JobName = $JobName
            OriginalState = $state; LaunchAttempted = $false
        }
    }
    catch {
        if ($null -ne $lock) { $lock.Dispose(); Remove-Item -LiteralPath $paths.ControlLockPath -Force -ErrorAction SilentlyContinue }
        throw
    }
}

function Exit-SmartM365OrchestratorPipelineLaunchGuard {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Guard)
    if ($null -ne $Guard.Lock) {
        try {
            if (-not $Guard.LaunchAttempted) {
                $path = Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $Guard.SharedDataFolderPath -BatchId $Guard.BatchId -JobName $Guard.JobName
                $state = Read-SmartM365OrchestratorPipelineJson -Path $path
                if ([string]$state.Status -eq 'Starting') {
                    $original = $Guard.OriginalState
                    $null = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $Guard.SharedDataFolderPath -BatchId $Guard.BatchId -JobName $Guard.JobName -Status $original.Status -Attempt $original.Attempt -OwnerServer $original.OwnerServer -Detail $original.Detail -NotBeforeUtc (ConvertTo-SmartM365OrchestratorPipelineTimestampText $original.NotBeforeUtc)
                }
            }
        }
        finally { $Guard.Lock.Dispose(); Remove-Item -LiteralPath $Guard.LockPath -Force -ErrorAction Stop }
    }
}

function Stop-SmartM365OrchestratorPipelineRequest {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][string]$BatchId, [Parameter(Mandatory)][string]$Tenant,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Reason,
        [string]$RequestedBy = ([Environment]::UserName),
        [string]$RequestedFrom = ([Environment]::MachineName),
        [int]$LockTimeoutSeconds = 15, [switch]$ValidateOnly)
    if ([string]::IsNullOrWhiteSpace($Reason)) { throw 'Cancellation reason is required.' }
    $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
    $before = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
    if ($before.Tenant -ne $Tenant) { throw 'Cancellation tenant does not match the request.' }
    if ($before.IsTerminal -and -not $before.CancellationRequested) { throw 'The request is already terminal; no cancellation was published.' }
    $readiness = @(Get-SmartM365OrchestratorPipelineCancellationReadiness -SharedDataFolderPath $SharedDataFolderPath -Tenant $Tenant)
    if (@($readiness | Where-Object { -not $_.Ready }).Count -gt 0) {
        throw "Cancellation unavailable: all expected residents must provide fresh Running heartbeats with cancellation protocol 1 for this tenant. $(@($readiness | Where-Object { -not $_.Ready } | ForEach-Object { "$($_.Server): $($_.Detail)" }) -join '; ')"
    }
    if ($ValidateOnly) { return $before }
    $submissionLock = Enter-SmartM365OrchestratorPipelineLock -Path $paths.SubmissionLockPath -TimeoutSeconds $LockTimeoutSeconds
    $controlLock = $null
    try {
        $controlLock = Enter-SmartM365OrchestratorPipelineLock -Path $paths.ControlLockPath -TimeoutSeconds $LockTimeoutSeconds
        $run = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
        if ($run.Tenant -ne $Tenant) { throw 'Cancellation tenant changed; no cancellation published.' }
        if ($run.IsTerminal -and -not $run.CancellationRequested) { throw 'The request completed before cancellation; results are unchanged.' }
        if (-not $run.CancellationRequested) {
            $run.Request | Add-Member -NotePropertyName Cancellation -NotePropertyValue ([pscustomobject]@{
                RequestedAtUtc = [datetime]::UtcNow.ToString('o'); RequestedBy = $RequestedBy
                RequestedFrom = $RequestedFrom; Reason = $Reason.Trim(); Protocol = 1
            })
            # Publish intent first. Interrupted writes can be completed by an idempotent retry.
            Write-SmartM365OrchestratorPipelineJsonAtomically -Path $paths.RequestPath -Document $run.Request
        }
        foreach ($job in @($run.Jobs | Where-Object Status -in @('Pending', 'RetryScheduled'))) {
            $null = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId -JobName $job.JobName -Status Cancelled -Attempt $job.Attempt -OwnerServer $job.OwnerServer -Detail "Cancelled: $($run.Request.Cancellation.Reason)"
        }
        Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
    }
    finally {
        if ($null -ne $controlLock) { $controlLock.Dispose(); Remove-Item -LiteralPath $paths.ControlLockPath -Force -ErrorAction SilentlyContinue }
        $submissionLock.Dispose(); Remove-Item -LiteralPath $paths.SubmissionLockPath -Force -ErrorAction SilentlyContinue
    }
}

Export-ModuleMember -Function @(
    'Test-SmartM365OrchestratorPipelineCancellation',
    'Get-SmartM365OrchestratorPipelineCancellationReadiness',
    'Enter-SmartM365OrchestratorPipelineLaunchGuard',
    'Exit-SmartM365OrchestratorPipelineLaunchGuard',
    'Stop-SmartM365OrchestratorPipelineRequest',
    'Get-SmartM365OrchestratorPipelinePaths',
    'Get-SmartM365OrchestratorPipelineJobStatusPath',
    'Get-SmartM365OrchestratorPipelineSelection',
    'Get-SmartM365OrchestratorPipelineRunStatus',
    'Get-SmartM365OrchestratorActivePipelineRuns',
    'New-SmartM365OrchestratorPipelineRequest',
    'Set-SmartM365OrchestratorPipelineJobStatus'
)

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC8O5rqRJrLN8cP
# cPRsL1qdkEaUyZsVvCLeeXJQ8BF4L6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDLOaibfJXYhmn7W8yEzBz0
# NxUghi782Ql3NVMfZlnbwTANBgkqhkiG9w0BAQEFAASCAYBIUu/A1Ox8mZd1eRcW
# IOFivSDAy2kpkUeyqBaFeTk9tF58Bx26pcj00MkfesBwOtC/J1or5tNU0zSyhbUJ
# +t1IRGq3qo6CmkxEixpxD5fDhLJSsSNVu6qct1H5tQpM2f0FxL8figQn3Yrf5YRC
# Y2fcV5wuMeHzZ8KSYQUgwHVUxeXGSADCGf9PNIIONYYCNN96zzF1xtLO3GeIp6iC
# MoGYYaUKQFahd9RNENwCNRH/PA8g2sHJgPYG2TZ8KYYQDiAYcZ+4bqhh1lIP8oL/
# +Xd0B8a+cpbB7Mqv/cOBKkyMqOr259GCKIIi3YlBzYrOkptC4FOMrq7oC4wtdCc5
# 77OGFkMBatKkQdbuBterrhXogD91wdOueb9SOI3TJ4IWAGlWcSO9su37DY2sP1nB
# Tei0DoMDqyMHyT1yU955nl/j7WL0DllBf1i7zgnm6lGN2BkFi7Smrptnzeo6s6gz
# 5Vqq7HbjkrPYavpd4prRbDGU+TgvrxWHTz7FZ8a5pUj4rFs=
# SIG # End signature block
