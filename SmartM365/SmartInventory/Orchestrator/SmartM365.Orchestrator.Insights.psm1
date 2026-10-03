# Read-only operational insights for the SmartM365 Orchestrator GUI: live server state, job
# health, dependency readiness, pipeline requests and job run requests. No WPF code here, so
# every function can be tested headless against a copy of the shared Orchestrator folder.
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Pipeline.psm1') -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Maintenance.psm1') -ErrorAction Stop

$script:InsightsTerminalSuccess = @('Success', 'CompletedWithWarnings')
$script:InsightsTerminalFailure = @('Failed', 'TimedOut', 'Interrupted')
# Same lists as SmartM365.Orchestrator.Pipeline.psm1 (Get-SmartM365OrchestratorPipelineRunStatus).
$script:InsightsPipelineFailure = @('Failed', 'TimedOut', 'Interrupted', 'BlockedDependencyFailed', 'BlockedDependencyTimeout', 'Rejected', 'MissingStatus')
$script:InsightsPipelineTerminal = @('Success', 'CompletedWithWarnings') + $script:InsightsPipelineFailure

function Get-InsightsProperty {
    param([AllowNull()]$Object, [Parameter(Mandatory)][string]$Name, [AllowNull()]$Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name) -and $null -ne $Object[$Name]) { return $Object[$Name] }; return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function ConvertTo-InsightsUtc {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return ([datetime]$Value).ToUniversalTime() }
    if ($Value -is [datetimeoffset]) { return ([datetimeoffset]$Value).UtcDateTime }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $parsed = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) { return $parsed.UtcDateTime }
    return $null
}

function Read-InsightsJson {
    param([Parameter(Mandatory)][string]$Path)
    $selected = Get-SmartM365JsonReadPath $Path -Optional
    if (-not $selected) { return $null }
    try { return (Read-SmartM365JsonDocument -Path $selected).Document } catch { return $null }
}

function Get-SmartM365OrchestratorScheduleOccurrences {
    <# Scheduled occurrences (local time) of a job schedule inside a window. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Schedule, [Parameter(Mandatory)][datetime]$From, [Parameter(Mandatory)][datetime]$To)
    $type = [string](Get-InsightsProperty $Schedule 'Type' 'Daily')
    $days = @(Get-InsightsProperty $Schedule 'DaysOfWeek' @() | ForEach-Object { [string]$_ })
    $times = @(Get-InsightsProperty $Schedule 'Times' @() | ForEach-Object { [string]$_ })
    $result = [Collections.Generic.List[datetime]]::new()
    for ($day = $From.Date; $day -le $To.Date; $day = $day.AddDays(1)) {
        if ($type -eq 'Weekly' -and [string]$day.DayOfWeek -notin $days) { continue }
        foreach ($timeText in $times) {
            $time = [timespan]::Zero
            if (-not [timespan]::TryParseExact($timeText, 'hh\:mm', [Globalization.CultureInfo]::InvariantCulture, [ref]$time)) { continue }
            $occurrence = $day.Add($time)
            if ($occurrence -ge $From -and $occurrence -le $To) { $result.Add($occurrence) }
        }
    }
    @($result | Sort-Object)
}

function Get-SmartM365OrchestratorExpectedMaxAgeHours {
    <# Same rule as the orchestrator FreshSuccess gate: longest schedule gap + 2 h; an override is a floor. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Job, [int]$OverrideHours = 0, [datetime]$Now = (Get-Date))
    $occurrences = @(Get-SmartM365OrchestratorScheduleOccurrences -Schedule $Job.Schedule -From $Now.Date.AddDays(-15) -To $Now)
    $maxGap = 24.0
    if ($occurrences.Count -ge 2) {
        $maxGap = 0.0
        for ($i = 1; $i -lt $occurrences.Count; $i++) { $gap = ($occurrences[$i] - $occurrences[$i - 1]).TotalHours; if ($gap -gt $maxGap) { $maxGap = $gap } }
    }
    [math]::Max($maxGap + 2.0, [double][math]::Max(0, $OverrideHours))
}

function Get-SmartM365OrchestratorRecentRuns {
    <# Job run rows of every server; only the daily CSV files inside the window are read. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [int]$Days = 21, [datetime]$Now = (Get-Date))
    $from = $Now.Date.AddDays(-1 * $Days)
    $rows = [Collections.Generic.List[object]]::new()
    if (-not (Test-Path -LiteralPath $SharedDataFolderPath)) { return @() }
    foreach ($serverFolder in @(Get-ChildItem -LiteralPath $SharedDataFolderPath -Directory -ErrorAction SilentlyContinue)) {
        $jobRuns = Join-Path $serverFolder.FullName 'JobRuns'
        if (-not (Test-Path -LiteralPath $jobRuns)) { continue }
        foreach ($csv in @(Get-ChildItem -LiteralPath $jobRuns -Filter 'Orchestrator_JobRuns_*.csv' -File -ErrorAction SilentlyContinue)) {
            $match = [regex]::Match($csv.Name, '_(\d{8})')
            if ($match.Success) {
                $fileDate = [datetime]::ParseExact($match.Groups[1].Value, 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture)
                if ($fileDate -lt $from.AddDays(-1)) { continue }
            }
            foreach ($row in @(Import-Csv -LiteralPath $csv.FullName -ErrorAction SilentlyContinue)) {
                $start = [datetime]::MinValue
                if (-not [datetime]::TryParse([string]$row.StartTime, [ref]$start) -or $start -lt $from) { continue }
                $end = [datetime]::MinValue
                $endValue = if ([datetime]::TryParse([string]$row.EndTime, [ref]$end)) { $end } else { $null }
                $rows.Add([pscustomobject]@{
                    Server = $serverFolder.Name; JobName = [string]$row.JobName; StartTime = $start; EndTime = $endValue
                    DurationSec = [double](Get-InsightsProperty $row 'DurationSec' 0); Status = [string]$row.Status; LogPath = [string]$row.LogPath
                })
            }
        }
    }
    @($rows | Sort-Object StartTime -Descending)
}

function Get-SmartM365OrchestratorJobHealth {
    <# One health row per job from the recent runs of all servers. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$JobsDocument, [AllowEmptyCollection()][object[]]$Runs = @(), [datetime]$Now = (Get-Date))
    $byJob = @{}
    foreach ($run in @($Runs)) { if (-not $byJob.ContainsKey($run.JobName)) { $byJob[$run.JobName] = [Collections.Generic.List[object]]::new() }; $byJob[$run.JobName].Add($run) }
    foreach ($job in @($JobsDocument.Jobs)) {
        $name = [string]$job.Name
        $jobRuns = if ($byJob.ContainsKey($name)) { @($byJob[$name] | Sort-Object StartTime -Descending) } else { @() }
        $final = @($jobRuns | Where-Object { $_.Status -ne 'Retried' })
        $successes = @($final | Where-Object { $_.Status -in $script:InsightsTerminalSuccess })
        $last = if ($final.Count) { $final[0] } else { $null }
        $lastSuccess = if ($successes.Count) { $successes[0] } else { $null }
        $lastSuccessEnd = if ($null -ne $lastSuccess) { if ($null -ne $lastSuccess.EndTime) { [datetime]$lastSuccess.EndTime } else { [datetime]$lastSuccess.StartTime } } else { $null }
        $avgMinutes = if ($successes.Count) { [math]::Round((@($successes | Select-Object -First 10 | ForEach-Object DurationSec) | Measure-Object -Average).Average / 60, 1) } else { $null }
        $mode = [string](Get-InsightsProperty $job 'AssignmentMode' 'Legacy')
        $expected = [math]::Round((Get-SmartM365OrchestratorExpectedMaxAgeHours -Job $job -Now $Now) + $(if ($avgMinutes) { $avgMinutes / 60 } else { 0 }), 1)
        $ageHours = if ($null -ne $lastSuccessEnd) { [math]::Round(($Now - $lastSuccessEnd).TotalHours, 1) } else { $null }
        $health = if (-not [bool](Get-InsightsProperty $job 'Enabled' $false)) { 'Disabled' }
        elseif ($mode -eq 'Manual') { 'Manual' }
        elseif ($null -ne $last -and $last.Status -in $script:InsightsTerminalFailure) { 'Failing' }
        elseif ($null -eq $lastSuccessEnd) { 'NoRecentSuccess' }
        elseif ($ageHours -gt $expected) { 'Stale' }
        else { 'OK' }
        [pscustomobject]@{
            Name = $name; Health = $health
            LastStatus = if ($null -ne $last) { $last.Status } else { '' }
            LastRun = if ($null -ne $last) { ([datetime]$last.StartTime).ToString('yyyy-MM-dd HH:mm') } else { '' }
            LastSuccess = if ($null -ne $lastSuccessEnd) { $lastSuccessEnd.ToString('yyyy-MM-dd HH:mm') } else { '' }
            SuccessAgeHours = $ageHours; ExpectedMaxAgeHours = $expected; AverageDurationMinutes = $avgMinutes
        }
    }
}

function Get-SmartM365OrchestratorDependents {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$JobsDocument, [Parameter(Mandatory)][string]$JobName, [switch]$EnabledOnly)
    @($JobsDocument.Jobs | Where-Object {
        (@(Get-InsightsProperty $_ 'DependsOn' @()) -contains $JobName) -and (-not $EnabledOnly -or [bool](Get-InsightsProperty $_ 'Enabled' $false))
    } | ForEach-Object { [string]$_.Name } | Sort-Object)
}

function Get-InsightsClaimFiles {
    param([string]$ClaimsRoot, [string]$JobName)
    $folder = Join-Path $ClaimsRoot $JobName
    if (-not (Test-Path -LiteralPath $folder)) { return @() }
    @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{8}T\d{9}Z\.json(\.txt)?$' } | Sort-Object Name -Descending)
}

function ConvertTo-InsightsClaim {
    param([Parameter(Mandatory)][IO.FileInfo]$File)
    $document = Read-InsightsJson -Path $File.FullName
    if ($null -eq $document) { return $null }
    [pscustomobject]@{
        Occurrence = [datetime]::ParseExact($File.Name.Substring(0, 19), 'yyyyMMddTHHmmssfffZ', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal)
        Status = [string](Get-InsightsProperty $document 'Status' '')
        UpdatedUtc = ConvertTo-InsightsUtc (Get-InsightsProperty $document 'UpdatedAtUtc' $null)
        OwnerServer = [string](Get-InsightsProperty $document 'OwnerServer' '')
    }
}

function Get-InsightsLatestSuccessClaim {
    # Newest-first; stops at the first success so a readiness check reads only a few files.
    param([string]$ClaimsRoot, [string]$JobName, [int]$MaximumFiles = 60)
    foreach ($file in @(Get-InsightsClaimFiles -ClaimsRoot $ClaimsRoot -JobName $JobName | Select-Object -First $MaximumFiles)) {
        $claim = ConvertTo-InsightsClaim -File $file
        if ($null -ne $claim -and $claim.Status -in $script:InsightsTerminalSuccess -and $null -ne $claim.UpdatedUtc) { return $claim }
    }
    return $null
}

function Get-InsightsOccurrenceClaim {
    param([string]$ClaimsRoot, [string]$JobName, [datetime]$OccurrenceUtc)
    $stem = $OccurrenceUtc.ToString('yyyyMMddTHHmmssfffZ', [Globalization.CultureInfo]::InvariantCulture)
    foreach ($name in @("$stem.json.txt", "$stem.json")) {
        $path = Join-Path (Join-Path $ClaimsRoot $JobName) $name
        if (Test-Path -LiteralPath $path -PathType Leaf) { return ConvertTo-InsightsClaim -File (Get-Item -LiteralPath $path) }
    }
    return $null
}
function Get-SmartM365OrchestratorDependencyReadiness {
    <# Why a job waits: each dependency evaluated with the job's scheduled dependency rule. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)]$JobsDocument, [Parameter(Mandatory)][string]$JobName,
        [AllowEmptyCollection()][object[]]$Runs = @(), [datetime]$Now = (Get-Date)
    )
    $job = @($JobsDocument.Jobs | Where-Object { [string]$_.Name -eq $JobName })[0]
    if ($null -eq $job) { throw "Job '$JobName' is absent from the jobs document." }
    $claimsRoot = Join-Path $SharedDataFolderPath 'Election\Claims'
    $mode = [string](Get-InsightsProperty $job 'DependencyMode' 'LatestOccurrence')
    $override = [int](Get-InsightsProperty $job 'DependencyMaxAgeHours' 0)
    foreach ($dependencyName in @(Get-InsightsProperty $job 'DependsOn' @() | ForEach-Object { [string]$_ })) {
        $dependency = @($JobsDocument.Jobs | Where-Object { [string]$_.Name -eq $dependencyName })[0]
        $row = [ordered]@{ Dependency = $dependencyName; Rule = $mode; State = ''; LastSuccess = ''; AgeHours = $null; MaxAgeHours = $null; Detail = '' }
        if ($null -eq $dependency) { $row.State = 'Failed'; $row.Detail = 'Unknown job'; [pscustomobject]$row; continue }
        if (-not [bool](Get-InsightsProperty $dependency 'Enabled' $false) -or [string](Get-InsightsProperty $dependency 'AssignmentMode' 'Legacy') -eq 'Manual') {
            $row.State = 'Ignored'; $row.Detail = 'Disabled or manual: ignored by the dependency gate'; [pscustomobject]$row; continue
        }
        $successClaim = Get-InsightsLatestSuccessClaim -ClaimsRoot $claimsRoot -JobName $dependencyName
        $successUtc = if ($null -ne $successClaim) { $successClaim.UpdatedUtc } else { $null }
        if ($null -eq $successUtc) {
            $runSuccess = @($Runs | Where-Object { $_.JobName -eq $dependencyName -and $_.Status -in $script:InsightsTerminalSuccess } | Select-Object -First 1)
            if ($runSuccess.Count) { $successUtc = ([datetime]$(if ($null -ne $runSuccess[0].EndTime) { $runSuccess[0].EndTime } else { $runSuccess[0].StartTime })).ToUniversalTime() }
        }
        if ($null -ne $successUtc) {
            $row.LastSuccess = $successUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
            $row.AgeHours = [math]::Round(($Now.ToUniversalTime() - $successUtc).TotalHours, 1)
        }
        if ($mode -eq 'FreshSuccess') {
            $row.MaxAgeHours = [math]::Round((Get-SmartM365OrchestratorExpectedMaxAgeHours -Job $dependency -OverrideHours $override -Now $Now), 1)
            if ($null -ne $row.AgeHours -and $row.AgeHours -le $row.MaxAgeHours) { $row.State = 'Ready' }
            else { $row.State = 'Waiting'; $row.Detail = if ($null -eq $row.AgeHours) { 'No recorded success' } else { "Last success older than $($row.MaxAgeHours) h" } }
        }
        else {
            $expected = @(Get-SmartM365OrchestratorScheduleOccurrences -Schedule $dependency.Schedule -From $Now.AddDays(-8) -To $Now | Select-Object -Last 1)
            if (-not $expected.Count) { $row.State = 'Waiting'; $row.Detail = 'No scheduled occurrence in the last 8 days' }
            else {
                $expectedUtc = $expected[0].ToUniversalTime()
                $claim = Get-InsightsOccurrenceClaim -ClaimsRoot $claimsRoot -JobName $dependencyName -OccurrenceUtc $expectedUtc
                $status = if ($null -ne $claim) { $claim.Status } else { '' }
                $row.Detail = "Latest occurrence $($expected[0].ToString('yyyy-MM-dd HH:mm')): $(if ($status) { $status } else { 'not started' })"
                $row.State = if ($status -in $script:InsightsTerminalSuccess) { 'Ready' } elseif ($status -in $script:InsightsTerminalFailure) { 'Failed' } else { 'Waiting' }
            }
        }
        [pscustomobject]$row
    }
}

function Get-SmartM365OrchestratorOperations {
    <# Live view from the shared heartbeats, states and orchestrator mail copies. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)]$ClusterDocument,
        [string]$MailFolderPath = '', [int]$MailHours = 24, [datetime]$Now = (Get-Date)
    )
    $servers = [Collections.Generic.List[object]]::new(); $running = [Collections.Generic.List[object]]::new()
    $pending = [Collections.Generic.List[object]]::new(); $incidents = [Collections.Generic.List[object]]::new()
    $staleMinutes = [double](Get-InsightsProperty $ClusterDocument 'PeerHeartbeatStaleMinutes' 5)
    $maintenance = $null; $maintenanceError = ''; $maintenanceServers = @()
    try {
        $maintenance = Get-SmartM365OrchestratorMaintenanceState $SharedDataFolderPath
        $maintenanceServers = @(Get-SmartM365OrchestratorMaintenanceReadiness $SharedDataFolderPath $ClusterDocument $maintenance -Now $Now)
    }
    catch { $maintenanceError = $_.Exception.Message }
    foreach ($serverName in @($ClusterDocument.ExpectedOrchestratorServers | ForEach-Object { [string]$_ } | Sort-Object -Unique)) {
        $folder = Join-Path $SharedDataFolderPath $serverName
        $heartbeat = Read-InsightsJson -Path (Join-Path $folder 'Orchestrator-Heartbeat.json')
        $timestamp = ConvertTo-InsightsUtc (Get-InsightsProperty $heartbeat 'Timestamp' $null)
        $age = if ($null -ne $timestamp) { [math]::Round(($Now.ToUniversalTime() - $timestamp).TotalMinutes, 1) } else { $null }
        $lifecycle = [string](Get-InsightsProperty $heartbeat 'Lifecycle' $(if ($null -ne $heartbeat) { 'Running' } else { '' }))
        $deadline = ConvertTo-InsightsUtc (Get-InsightsProperty $heartbeat 'LifetimeDeadline' $null)
        $jobsRunning = @(Get-InsightsProperty $heartbeat 'RunningJobs' @())
        $jobsPending = @(Get-InsightsProperty $heartbeat 'PendingJobs' @())
        $state = if ($null -eq $heartbeat) { 'No heartbeat' } elseif ($lifecycle -eq 'Recycling') { 'Recycling' } elseif ($lifecycle -eq 'Starting') { 'Starting' } elseif ($age -gt $staleMinutes) { 'Stale' } else { 'Online' }
        $recycleIn = ''
        if ($null -ne $deadline) {
            $remaining = $deadline - $Now.ToUniversalTime()
            $recycleIn = if ($remaining.TotalMinutes -gt 0) { '{0:0}h{1:00}' -f [math]::Floor($remaining.TotalHours), $remaining.Minutes } else { 'due' }
        }
        $servers.Add([pscustomobject]@{
            Server = $serverName; State = $state; Lifecycle = $lifecycle; HeartbeatAgeMinutes = $age
            Version = [string](Get-InsightsProperty $heartbeat 'ScriptVersion' ''); Pid = [string](Get-InsightsProperty $heartbeat 'Pid' '')
            Running = $jobsRunning.Count; Pending = $jobsPending.Count
            RecycleIn = $recycleIn
            Maintenance = if ($maintenanceError) { 'Control unavailable' } else { [string](@($maintenanceServers | Where-Object Server -eq $serverName)[0].Status) }
            StatePersistence = if ([bool](Get-InsightsProperty $heartbeat 'StatePersistenceHealthy' $true)) { 'OK' } else { 'Launches paused: ' + [string](Get-InsightsProperty $heartbeat 'StatePersistenceLastError' '') }
        })
        foreach ($job in $jobsRunning) {
            $start = ConvertTo-InsightsUtc (Get-InsightsProperty $job 'StartTime' $null)
            $running.Add([pscustomobject]@{
                Server = $serverName; Job = [string](Get-InsightsProperty $job 'Name' ''); Pid = [string](Get-InsightsProperty $job 'Pid' '')
                Started = if ($null -ne $start) { $start.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
                DurationMinutes = if ($null -ne $start) { [math]::Round(($Now.ToUniversalTime() - $start).TotalMinutes, 0) } else { $null }
                Occurrence = [string](Get-InsightsProperty $job 'ScheduledOccurrence' '')
            })
        }
        foreach ($job in $jobsPending) {
            $firstSeen = ConvertTo-InsightsUtc (Get-InsightsProperty $job 'FirstSeen' $null)
            $pending.Add([pscustomobject]@{
                Server = $serverName; Job = [string](Get-InsightsProperty $job 'Name' ''); Reason = [string](Get-InsightsProperty $job 'Reason' '')
                WaitingMinutes = if ($null -ne $firstSeen) { [math]::Round(($Now.ToUniversalTime() - $firstSeen).TotalMinutes, 0) } else { $null }
                Occurrence = [string](Get-InsightsProperty $job 'ScheduledOccurrence' ''); Details = [string](Get-InsightsProperty $job 'Details' '')
            })
        }
        $orchestratorState = Read-InsightsJson -Path (Join-Path $folder 'Orchestrator-State.json')
        $summary = [string](Get-InsightsProperty (Get-InsightsProperty $orchestratorState 'PeerMonitoring' $null) 'ActiveIssueSummary' '')
        if (-not [string]::IsNullOrWhiteSpace($summary)) {
            foreach ($line in @($summary -split '\r?\n|;\s*' | Where-Object { $_.Trim() })) { $incidents.Add([pscustomobject]@{ ObservedBy = $serverName; Issue = $line.Trim() }) }
        }
    }
    $mails = [Collections.Generic.List[object]]::new()
    if (-not [string]::IsNullOrWhiteSpace($MailFolderPath) -and (Test-Path -LiteralPath $MailFolderPath)) {
        $since = $Now.AddHours(-1 * $MailHours)
        foreach ($file in @(Get-ChildItem -LiteralPath $MailFolderPath -Recurse -File -Filter '*.htm*' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $since })) {
            $match = [regex]::Match($file.BaseName, '^(\d{4}-\d{2}-\d{2})_(\d{2})-(\d{2})-(\d{2})_(?<Subject>.*)$')
            $subject = if ($match.Success) { ($match.Groups['Subject'].Value -replace '^.*\]\s*-\s*', '') } else { $file.BaseName }
            $mails.Add([pscustomobject]@{ Time = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm'); Server = $file.Directory.Name; Subject = $subject; Path = $file.FullName })
        }
    }
    [pscustomobject]@{
        Servers = @($servers); Running = @($running | Sort-Object Server, Job); Pending = @($pending | Sort-Object Server, Job)
        Incidents = @($incidents); Mails = @($mails | Sort-Object Time -Descending); RefreshedAt = $Now
        Maintenance = $maintenance; MaintenanceError = $maintenanceError; MaintenanceServers = $maintenanceServers
    }
}

function Get-SmartM365OrchestratorRecentPipelineRuns {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [int]$Count = 20)
    $root = Join-Path $SharedDataFolderPath 'PipelineRuns'
    if (-not (Test-Path -LiteralPath $root)) { return @() }
    # Direct read-only reads: the pipeline module resolves and validates request.json once per
    # job status, which costs about one second per batch on a shared folder.
    foreach ($folder in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First $Count)) {
        $request = Read-InsightsJson -Path (Join-Path $folder.FullName 'request.json')
        if ($null -eq $request -or [string](Get-InsightsProperty $request 'BatchId' '') -ne $folder.Name) {
            [pscustomobject]@{ BatchId = $folder.Name; Selection = ''; Created = ''; Status = 'Unreadable'; Jobs = 0; Pending = 0; Failed = 0; RequestedBy = ''; JobRows = @() }
            continue
        }
        $jobRows = @(foreach ($selectedJob in @(Get-InsightsProperty $request 'SelectedJobs' @())) {
            $name = if ($selectedJob -is [string]) { [string]$selectedJob } else { [string](Get-InsightsProperty $selectedJob 'Name' '') }
            $jobStatus = Read-InsightsJson -Path (Join-Path (Join-Path $folder.FullName 'Jobs') ($name + '.json'))
            $updated = ConvertTo-InsightsUtc (Get-InsightsProperty $jobStatus 'UpdatedAtUtc' $null)
            [pscustomobject]@{
                JobName = $name; Status = if ($null -ne $jobStatus) { [string](Get-InsightsProperty $jobStatus 'Status' '') } else { 'MissingStatus' }
                OwnerServer = [string](Get-InsightsProperty $jobStatus 'OwnerServer' '')
                UpdatedAtUtc = if ($null -ne $updated) { $updated.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
                Detail = [string](Get-InsightsProperty $jobStatus 'Detail' '')
            }
        })
        $pendingCount = @($jobRows | Where-Object { $_.Status -notin $script:InsightsPipelineTerminal }).Count
        $failedCount = @($jobRows | Where-Object { $_.Status -in $script:InsightsPipelineFailure }).Count
        $warningCount = @($jobRows | Where-Object { $_.Status -eq 'CompletedWithWarnings' }).Count
        $created = ConvertTo-InsightsUtc (Get-InsightsProperty $request 'CreatedAtUtc' $null)
        [pscustomobject]@{
            BatchId = $folder.Name; Selection = [string](Get-InsightsProperty $request 'Pipeline' '')
            Created = if ($null -ne $created) { $created.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
            Status = if ($pendingCount -gt 0) { 'Running' } elseif ($failedCount -gt 0) { 'Failed' } elseif ($warningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
            Jobs = $jobRows.Count; Pending = $pendingCount; Failed = $failedCount
            RequestedBy = [string](Get-InsightsProperty $request 'RequestedBy' ''); JobRows = $jobRows
        }
    }
}

function New-SmartM365OrchestratorJobRunRequest {
    <# Same contract as SmartM365-Inventory-Pipeline.ps1 -Job ... -Collect -NoWait, on the published configuration. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)][string]$JobsPath, [Parameter(Mandatory)][string[]]$JobName,
        [Parameter(Mandatory)][string]$Tenant, [switch]$IncludeDependencies
    )
    $jobsDocument = (Read-SmartM365JsonDocument -Path $JobsPath).Document
    $selection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $jobsDocument -JobName $JobName -IncludeDependencies:$IncludeDependencies
    $plan = Read-InsightsJson -Path (Join-Path $SharedDataFolderPath 'Election\Orchestrator-ElectionPlan.json')
    $missing = @(foreach ($selected in @($selection.SelectedJobs | Where-Object AssignmentMode -eq 'Elected')) {
        if (-not @(Get-InsightsProperty $plan 'Assignments' @() | Where-Object { [string]$_.JobName -eq [string]$selected.Name -and [string]$_.OwnerServer }).Count) { [string]$selected.Name }
    })
    if ($missing.Count) { throw "No elected owner in the current shared plan for: $($missing -join ', ')." }
    $hash = (Get-FileHash -LiteralPath $JobsPath -Algorithm SHA256).Hash
    New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $SharedDataFolderPath -Tenant $Tenant -Pipeline Jobs -Selection $selection -ManifestHash $hash -RequestedBy ([Environment]::UserName) -RequestedFrom ([Environment]::MachineName)
}

Export-ModuleMember -Function @(
    'Get-SmartM365OrchestratorScheduleOccurrences', 'Get-SmartM365OrchestratorExpectedMaxAgeHours', 'Get-SmartM365OrchestratorRecentRuns',
    'Get-SmartM365OrchestratorJobHealth', 'Get-SmartM365OrchestratorDependents', 'Get-SmartM365OrchestratorDependencyReadiness',
    'Get-SmartM365OrchestratorOperations', 'Get-SmartM365OrchestratorRecentPipelineRuns', 'New-SmartM365OrchestratorJobRunRequest'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZ414hZoffH6e5
# zDPQyjqsPLazlRAGI/mIyfvqpzylYaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGdzglAy2/TwmeHSFtWgfRgepOiIDBYO/eGGpSeuRSRqMA0GCSqG
# SIb3DQEBAQUABIIBgI6ITIgd3I95QOWxsPQDR7EfqLHZpC+fC5bGZXU1rqcmH6LP
# MAIyUdat2zpuB7vK5XJh85qqkPXhp3TYTG8biN/3gxdeogXLViC5kP/Pk8z4CQF2
# yNcF/0mdzqh+DxE4CsilTrRGJOOJ2QBXJajwUj7wPXZG27xaaY/L/cRfX4WJt2uB
# IX2LAV3aXV2wxSEXMNRj63P3QozR2Wcqqwf9mDgv0R+d/EhHulRnwYvxfqyidG8e
# SNocPpDyPlqjiRy0ytyvXOQ2n4z8LFGUxS2p+7J2sQvgi6yX6sY2wAdRUPOxK5lz
# UqAuaNSmv+D0uOiY7if7yjV0swmV5NIJ0sux1RyoqQ0vVPYSLeRGX2jlvqbN1DFN
# m0s+4espTajlCdWCkiNg5RNJEvPlG386DE+DWPDr2noLrhnxe/2RasZq/bB3IEkx
# d95nITQBoL2BIa39vhvtMr/kkpnaEEAMq+bFk5pxer0zuvMPWJyk1un0H3GiJVNM
# dO5bqBV5HLZSJuWfiqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMyMDEy
# MTRaMC8GCSqGSIb3DQEJBDEiBCB+/0T4vRfGzUO368FcEjmRTjOZ8hzhpF9x7ir3
# 3Igz/TANBgkqhkiG9w0BAQEFAASCAgBsTDBinlWYHAQPw1pXoMoy0poCefq6XbBc
# yQwDZ864cVa5AkT7VEswSKnyxPG0kvPCBQQrdV1Ro767gn36z3DxtnDWiIckgJdx
# 9HosbXFmgjI1IEqMEtKnrfCPVsFN7CKkBy+0CzKeU3VUqxkKzo22YHR8irMQb1E8
# PSXnBrsCG/+lIy6G/PKtUFCGCzl40wZKfJWYnX6RACZDPWp+HsxEkUXdTGla6FU+
# qmaWgcYQdLAVFkqOUwrcCmDDjsvxOiI8QLTQ1I9ajS0ck7IpOB7U0s/d9AP0Wazk
# TTaGTQFaP/cO2/khmqQhtPZAXMHEitVMB+/jOKga74ITtdMD+7c0NTaCS8sZ5HcG
# rBjVztdIFuDSv59sA52i5wZeO3GtGkDPiJA1Q3Us2uQMxbeReHHbFIv8YofTMmik
# s47dTTOLylTZ7bH7dGAkrFOdizanJ54RUT33PeKMJ9/FYjIv097tsTBT7g0Sccer
# j+ICkJ8oxvJswYQCfrVhjUZ10S+BiaSUfs2ALPVyh4LffutfAzxfcYDOjo9U6wud
# Xs/7WuU9qq4UMrS2mbzeLO8XXTypGHHy+K7wRPEVu/BWTOA0qE/mjUMKTv38x3NC
# 94Ggv7wKH0nMLgdDWhZbMH7Fe9JY7G9gfPUAlAmt5ywa5UoW1+zIUB0mEAN81CEk
# bDCqgCikdg==
# SIG # End signature block
