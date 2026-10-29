# Read-only operational insights for the SmartM365 Orchestrator GUI: live server state, job
# health, dependency readiness, pipeline requests and job run requests. No WPF code here, so
# every function can be tested headless against a copy of the shared Orchestrator folder.
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Pipeline.psm1') -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Maintenance.psm1') -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Distributed.psm1') -ErrorAction Stop

$script:InsightsTerminalSuccess = @('Success', 'CompletedWithWarnings')
$script:InsightsTerminalFailure = @('Failed', 'TimedOut', 'Interrupted', 'Cancelled')
# Same lists as SmartM365.Orchestrator.Pipeline.psm1 (Get-SmartM365OrchestratorPipelineRunStatus).
$script:InsightsPipelineFailure = @('Failed', 'TimedOut', 'Interrupted', 'BlockedDependencyFailed', 'BlockedDependencyTimeout', 'BlockedPrerequisite', 'Rejected', 'MissingStatus')
$script:InsightsPipelineTerminal = @('Success', 'CompletedWithWarnings', 'Cancelled') + $script:InsightsPipelineFailure

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
        CreatedUtc = ConvertTo-InsightsUtc (Get-InsightsProperty $document 'CreatedAtUtc' $null)
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
function Get-InsightsLaterOccurrenceClaim {
    param([string]$ClaimsRoot, [string]$JobName, [datetime]$ExpectedUtc, [datetime]$NowUtc)
    foreach ($file in @(Get-InsightsClaimFiles -ClaimsRoot $ClaimsRoot -JobName $JobName)) {
        $occurrence = [datetime]::ParseExact($file.Name.Substring(0, 19), 'yyyyMMddTHHmmssfffZ', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal)
        if ($occurrence -le $ExpectedUtc) { break }
        if ($occurrence -gt $NowUtc) { continue }
        $claim = Get-InsightsOccurrenceClaim -ClaimsRoot $ClaimsRoot -JobName $JobName -OccurrenceUtc $occurrence
        if ($null -eq $claim) { return $null }
        if ($null -eq $claim.CreatedUtc -or $null -eq $claim.UpdatedUtc -or
            $claim.CreatedUtc -lt $ExpectedUtc -or $claim.UpdatedUtc -lt $claim.CreatedUtc -or $claim.UpdatedUtc -gt $NowUtc) { return $null }
        return $claim
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
    $defaultMode = [string](Get-InsightsProperty $job 'DependencyMode' 'LatestOccurrence')
    $override = [int](Get-InsightsProperty $job 'DependencyMaxAgeHours' 0)
    foreach ($dependencyName in @(Get-InsightsProperty $job 'DependsOn' @() | ForEach-Object { [string]$_ })) {
        $mode = if ($dependencyName -in @(Get-InsightsProperty $job 'LatestOccurrenceDependencies' @())) { 'LatestOccurrence' } else { $defaultMode }
        $dependency = @($JobsDocument.Jobs | Where-Object { [string]$_.Name -eq $dependencyName })[0]
        $row = [ordered]@{ Dependency = $dependencyName; Rule = $mode; State = ''; LastSuccess = ''; AgeHours = $null; MaxAgeHours = $null; Detail = '' }
        if ($null -eq $dependency) { $row.State = 'Failed'; $row.Detail = 'Unknown job'; [pscustomobject]$row; continue }
        if (-not [bool](Get-InsightsProperty $dependency 'Enabled' $false) -or [string](Get-InsightsProperty $dependency 'AssignmentMode' 'Legacy') -eq 'Manual') {
            $row.State = 'Ignored'; $row.Detail = 'Disabled or manual: ignored by the dependency gate'; [pscustomobject]$row; continue
        }
        $maxAgeHours = if ($mode -eq 'FreshSuccess') { Get-SmartM365OrchestratorExpectedMaxAgeHours -Job $dependency -OverrideHours $override -Now $Now } else { 0 }
        if ($mode -eq 'FreshSuccess' -and [string](Get-InsightsProperty $dependency 'AssignmentMode' '') -eq 'Pinned') {
            $cluster = Read-InsightsJson -Path (Join-Path $SharedDataFolderPath 'Config/Orchestrator-Cluster.json')
            $profile = [string](Get-Variable SmartM365ProfileKey -Scope Global -ValueOnly -ErrorAction SilentlyContinue)
            $pinned = Get-SmartM365OrchestratorPinnedDependencyReadiness -SharedDataFolderPath $SharedDataFolderPath -Job $dependency `
                -Now $Now -MaxAgeHours $maxAgeHours -ProfileKey $profile `
                -HeartbeatStaleMinutes ([double](Get-InsightsProperty $cluster 'PeerHeartbeatStaleMinutes' 5))
            $row.State = $pinned.State; $row.Detail = $pinned.Detail; $row.MaxAgeHours = $maxAgeHours
            if ($null -ne $pinned.LastSuccessUtc) {
                $row.LastSuccess = $pinned.LastSuccessUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
                $row.AgeHours = [math]::Round(($Now.ToUniversalTime() - $pinned.LastSuccessUtc).TotalHours, 1)
            }
            [pscustomobject]$row; continue
        }
        if ($mode -eq 'FreshSuccess' -and [string](Get-InsightsProperty $dependency 'AssignmentMode' '') -eq 'Elected') {
            # Same activity rule as the scheduler, evaluated for a start now.
            $scriptName = [IO.Path]::GetFileName([string](Get-InsightsProperty $dependency 'ScriptPath' ''))
            $sharedScript = @($JobsDocument.Jobs | Where-Object { [bool](Get-InsightsProperty $_ 'Enabled' $false) -and [IO.Path]::GetFileName([string](Get-InsightsProperty $_ 'ScriptPath' '')) -ieq $scriptName }).Count -gt 1
            $runState = if ($sharedScript) { $null } else { Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath $SharedDataFolderPath -Job $dependency }
            $latestPast = @(Get-SmartM365OrchestratorScheduleOccurrences -Schedule $dependency.Schedule -From $Now.AddDays(-8) -To $Now | Select-Object -Last 1)
            $activity = $null
            try {
                $activity = Get-SmartM365OrchestratorDependencyActivity -ClaimsRootPath $claimsRoot -DependencyJob $dependency -DownstreamOccurrence $Now -Now $Now `
                    -LatestPastOccurrence $(if ($latestPast.Count) { $latestPast[0] } else { $null }) -RunState $runState
            }
            catch { $activity = [pscustomobject]@{ State = 'Waiting'; Detail = 'Dependency activity could not be read: ' + $_.Exception.Message } }
            if ($activity.State -ne 'Clear') {
                $row.State = 'Waiting'; $row.Detail = $activity.Detail
                $row.MaxAgeHours = [math]::Round($maxAgeHours, 1)
                [pscustomobject]$row; continue
            }
        }
        $successClaim = Get-InsightsLatestSuccessClaim -ClaimsRoot $claimsRoot -JobName $dependencyName
        $successUtc = if ($null -ne $successClaim) { $successClaim.UpdatedUtc } else { $null }
        $runSuccess = @($Runs | Where-Object { $_.JobName -eq $dependencyName -and $_.Status -in $script:InsightsTerminalSuccess } | Select-Object -First 1)
        if ($runSuccess.Count) {
            $runUtc = ([datetime]$(if ($null -ne $runSuccess[0].EndTime) { $runSuccess[0].EndTime } else { $runSuccess[0].StartTime })).ToUniversalTime()
            if ($null -eq $successUtc -or $runUtc -gt $successUtc) { $successUtc = $runUtc }
        }
        if ($mode -eq 'FreshSuccess') {
            $receiptUtc = Get-SmartM365OrchestratorFreshProducerReceipt -SharedDataFolderPath $SharedDataFolderPath -Job $dependency -TenantKey ([string]$global:SmartM365TenantKey) -Now $Now -MaxAgeHours $maxAgeHours
            if ($null -ne $receiptUtc -and ($null -eq $successUtc -or $receiptUtc -gt $successUtc)) {
                $successUtc = $receiptUtc
                $row.Detail = 'Qualified producer receipt'
            }
        }
        if ($null -ne $successUtc) {
            $row.LastSuccess = $successUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
            $row.AgeHours = [math]::Round(($Now.ToUniversalTime() - $successUtc).TotalHours, 1)
        }
        if ($mode -eq 'FreshSuccess') {
            $row.MaxAgeHours = [math]::Round($maxAgeHours, 1)
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
                if ($null -eq $claim -or $status -in $script:InsightsTerminalFailure) {
                    $later = Get-InsightsLaterOccurrenceClaim -ClaimsRoot $claimsRoot -JobName $dependencyName -ExpectedUtc $expectedUtc -NowUtc $Now.ToUniversalTime()
                    if ($null -ne $later) {
                        $status = $later.Status
                        $row.Detail = "Latest occurrence $($expected[0].ToString('yyyy-MM-dd HH:mm')); later run $($later.Occurrence.ToLocalTime().ToString('yyyy-MM-dd HH:mm')): $status"
                    }
                }
                $row.State = if ($status -in $script:InsightsTerminalSuccess) { 'Ready' } elseif ($status -in $script:InsightsTerminalFailure) { 'Failed' } else { 'Waiting' }
            }
        }
        [pscustomobject]$row
    }
}

function Get-SmartM365OrchestratorServerMetrics {
    <# CPU and memory samples of one server (Orchestrator-Metrics.json) within the last Hours. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ServerFolder, [string]$Server = '', [int]$Hours = 12, [datetime]$Now = (Get-Date))
    $document = Read-InsightsJson -Path (Join-Path $ServerFolder 'Orchestrator-Metrics.json')
    $from = $Now.ToUniversalTime().AddHours(-1 * $Hours)
    $samples = @(foreach ($sample in @(Get-InsightsProperty $document 'Samples' @())) {
        $time = ConvertTo-InsightsUtc (Get-InsightsProperty $sample 'TimeUtc' $null)
        if ($null -eq $time -or $time -lt $from -or $time -gt $Now.ToUniversalTime().AddMinutes(5)) { continue }
        [pscustomobject]@{ Time = $time.ToLocalTime(); Cpu = Get-InsightsProperty $sample 'Cpu' $null; Memory = Get-InsightsProperty $sample 'Memory' $null }
    })
    [pscustomobject]@{ Server = $Server; From = $from.ToLocalTime(); To = $Now; SampleMinutes = [int](Get-InsightsProperty $document 'SampleMinutes' 5); Samples = @($samples | Sort-Object Time) }
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
    $metrics = [Collections.Generic.List[object]]::new()
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
        $usage = Get-InsightsProperty $heartbeat 'ResourceUsage' $null
        $metrics.Add((Get-SmartM365OrchestratorServerMetrics -ServerFolder $folder -Server $serverName -Now $Now))
        $state = if ($null -eq $heartbeat) { 'No heartbeat' } elseif ($null -eq $age) { 'No timestamp' } elseif ($age -gt $staleMinutes) { 'Stale' } elseif ($lifecycle -eq 'Recycling') { 'Recycling' } elseif ($lifecycle -eq 'Starting') { 'Starting' } else { 'Online' }
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
            CpuPercent = Get-InsightsProperty $usage 'Cpu' $null; MemoryPercent = Get-InsightsProperty $usage 'Memory' $null
            Maintenance = if ($maintenanceError) { 'Unknown' } elseif ($maintenance.Enabled) { 'ON' } else { 'OFF' }
            MaintenanceAcknowledgement = if ($maintenanceError) { 'Control unavailable' } else { [string](@($maintenanceServers | Where-Object Server -eq $serverName)[0].Status) }
            StatePersistence = if ([bool](Get-InsightsProperty $heartbeat 'StatePersistenceHealthy' $true)) { 'OK' } else { 'Launches paused: ' + [string](Get-InsightsProperty $heartbeat 'StatePersistenceLastError' '') }
            ReadyCapabilities = (@(Get-InsightsProperty $heartbeat 'ReadyCapabilities' @() | ForEach-Object { [string]$_ } | Sort-Object -Unique) -join ', ')
            CertificateDaysRemaining = Get-InsightsProperty $heartbeat 'CertificateDaysRemaining' $null
            PrerequisiteBlocks = (@(Get-InsightsProperty $heartbeat 'PrerequisiteBlocks' @() | ForEach-Object { if ($_ -is [string]) { $_ } else { '{0}: {1}' -f (Get-InsightsProperty $_ 'Name' ''), (Get-InsightsProperty $_ 'Detail' '') } }) -join '; ')
        })
        foreach ($job in $jobsRunning) {
            $rawStart = Get-InsightsProperty $job 'StartTime' $null
            $start = ConvertTo-InsightsUtc $rawStart
            $startText = if ($rawStart -is [datetime] -or $rawStart -is [datetimeoffset]) { $rawStart.ToString('o') } else { [string]$rawStart }
            $running.Add([pscustomobject]@{
                Server = $serverName; ServerState = $state; HeartbeatAgeMinutes = $age; Job = [string](Get-InsightsProperty $job 'Name' ''); Pid = [string](Get-InsightsProperty $job 'Pid' '')
                StartTime = $startText
                StopSupported = $state -eq 'Online' -and [int](Get-InsightsProperty $heartbeat 'JobStopProtocol' 0) -ge 1
                Started = if ($null -ne $start) { $start.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
                DurationMinutes = if ($null -ne $start) { [math]::Round(($Now.ToUniversalTime() - $start).TotalMinutes, 0) } else { $null }
                Occurrence = [string](Get-InsightsProperty $job 'ScheduledOccurrence' '')
            })
        }
        foreach ($job in $jobsPending) {
            $firstSeen = ConvertTo-InsightsUtc (Get-InsightsProperty $job 'FirstSeen' $null)
            $pending.Add([pscustomobject]@{
                Server = $serverName; ServerState = $state; HeartbeatAgeMinutes = $age; Job = [string](Get-InsightsProperty $job 'Name' ''); Reason = [string](Get-InsightsProperty $job 'Reason' '')
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
        Metrics = @($metrics)
    }
}

function Get-SmartM365OrchestratorRecentPipelineRuns {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [int]$Count = 20)
    $root = Join-Path $SharedDataFolderPath 'PipelineRuns'
    if (-not (Test-Path -LiteralPath $root)) { return @() }
    # Direct read-only reads: the pipeline module resolves and validates request.json once per
    # job status, which costs about one second per batch on a shared folder.
    $rows = @(foreach ($folder in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First $Count)) {
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
        $cancellationRequested = $null -ne (Get-InsightsProperty $request 'Cancellation' $null)
        $cancelledCount = @($jobRows | Where-Object Status -eq 'Cancelled').Count
        $failedCount = @($jobRows | Where-Object { $_.Status -in $script:InsightsPipelineFailure }).Count
        if (-not $cancellationRequested) { $failedCount += $cancelledCount }
        $warningCount = @($jobRows | Where-Object { $_.Status -eq 'CompletedWithWarnings' }).Count
        $created = ConvertTo-InsightsUtc (Get-InsightsProperty $request 'CreatedAtUtc' $null)
        $activated = ConvertTo-InsightsUtc (Get-InsightsProperty $request 'ActivatedAtUtc' $null)
        [pscustomobject]@{
            BatchId = $folder.Name; Selection = [string](Get-InsightsProperty $request 'Pipeline' '')
            Created = if ($null -ne $created) { $created.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
            Activated = if ($null -ne $activated) { $activated.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
            Status = if ($pendingCount -gt 0) { if ($cancellationRequested) { 'Cancelling' } else { 'Running' } } elseif ($failedCount -gt 0) { 'Failed' } elseif ($cancellationRequested) { 'Cancelled' } elseif ($warningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
            Cancelled = $cancelledCount
            Jobs = $jobRows.Count; Pending = $pendingCount; Failed = $failedCount
            RequestedBy = [string](Get-InsightsProperty $request 'RequestedBy' ''); JobRows = $jobRows
            QueuePosition = 0; CreatedSortKey = if ($null -ne $created) { $created.ToString('o') } else { '' }
        }
    })
    # Same FIFO order as the residents: the oldest active request runs, the others are queued.
    $active = @($rows | Where-Object { $_.Status -in @('Running', 'Cancelling') } | Sort-Object CreatedSortKey, BatchId)
    for ($index = 1; $index -lt $active.Count; $index++) {
        $active[$index].QueuePosition = $index
        if ($active[$index].Status -eq 'Running') { $active[$index].Status = 'Queued' }
    }
    $rows
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

function Get-SmartM365OrchestratorDashboard {
    <# Pure dashboard aggregation (banner, tiles, items to review, servers, next 24 h timeline). #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$JobsDocument, [Parameter(Mandatory)]$Operations,
        [AllowEmptyCollection()][object[]]$Health = @(), [AllowEmptyCollection()][object[]]$Runs = @(),
        [AllowNull()]$Plan = $null, [AllowEmptyCollection()][object[]]$Requests = @(), [AllowNull()]$LogEvents = $null,
        [AllowEmptyCollection()][object[]]$SwiRows = @(),
        [datetime]$Now = (Get-Date), [int]$CertificateWarningDays = 30, [int]$RequestStaleHours = 6, [int]$TimelineHours = 24
    )
    $actions = [Collections.Generic.List[object]]::new()
    function Add-DashboardAction([string]$Severity, [string]$Area, [string]$Item, [string]$Detail, [string]$Target) {
        $actions.Add([pscustomobject]@{ Severity = $Severity; Area = $Area; Item = $Item; Detail = $Detail; Target = $Target })
    }
    $servers = @($Operations.Servers)
    $jobs = @($JobsDocument.Jobs)
    $enabledJobs = @($jobs | Where-Object { [bool](Get-InsightsProperty $_ 'Enabled' $false) })
    $owners = @{}; $ownedCount = @{}
    foreach ($assignment in @(Get-InsightsProperty $Plan 'Assignments' @())) {
        $owner = [string](Get-InsightsProperty $assignment 'OwnerServer' '')
        $owners[[string](Get-InsightsProperty $assignment 'JobName' '')] = $owner
        if ($owner) { $ownedCount[$owner] = 1 + [int]$(if ($ownedCount.ContainsKey($owner)) { $ownedCount[$owner] } else { 0 }) }
    }

    # Items to review, most severe first.
    foreach ($server in $servers) {
        switch ([string]$server.State) {
            'Online' { }
            { $_ -in @('Recycling', 'Starting') } { Add-DashboardAction 'Info' 'Server' $server.Server ('{0} (heartbeat {1} min).' -f $server.State, $server.HeartbeatAgeMinutes) 'Operations' }
            default { Add-DashboardAction 'Critical' 'Server' $server.Server ('{0}; heartbeat age {1} min.' -f $server.State, $(if ($null -ne $server.HeartbeatAgeMinutes) { $server.HeartbeatAgeMinutes } else { 'unknown' })) 'Operations' }
        }
        if ([string]$server.StatePersistence -ne 'OK') { Add-DashboardAction 'Critical' 'State persistence' $server.Server ([string]$server.StatePersistence) 'Events' }
        if ([string](Get-InsightsProperty $server 'PrerequisiteBlocks' '')) { Add-DashboardAction 'Critical' 'Launch prerequisite' $server.Server ([string]$server.PrerequisiteBlocks) 'Operations' }
        $days = Get-InsightsProperty $server 'CertificateDaysRemaining' $null
        if ($null -ne $days -and [double]$days -lt $CertificateWarningDays) {
            Add-DashboardAction $(if ([double]$days -lt 7) { 'Critical' } else { 'Warning' }) 'Certificate' $server.Server ('Application certificate expires in {0} day(s).' -f [math]::Floor([double]$days)) 'Operations'
        }
    }
    if ([string](Get-InsightsProperty $Operations 'MaintenanceError' '')) { Add-DashboardAction 'Critical' 'Maintenance' 'Shared control' ('Maintenance control unavailable: ' + $Operations.MaintenanceError) 'Operations' }
    elseif ($null -ne $Operations.Maintenance -and [bool](Get-InsightsProperty $Operations.Maintenance 'Enabled' $false)) { Add-DashboardAction 'Warning' 'Maintenance' 'Scheduling' ('Maintenance is ON: scheduled launches are suspended. Reason: ' + [string](Get-InsightsProperty $Operations.Maintenance 'Reason' '')) 'Operations' }
    foreach ($row in @($Health | Where-Object { $_.Health -in @('Failing', 'Stale', 'NoRecentSuccess') })) {
        $detail = switch ($row.Health) {
            'Failing' { 'Last run {0} at {1}; last success {2}.' -f $row.LastStatus, $row.LastRun, $(if ($row.LastSuccess) { $row.LastSuccess } else { 'none in 21 days' }) }
            'Stale' { 'Last success {0} ({1} h, expected at most {2} h).' -f $row.LastSuccess, $row.SuccessAgeHours, $row.ExpectedMaxAgeHours }
            default { 'No successful run in the last 21 days.' }
        }
        Add-DashboardAction $(if ($row.Health -eq 'Failing') { 'Critical' } else { 'Warning' }) ('Job ' + $row.Health) $row.Name $detail 'Planning'
    }
    foreach ($group in @(Get-InsightsProperty $Plan 'UnassignedGroups' @())) {
        $groupName = [string](Get-InsightsProperty $group 'GroupKey' (Get-InsightsProperty $group 'Group' $group))
        $reason = [string](Get-InsightsProperty $group 'Reason' 'No live server satisfies the job requirements.')
        Add-DashboardAction 'Warning' 'Election' $groupName ('Unassigned: ' + $reason) 'Servers and versions'
    }
    foreach ($incident in @($Operations.Incidents)) { Add-DashboardAction 'Warning' 'Peer monitoring' ([string]$incident.ObservedBy) ([string]$incident.Issue) 'Events' }
    # A required SWI input that WorkplaceEvidence-Prepare would refuse (Get-SmartM365OrchestratorSwiValidity).
    foreach ($swi in @($SwiRows | Where-Object { [bool](Get-InsightsProperty $_ 'Blocks' $false) })) {
        Add-DashboardAction 'Critical' 'SWI validity' ([string]$swi.Name) ('{0}. {1}' -f [string]$swi.Swi, [string](Get-InsightsProperty $swi 'Detail' '')) 'Planning'
    }
    # Queued requests are not stale: their wait starts when the residents activate them.
    foreach ($request in @($Requests | Where-Object { $_.Status -in @('Running', 'Cancelling') })) {
        $created = [datetime]::MinValue
        $since = if ($request.PSObject.Properties['Activated'] -and [string]$request.Activated) { [string]$request.Activated } else { [string]$request.Created }
        if ([datetime]::TryParseExact($since, 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$created) -and ($Now - $created).TotalHours -gt $RequestStaleHours) {
            Add-DashboardAction 'Warning' 'Pipeline request' ([string]$request.BatchId) ('{0} for {1:0} h; {2} job(s) not finished.' -f $request.Status, ($Now - $created).TotalHours, $request.Pending) 'Requests'
        }
    }
    $recentEvents = @()
    if ($null -ne $LogEvents) {
        $recentEvents = @($LogEvents.Events | Where-Object { $_.Time -ge $Now.AddHours(-24) })
        # A resident restarted after the last problem line no longer carries that state.
        $lastStart = @{}
        foreach ($start in @($LogEvents.Events | Where-Object { $_.Category -eq 'Lifecycle' -and $_.Message -match '^\S+ v\S+ starting \(PID ' })) {
            if (-not $lastStart.ContainsKey($start.Server) -or $start.Time -gt $lastStart[$start.Server]) { $lastStart[$start.Server] = $start.Time }
        }
        foreach ($group in @($LogEvents.Events | Where-Object { $_.Time -ge $Now.AddHours(-1) -and $_.Category -in @('Heartbeat', 'StatePersistence', 'ConfigReload') -and $_.Level -ne 'INFO' } | Group-Object Server, Category)) {
            $last = @($group.Group | Sort-Object Time -Descending)[0]
            if ($lastStart.ContainsKey($last.Server) -and $lastStart[$last.Server] -gt $last.Time) { continue }
            Add-DashboardAction 'Warning' ('Log ' + $last.Category) $last.Server ('{0} line(s) in the last hour; latest {1:HH:mm}: {2}' -f $group.Count, $last.Time, $last.Message) 'Events'
        }
    }
    $rank = @{ Critical = 0; Warning = 1; Info = 2 }
    $sortedActions = @($actions | Sort-Object @{ Expression = { $rank[$_.Severity] } }, Area, Item)

    # Tiles.
    $online = @($servers | Where-Object State -eq 'Online').Count
    $unhealthy = @($Health | Where-Object { $_.Health -in @('Failing', 'Stale', 'NoRecentSuccess') })
    $failingCount = @($unhealthy | Where-Object Health -eq 'Failing').Count
    $failures = @($Runs | Where-Object { $_.Status -in $script:InsightsTerminalFailure -and $(if ($null -ne $_.EndTime) { [datetime]$_.EndTime } else { [datetime]$_.StartTime }) -ge $Now.AddHours(-24) })
    $pendingRows = @($Operations.Pending)
    $topReason = @($pendingRows | Group-Object Reason | Sort-Object Count -Descending | Select-Object -First 1)
    $certificateDays = @($servers | ForEach-Object { Get-InsightsProperty $_ 'CertificateDaysRemaining' $null } | Where-Object { $null -ne $_ } | ForEach-Object { [double]$_ } | Sort-Object)
    $errorEvents = @($recentEvents | Where-Object Level -eq 'ERROR').Count
    $warnEvents = @($recentEvents | Where-Object Level -eq 'WARN').Count
    $tiles = @(
        [pscustomobject]@{ Key = 'Jobs'; Label = 'Jobs'; Value = [string]$jobs.Count; Detail = '{0} manual, {1} disabled' -f @($jobs | Where-Object { [string](Get-InsightsProperty $_ 'AssignmentMode' '') -eq 'Manual' }).Count, ($jobs.Count - $enabledJobs.Count); Level = 'Normal'; Target = 'Planning' }
        [pscustomobject]@{ Key = 'Enabled'; Label = 'Enabled'; Value = [string]$enabledJobs.Count; Detail = 'scheduled or manual'; Level = 'Normal'; Target = 'Planning' }
        [pscustomobject]@{ Key = 'Servers'; Label = 'Servers online'; Value = '{0}/{1}' -f $online, $servers.Count; Detail = $(if ($online -eq $servers.Count) { 'all heartbeats fresh' } else { (@($servers | Where-Object State -ne 'Online' | ForEach-Object { '{0} {1}' -f $_.Server, $_.State }) -join ', ') }); Level = $(if ($servers.Count -and $online -eq $servers.Count) { 'Normal' } elseif (@($servers | Where-Object { $_.State -notin @('Online', 'Recycling', 'Starting') }).Count) { 'Critical' } else { 'Warning' }); Target = 'Operations' }
        [pscustomobject]@{ Key = 'Running'; Label = 'Running'; Value = [string]@($Operations.Running).Count; Detail = 'jobs in progress'; Level = 'Normal'; Target = 'Operations' }
        [pscustomobject]@{ Key = 'Pending'; Label = 'Pending'; Value = [string]$pendingRows.Count; Detail = $(if ($topReason.Count) { 'mostly ' + $(if ($topReason[0].Name) { $topReason[0].Name } else { 'waiting' }) } else { 'nothing waiting' }); Level = $(if ($pendingRows.Count) { 'Info' } else { 'Normal' }); Target = 'Operations' }
        [pscustomobject]@{ Key = 'Health'; Label = 'Job health'; Value = [string]$unhealthy.Count; Detail = '{0} failing, {1} stale, {2} no success' -f $failingCount, @($unhealthy | Where-Object Health -eq 'Stale').Count, @($unhealthy | Where-Object Health -eq 'NoRecentSuccess').Count; Level = $(if ($failingCount) { 'Critical' } elseif ($unhealthy.Count) { 'Warning' } else { 'Normal' }); Target = 'Planning' }
        [pscustomobject]@{ Key = 'Failures'; Label = 'Failures 24 h'; Value = [string]$failures.Count; Detail = $(if ($failures.Count) { (@($failures | ForEach-Object JobName | Sort-Object -Unique) | Select-Object -First 3) -join ', ' } else { 'no failed run' }); Level = $(if ($failures.Count) { 'Critical' } else { 'Normal' }); Target = 'Failures24h' }
        [pscustomobject]@{ Key = 'Events'; Label = 'Log errors 24 h'; Value = $(if ($null -eq $LogEvents) { '-' } else { [string]$errorEvents }); Detail = $(if ($null -eq $LogEvents) { 'resident logs not read' } else { '{0} warning(s)' -f $warnEvents }); Level = $(if ($errorEvents) { 'Warning' } else { 'Normal' }); Target = 'Events' }
        [pscustomobject]@{ Key = 'Certificate'; Label = 'Certificate'; Value = $(if ($certificateDays.Count) { '{0} d' -f [math]::Floor($certificateDays[0]) } else { '-' }); Detail = $(if ($certificateDays.Count) { 'days before expiry (lowest server)' } else { 'not reported by heartbeats' }); Level = $(if (-not $certificateDays.Count) { 'Normal' } elseif ($certificateDays[0] -lt 7) { 'Critical' } elseif ($certificateDays[0] -lt $CertificateWarningDays) { 'Warning' } else { 'Normal' }); Target = 'Operations' }
    )

    # Servers.
    $serverRows = @(foreach ($server in $servers) {
        [pscustomobject]@{
            Server = $server.Server; State = $server.State; Version = $server.Version; StatePersistence = $server.StatePersistence
            Running = $server.Running; Pending = $server.Pending; RecycleIn = $server.RecycleIn
            CpuPercent = Get-InsightsProperty $server 'CpuPercent' $null; MemoryPercent = Get-InsightsProperty $server 'MemoryPercent' $null
            Owned = $(if ($ownedCount.ContainsKey([string]$server.Server)) { $ownedCount[[string]$server.Server] } else { 0 })
            Certificate = $(if ($null -ne (Get-InsightsProperty $server 'CertificateDaysRemaining' $null)) { '{0} d' -f [math]::Floor([double]$server.CertificateDaysRemaining) } else { '' })
            Capabilities = [string](Get-InsightsProperty $server 'ReadyCapabilities' '')
            HeartbeatAgeMinutes = $server.HeartbeatAgeMinutes
        }
    })

    # Next hours: scheduled occurrences on their owner lane, plus running jobs.
    $healthByName = @{}; foreach ($row in @($Health)) { $healthByName[[string]$row.Name] = $row }
    $maintenanceOn = $null -ne $Operations.Maintenance -and [bool](Get-InsightsProperty $Operations.Maintenance 'Enabled' $false)
    $to = $Now.AddHours($TimelineHours)
    $items = [Collections.Generic.List[object]]::new()
    foreach ($job in @($enabledJobs | Where-Object { [string](Get-InsightsProperty $_ 'AssignmentMode' '') -ne 'Manual' })) {
        $name = [string]$job.Name
        $mode = [string](Get-InsightsProperty $job 'AssignmentMode' 'Legacy')
        $allowed = @(Get-InsightsProperty $job 'AllowedServers' @() | ForEach-Object { [string]$_ } | Where-Object { $_ })
        $lane = if ($mode -eq 'Elected') { if ($owners.ContainsKey($name) -and $owners[$name]) { $owners[$name] } else { 'Unassigned' } } elseif ($allowed.Count -eq 1) { $allowed[0] } else { 'Any server' }
        $minutes = [double](Get-InsightsProperty $healthByName[$name] 'AverageDurationMinutes' 0)
        if ($minutes -le 0) { $minutes = [double](Get-InsightsProperty $job 'EstimatedDurationMinutes' 5) }
        if ($minutes -le 0) { $minutes = 5 }
        foreach ($occurrence in @(Get-SmartM365OrchestratorScheduleOccurrences -Schedule $job.Schedule -From $Now -To $to)) {
            $kind = if ($lane -eq 'Unassigned') { 'Unassigned' } elseif ($maintenanceOn) { 'Suspended' } else { 'Planned' }
            $items.Add([pscustomobject]@{ Lane = $lane; Job = $name; Start = $occurrence; End = $occurrence.AddMinutes($minutes); Kind = $kind })
        }
    }
    foreach ($running in @($Operations.Running)) {
        $start = $null; $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact([string]$running.Started, 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { $start = $parsed } else { $start = $Now }
        $minutes = [double](Get-InsightsProperty $healthByName[[string]$running.Job] 'AverageDurationMinutes' 0)
        $end = if ($minutes -gt 0) { $start.AddMinutes($minutes) } else { $Now.AddMinutes(15) }
        if ($end -lt $Now.AddMinutes(5)) { $end = $Now.AddMinutes(5) }
        $items.Add([pscustomobject]@{ Lane = [string]$running.Server; Job = [string]$running.Job; Start = $start; End = $end; Kind = 'Running' })
    }
    $lanes = @(@($servers | ForEach-Object { [string]$_.Server }) + @($items | ForEach-Object Lane | Where-Object { $_ -notin @($servers | ForEach-Object { [string]$_.Server }) } | Sort-Object -Unique))

    $critical = @($sortedActions | Where-Object Severity -eq 'Critical').Count
    $warning = @($sortedActions | Where-Object Severity -eq 'Warning').Count
    $summary = '{0}/{1} server(s) online | {2} running | {3} pending | {4} failure(s) in 24 h | maintenance {5}' -f $online, $servers.Count, @($Operations.Running).Count, $pendingRows.Count, $failures.Count, $(if ($maintenanceOn) { 'ON' } else { 'off' })
    $banner = if ($critical) { [pscustomobject]@{ Level = 'Critical'; Text = '{0} critical and {1} warning item(s) to review. {2}' -f $critical, $warning, $summary } }
    elseif ($warning) { [pscustomobject]@{ Level = 'Warning'; Text = '{0} warning item(s) to review. {1}' -f $warning, $summary } }
    else { [pscustomobject]@{ Level = 'Normal'; Text = 'Everything looks healthy. ' + $summary } }
    [pscustomobject]@{
        Banner = $banner; Tiles = $tiles; Actions = $sortedActions; Servers = $serverRows
        Metrics = @(Get-InsightsProperty $Operations 'Metrics' @())
        Timeline = [pscustomobject]@{ From = $Now; To = $to; Lanes = $lanes; Items = @($items | Sort-Object Start) }
        GeneratedAt = $Now
    }
}

function Get-SmartM365OrchestratorDashboardData {
    <# Reads the remaining read-only sources (runs, election plan) and builds the dashboard. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)]$JobsDocument, [Parameter(Mandatory)]$Operations,
        [AllowNull()][object[]]$Runs = $null, [AllowEmptyCollection()][object[]]$Requests = @(), [AllowNull()]$LogEvents = $null,
        [AllowEmptyCollection()][object[]]$SwiRows = @(), [datetime]$Now = (Get-Date)
    )
    if ($null -eq $Runs) { $Runs = @(Get-SmartM365OrchestratorRecentRuns -SharedDataFolderPath $SharedDataFolderPath -Days 21 -Now $Now) }
    $health = @(Get-SmartM365OrchestratorJobHealth -JobsDocument $JobsDocument -Runs $Runs -Now $Now)
    $plan = Read-InsightsJson -Path (Join-Path $SharedDataFolderPath 'Election\Orchestrator-ElectionPlan.json')
    Get-SmartM365OrchestratorDashboard -JobsDocument $JobsDocument -Operations $Operations -Health $health -Runs $Runs -Plan $plan -Requests $Requests -LogEvents $LogEvents -SwiRows $SwiRows -Now $Now
}

# Resident log lines kept as events. ERROR/WARN lines are always kept; INFO lines only when
# they match a rule. Messages are those written by SmartM365-Inventory-Orchestrator.ps1.
$script:InsightsLogEventRules = @(
    @{ Category = 'Lifecycle'; Pattern = '^\S+ v\S+ starting \(PID |^\S+ exiting normally|^Runtime update candidate detected|^Validated runtime update|^Heartbeat published as Recycling|^Maximum lifetime' }
    @{ Category = 'StopRequest'; Pattern = '(?i)stop request' }
    @{ Category = 'Heartbeat'; Pattern = '^Failed to write heartbeat' }
    @{ Category = 'StatePersistence'; Pattern = '^State persistence ' }
    @{ Category = 'ConfigReload'; Pattern = 'reload rejected' }
    @{ Category = 'Maintenance'; Pattern = '^Maintenance control ' }
)
$script:InsightsLogEventNoise = '^(Email sent |Mail HTML copy |SharePoint |Acquiring )'
$script:InsightsLogLinePattern = [regex]'^\[?(?<Time>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]?\s?\[(?<Level>[A-Z]+)\]\s?(?<Message>.*)$'
$script:InsightsLogMaxReadBytes = 33554432

function Get-SmartM365OrchestratorLogEventCategory {
    <# Category of a resident log line, or $null when the line is not an orchestrator event. #>
    param([Parameter(Mandatory)][string]$Level, [AllowEmptyString()][string]$Message)
    if ($Message -notmatch $script:InsightsLogEventNoise) {
        foreach ($rule in $script:InsightsLogEventRules) { if ($Message -match $rule.Pattern) { return $rule.Category } }
    }
    if ($Level -eq 'ERROR') { return 'Error' }
    if ($Level -eq 'WARN') { return 'Warning' }
    return $null
}

function Read-InsightsLogTail {
    # Shared read of complete lines appended after Offset; the resident keeps writing the file.
    param([Parameter(Mandatory)][string]$Path, [long]$Offset = 0)
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try {
        $length = $stream.Length; $reset = $false; $skipPartial = $false
        if ($Offset -gt $length) { $Offset = 0; $reset = $true }
        if (($length - $Offset) -gt $script:InsightsLogMaxReadBytes) { $Offset = $length - $script:InsightsLogMaxReadBytes; $skipPartial = $true }
        $count = [int]($length - $Offset)
        if ($count -le 0) { return [pscustomobject]@{ Lines = @(); Offset = $Offset; Reset = $reset } }
        $buffer = [byte[]]::new($count); $stream.Position = $Offset; $read = 0
        while ($read -lt $count) { $chunk = $stream.Read($buffer, $read, $count - $read); if ($chunk -le 0) { break }; $read += $chunk }
        $last = [Array]::LastIndexOf($buffer, [byte]10, $read - 1)
        if ($read -le 0 -or $last -lt 0) { return [pscustomobject]@{ Lines = @(); Offset = $Offset; Reset = $reset } }
        $start = if ($Offset -eq 0 -and $read -ge 3 -and $buffer[0] -eq 0xEF -and $buffer[1] -eq 0xBB -and $buffer[2] -eq 0xBF) { 3 } else { 0 }
        $lines = @([Text.Encoding]::UTF8.GetString($buffer, $start, $last + 1 - $start) -split '\r?\n')
        if ($skipPartial -and $lines.Count) { $lines = @($lines | Select-Object -Skip 1) }
        return [pscustomobject]@{ Lines = $lines; Offset = $Offset + $last + 1; Reset = $reset }
    }
    finally { $stream.Dispose() }
}

function Get-SmartM365OrchestratorLogEvents {
    <# Read-only events from the resident daily logs of each server. Cache (path -> offset and
       parsed events) lets periodic refreshes read only the lines appended since the last call. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LogFolderPath, [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Server,
        [ValidateRange(1, 744)][int]$Hours = 24, [datetime]$Now = (Get-Date), [hashtable]$Cache = $null
    )
    if ($null -eq $Cache) { $Cache = @{} }
    $since = $Now.AddHours(-1 * $Hours)
    $events = [Collections.Generic.List[object]]::new(); $sources = [Collections.Generic.List[object]]::new()
    $usedKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($serverName in @($Server | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { [string]$_ } | Sort-Object -Unique)) {
        $folder = Join-Path $LogFolderPath $serverName
        $fileCount = 0; $failure = ''
        if (Test-Path -LiteralPath $folder -PathType Container) {
            for ($day = $since.Date; $day -le $Now.Date; $day = $day.AddDays(1)) {
                $path = Join-Path $folder ('SmartM365-Inventory-Orchestrator_{0}_{1}.log' -f $serverName, $day.ToString('yyyyMMdd'))
                if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
                [void]$usedKeys.Add($path)
                $entry = $Cache[$path]
                if ($null -eq $entry) { $entry = @{ Offset = [long]0; Events = [Collections.Generic.List[object]]::new() }; $Cache[$path] = $entry }
                try {
                    $tail = Read-InsightsLogTail -Path $path -Offset $entry.Offset
                    if ($tail.Reset) { $entry.Events.Clear() }
                    foreach ($line in $tail.Lines) {
                        $match = $script:InsightsLogLinePattern.Match([string]$line)
                        if (-not $match.Success) { continue }
                        $level = switch ($match.Groups['Level'].Value) { 'ERROR' { 'ERROR' } { $_ -in @('WARN', 'WARNING') } { 'WARN' } default { 'INFO' } }
                        $message = $match.Groups['Message'].Value.TrimEnd()
                        $category = Get-SmartM365OrchestratorLogEventCategory -Level $level -Message $message
                        if ($null -eq $category) { continue }
                        $time = [datetime]::ParseExact($match.Groups['Time'].Value, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
                        $entry.Events.Add([pscustomobject]@{ Time = $time; Server = $serverName; Level = $level; Category = $category; Message = $message; Source = $path })
                    }
                    $entry.Offset = $tail.Offset
                    $fileCount++
                }
                catch { $failure = $_.Exception.Message }
                foreach ($logEvent in $entry.Events) { if ($logEvent.Time -ge $since -and $logEvent.Time -le $Now) { $events.Add($logEvent) } }
            }
        }
        $status = if ($failure) { 'Read error: ' + $failure } elseif (-not (Test-Path -LiteralPath $folder -PathType Container)) { 'Log folder not found' } elseif ($fileCount) { 'OK' } else { 'No log in period' }
        $sources.Add([pscustomobject]@{ Server = $serverName; Files = $fileCount; Status = $status })
    }
    foreach ($key in @($Cache.Keys)) { if (-not $usedKeys.Contains([string]$key)) { $Cache.Remove($key) } }
    [pscustomobject]@{
        Events = @($events | Sort-Object Time -Descending); Sources = @($sources); Since = $since; Hours = $Hours; ReadAt = $Now; Cache = $Cache
    }
}

function Select-SmartM365OrchestratorLogEvents {
    <# Filters events for display. GroupRepeats merges identical messages of one server. #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Events = @(), [string]$Server = '',
        [ValidateSet('All', 'ErrorsAndWarnings', 'ERROR', 'WARN', 'INFO')][string]$Level = 'All',
        [string]$Category = '', [string]$Text = '', [datetime]$Since = [datetime]::MinValue,
        [switch]$GroupRepeats, [ValidateRange(1, 100000)][int]$MaxRows = 5000
    )
    $matching = @(foreach ($logEvent in @($Events | Where-Object { $null -ne $_ })) {
        if ($Server -and $logEvent.Server -ne $Server) { continue }
        if ($Level -eq 'ErrorsAndWarnings' -and $logEvent.Level -notin @('ERROR', 'WARN')) { continue }
        if ($Level -in @('ERROR', 'WARN', 'INFO') -and $logEvent.Level -ne $Level) { continue }
        if ($Category -and $logEvent.Category -ne $Category) { continue }
        if ($logEvent.Time -lt $Since) { continue }
        if ($Text -and $logEvent.Message.IndexOf($Text, [StringComparison]::OrdinalIgnoreCase) -lt 0 -and $logEvent.Category.IndexOf($Text, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        $logEvent
    })
    $rows = [Collections.Generic.List[object]]::new()
    if ($GroupRepeats) {
        $groups = [ordered]@{}
        foreach ($logEvent in $matching) {
            $key = '{0}|{1}|{2}|{3}' -f $logEvent.Server, $logEvent.Level, $logEvent.Category, $logEvent.Message
            $group = $groups[$key]
            if ($null -eq $group) { $groups[$key] = @{ Last = $logEvent; First = $logEvent.Time; Count = 1 }; continue }
            $group.Count++
            if ($logEvent.Time -gt $group.Last.Time) { $group.Last = $logEvent }
            if ($logEvent.Time -lt $group.First) { $group.First = $logEvent.Time }
        }
        foreach ($group in $groups.Values) { $rows.Add([pscustomobject]@{ Event = $group.Last; First = $group.First; Count = $group.Count }) }
    }
    else { foreach ($logEvent in $matching) { $rows.Add([pscustomobject]@{ Event = $logEvent; First = $logEvent.Time; Count = 1 }) } }
    $sorted = @($rows | Sort-Object { $_.Event.Time } -Descending)
    [pscustomobject]@{
        Total = $sorted.Count; Truncated = $sorted.Count -gt $MaxRows; Matching = $matching.Count
        Rows = @($sorted | Select-Object -First $MaxRows | ForEach-Object {
            [pscustomobject]@{
                Time = $_.Event.Time.ToString('yyyy-MM-dd HH:mm:ss'); Server = $_.Event.Server; Level = $_.Event.Level; Category = $_.Event.Category
                Count = $_.Count; First = $_.First.ToString('yyyy-MM-dd HH:mm:ss'); Message = $_.Event.Message; Source = $_.Event.Source
            }
        })
    }
}

function Get-SmartM365OrchestratorScriptVersion {
    <#
    Reads the .VERSION header of each job script (first 8 KB only, never parsed or executed).
    Returns a case-insensitive map ScriptPath -> version text: the version, 'missing' when the file
    is absent, '?' when no valid .VERSION header is found. The optional cache (FullPath -> entry)
    makes a later call re-read only the scripts whose size or last write time changed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SmartInventoryRoot,
        [AllowEmptyCollection()][string[]]$ScriptPath = @(),
        [AllowNull()][System.Collections.IDictionary]$Cache = $null
    )
    $versions = @{}
    $pattern = [regex]'(?im)^[ \t]*\.VERSION(?:[ \t]+(?<Inline>[^\r\n]+)|[ \t]*\r?\n[ \t]*(?<Version>[^\r\n]+))'
    foreach ($relative in @($ScriptPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $key = $relative.Trim()
        if ($versions.ContainsKey($key)) { continue }
        $version = '?'
        try {
            $fullPath = [IO.Path]::GetFullPath((Join-Path $SmartInventoryRoot $key))
            $file = [IO.FileInfo]::new($fullPath)
            if (-not $file.Exists) { $version = 'missing' }
            elseif ($file.Extension -eq '.ps1') {
                $stamp = '{0}|{1}' -f $file.Length, $file.LastWriteTimeUtc.Ticks
                if ($null -ne $Cache -and $Cache.Contains($fullPath) -and $Cache[$fullPath].Stamp -eq $stamp) { $version = [string]$Cache[$fullPath].Version }
                else {
                    $stream = [IO.FileStream]::new($fullPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
                    try {
                        $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true)
                        $buffer = New-Object char[] 8192
                        $count = $reader.ReadBlock($buffer, 0, $buffer.Length)
                    }
                    finally { $stream.Dispose() }
                    $match = $pattern.Match([string]::new($buffer, 0, $count))
                    if ($match.Success) {
                        $text = if ($match.Groups['Inline'].Success) { $match.Groups['Inline'].Value } else { $match.Groups['Version'].Value }
                        if ($text.Trim() -match '^[0-9][0-9A-Za-z.+-]{0,31}$') { $version = $text.Trim() }
                    }
                    if ($null -ne $Cache) { $Cache[$fullPath] = [pscustomobject]@{ Stamp = $stamp; Version = $version } }
                }
            }
        }
        catch { $version = '?' }
        $versions[$key] = $version
    }
    return $versions
}

# Producer script of every SWI input without its own source receipt. Receipt producers come from
# the CMDB source registry and the Intune operational contract. Test-SmartM365OrchestratorSwiValidityOffline
# keeps this map aligned with the SWI contracts and the jobs template.
$script:SwiFileProducers = [ordered]@{
    'SmartM365-ActiveDirectory-HealthCheck.ps1' = @('AD_HealthCheck.csv')
    'SmartM365-ActiveDirectory-Inventory.ps1' = @('AD_Users_DuplicateSMTP.csv', 'AD_Users_DuplicateUPN.csv')
    'SmartM365-AuthenticationMethodsRegistration-Inventory.ps1' = @('M365_Entra_AuthenticationMethodsRegistration.csv')
    'SmartM365-AzureADConnect-SyncHealth-Inventory.ps1' = @('M365_Entra_AzureADConnect_SyncHealth.csv')
    'SmartM365-ConditionalAccess-Inventory.ps1' = @('M365_Entra_ConditionalAccessPolicies.csv', 'M365_Entra_ConditionalAccessPolicyTargets.csv')
    'SmartM365-CopilotUsage-Inventory.ps1' = @('M365_Copilot_UserCountTrend.csv', 'M365_Copilot_UserUsage.csv')
    'SmartM365-Device-System-Inventory.ps1' = @('Intune_Devices_LocalSystem.csv')
    'SmartM365-Devices-Compliance-Inventory.ps1' = @('Intune_Devices_Compliance_Policies.csv')
    'SmartM365-Devices-Inventory.ps1' = @('Intune_Devices_Inventory.csv')
    'SmartM365-EXO-AcceptedDomains-Inventory.ps1' = @('Exchange_EXO_AcceptedDomains.csv')
    'SmartM365-EXO-Mailboxes-Inventory.ps1' = @('Exchange_EXO_Mailboxes_AllDomains_Archive.csv', 'Exchange_EXO_Mailboxes_AllDomains_Permissions.csv', 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv')
    'SmartM365-EXO-MigJob-Inventory.ps1' = @('Exchange_EXO_MigrationJobs.csv')
    'SmartM365-EndpointAnalytics-Inventory.ps1' = @('Intune_EndpointAnalytics_StartupDevices.csv')
    'SmartM365-EntraDevices-Inventory.ps1' = @('M365_Entra_Devices.csv', 'M365_Entra_Devices_HardwareIdConflicts.csv', 'M365_Entra_Devices_RemovalCandidates.csv')
    'SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1' = @('Exchange_Hybrid_FreeBusy_Findings.csv', 'Exchange_Hybrid_FreeBusy_Summary.csv')
    'SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1' = @('Exchange_OnPrem_HealthChecker_Findings.csv', 'Exchange_OnPrem_HealthChecker_Summary.csv', 'Exchange_OnPrem_Infrastructure_PerServerSummary.csv', 'Exchange_OnPrem_MailboxDatabases_Paths.csv', 'Exchange_OnPrem_MigrationReadiness_Config.csv', 'Exchange_OnPrem_Servers_Compute.csv', 'Exchange_OnPrem_Servers_DiskDrives.csv', 'Exchange_OnPrem_Servers_Inventory.csv', 'Exchange_OnPrem_Servers_Inventory_Summary.csv', 'Exchange_OnPrem_Servers_LogicalDisks.csv', 'Exchange_OnPrem_Servers_RemoteAccess.csv', 'Exchange_OnPrem_Servers_ServiceHealth.csv')
    'SmartM365-Intune-EndpointSecurityHealth-Inventory.ps1' = @('Intune_EndpointSecurity_DefenderAgents.csv', 'Intune_EndpointSecurity_FirewallStatus.csv')
    'SmartM365-Licences-Inventory.ps1' = @('M365_Licenses_ReportAssignmentIssues.csv', 'M365_Licenses_ReportCandidates.csv', 'M365_Licenses_ReportGaps.csv', 'M365_Licenses_ReportSkuCandidates.csv', 'M365_Licenses_ReportSkuCapacity.csv', 'M365_Licenses_ReportSnapshot.json.txt', 'M365_Licenses_ReportSources.csv', 'M365_Licenses_ReportSummary.csv')
    'SmartM365-LicensePricing-Inventory.ps1' = @('M365_License_Prices.csv')
    'SmartM365-M365-BackupPolicyScope-Inventory.ps1' = @('M365_BackupPolicyScope_GroupMembers.csv', 'M365_BackupPolicyScope_MailboxCoverage.csv')
    'SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1' = @('M365_Backup_ProtectedMailboxes.csv')
    'SmartM365-M365-BackupProtectedSitesAndDrives-Inventory.ps1' = @('M365_Backup_ProtectedOneDriveAccounts.csv', 'M365_Backup_ProtectedSharePointSites.csv')
    'SmartM365-M365UserActivity-Inventory.ps1' = @('M365_Apps_Activations.csv', 'M365_Apps_Usage_180D.csv', 'M365_Apps_Usage_30D.csv', 'M365_Email_Activity.csv', 'M365_Email_AppUsage.csv', 'M365_Groups_Activity.csv', 'M365_Mailbox_Usage.csv', 'M365_OneDrive_Usage.csv', 'M365_OneDrive_UserActivity.csv', 'M365_SharePoint_UserActivity.csv', 'M365_Teams_UserActivity.csv')
    'SmartM365-Mailboxes-AllSources-PermissionsByUserReport.ps1' = @('Exchange_Mailboxes_AllSources_PermissionsByUser.csv')
    'SmartM365-SPO-Inventory.ps1' = @('M365_SPO_Tenant.csv')
    'SmartM365-SecureScore-Inventory.ps1' = @('M365_Security_SecureScore.csv')
    'SmartM365-SharePoint-OnPrem-Content-Inventory.ps1' = @('SharePoint_OnPrem_CollectionCoverage.csv', 'SharePoint_OnPrem_SiteAdministrators.csv', 'SharePoint_OnPrem_SiteCollections.csv', 'SharePoint_OnPrem_Webs.csv')
    'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.ps1' = @('SharePoint_OnPrem_ContentDatabases.csv', 'SharePoint_OnPrem_Farms.csv', 'SharePoint_OnPrem_Servers.csv', 'SharePoint_OnPrem_ServiceApplications.csv', 'SharePoint_OnPrem_WebApplicationZones.csv', 'SharePoint_OnPrem_WebApplications.csv')
    'SmartM365-Teams-Inventory.ps1' = @('M365_Teams_Channels.csv')
    'SmartM365-WinUpdate_Status_From_Intune.ps1' = @('Intune_WindowsUpdate_Status.csv')
}

function Get-SmartM365OrchestratorLatestFolderPath {
    <# DATA-LAST beside the shared DATA-ALL\Orchestrator folder, as the producer receipts use it. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)
    $shared = $SharedDataFolderPath.TrimEnd([char[]]'\/')
    $dataAll = Split-Path -Path $shared -Parent
    if ([IO.Path]::GetFileName($shared) -ne 'Orchestrator' -or [IO.Path]::GetFileName($dataAll) -ne 'DATA-ALL') {
        throw "DATA-LAST cannot be located: the shared folder is not <tenant>\DATA-ALL\Orchestrator ($SharedDataFolderPath)."
    }
    return (Join-Path (Split-Path -Path $dataAll -Parent) 'DATA-LAST')
}

function Get-InsightsSwiPrepareSettings {
    # Same values as WorkplaceEvidence-Prepare: its local configuration, else its template.
    param([string]$ConfigPath, [string]$TemplatePath)
    $maximum = 0; $overrides = @{}; $source = ''
    foreach ($candidate in @(@{ Path = $ConfigPath; Name = 'Prepare configuration' }, @{ Path = $TemplatePath; Name = 'Prepare template (no local configuration found)' })) {
        if ([string]::IsNullOrWhiteSpace([string]$candidate.Path)) { continue }
        if ($candidate.Path -like '*.template') {
            # The committed template is plain JSON outside the owned .json/.json.txt transport.
            if (-not (Test-Path -LiteralPath $candidate.Path -PathType Leaf)) { continue }
            $document = Get-Content -LiteralPath $candidate.Path -Raw | ConvertFrom-Json
        }
        else {
            if (-not (Get-SmartM365JsonReadPath $candidate.Path -Optional)) { continue }
            $document = (Read-SmartM365JsonDocument -Path $candidate.Path).Document
        }
        $value = Get-InsightsProperty $document 'PreparedMaxSourceAgeHours' $null
        $hours = 0
        if (-not [int]::TryParse([string]$value, [ref]$hours) -or $hours -lt 1 -or $hours -gt 8760) { continue }
        $maximum = $hours; $source = [string]$candidate.Name
        $table = Get-InsightsProperty $document 'PreparedSourceAgeOverrides' $null
        if ($null -ne $table) {
            foreach ($property in $table.PSObject.Properties) {
                $limit = 0
                if ([int]::TryParse([string]$property.Value, [ref]$limit) -and $limit -ge 1 -and $limit -le 8760) { $overrides[[string]$property.Name] = $limit }
            }
        }
        break
    }
    if ($maximum -eq 0) { $maximum = 168; $source = 'Prepare code default (no configuration or template found)' }
    [pscustomobject]@{ MaxSourceAgeHours = $maximum; AgeOverrides = $overrides; Source = $source }
}

function Read-InsightsSwiReceipt {
    param([string]$LatestFolderPath, [string]$ReceiptName, [hashtable]$Cache)
    if ($Cache.ContainsKey($ReceiptName)) { return $Cache[$ReceiptName] }
    $basePath = Join-Path $LatestFolderPath ($ReceiptName -replace '\.txt$', '')
    $receipt = $null
    if (Get-SmartM365JsonReadPath $basePath -Optional) {
        try { $receipt = (Read-SmartM365JsonDocument -Path $basePath).Document } catch { $receipt = 'Unreadable' }
    }
    $Cache[$ReceiptName] = $receipt
    return $receipt
}

function Get-SmartM365OrchestratorSwiValidity {
    <#
    Read-only SWI validity of each job's latest collection, with the same publication-age rules as
    WorkplaceEvidence-Prepare: DATA-LAST file write time (PreparedMaxSourceAgeHours and its
    overrides, or the optional file limit), producer receipt start time for native sources
    (native evidence contract) and receipt completion time for the Intune operational groups.
    A job's validity is the earliest expiry among its required inputs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LatestFolderPath,
        [Parameter(Mandatory)]$JobsDocument,
        [Parameter(Mandatory)][string]$SwiConfigRoot,
        [Parameter(Mandatory)][string]$RegistryPath,
        [string]$PrepareConfigPath = '',
        [string]$PrepareTemplatePath = '',
        [datetimeoffset]$NowUtc = [datetimeoffset]::UtcNow
    )
    $readContract = { param($Name) (Read-SmartM365JsonDocument -Path (Join-Path $SwiConfigRoot $Name)).Document }
    $source = & $readContract 'prepared-source-contract.json.txt'
    $artifact = & $readContract 'prepared-artifact-contract.json.txt'
    $native = & $readContract 'native-evidence-contract.json.txt'
    $intune = & $readContract 'intune-operational-contract.json.txt'
    $registry = (Read-SmartM365JsonDocument -Path $RegistryPath).Document
    $settings = Get-InsightsSwiPrepareSettings -ConfigPath $PrepareConfigPath -TemplatePath $PrepareTemplatePath

    # Producer script of each input file.
    $producerByFile = @{}
    foreach ($script in $script:SwiFileProducers.Keys) { foreach ($file in $script:SwiFileProducers[$script]) { $producerByFile[$file] = [string]$script } }
    $receiptByFile = @{}
    foreach ($producer in @($registry.Producers)) {
        foreach ($file in @($producer.Files)) { $producerByFile[[string]$file] = [string]$producer.Script; $receiptByFile[[string]$file] = [string]$producer.Receipt }
    }
    foreach ($group in @($intune.groups)) { foreach ($entry in @($group.files)) { $producerByFile[[string]$entry.file] = [string]$group.producer } }

    # Rules, as Get-PreparedSourcePlan, the native freshness policy and the Intune source groups apply them.
    $rules = [Collections.Generic.List[object]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $limitFor = { param($File) if ($settings.AgeOverrides.ContainsKey($File)) { [int]$settings.AgeOverrides[$File] } else { [int]$settings.MaxSourceAgeHours } }
    foreach ($file in @(@($source.currentFiles) + @($artifact.currentFiles))) {
        if ($seen.Add('File|' + $file)) { $rules.Add([pscustomobject]@{ File = [string]$file; Kind = 'File'; LimitHours = (& $limitFor $file); WarningHours = $null; Optional = $false; Group = '' }) }
    }
    foreach ($entry in @(@(Get-InsightsProperty $source 'optionalCurrentFiles' @()) + @(Get-InsightsProperty $artifact 'optionalCurrentFiles' @()))) {
        if ($seen.Add('File|' + $entry.file)) { $rules.Add([pscustomobject]@{ File = [string]$entry.file; Kind = 'File'; LimitHours = [int]$entry.maxAgeHours; WarningHours = $null; Optional = $true; Group = '' }) }
    }
    $freshnessGroup = @{}
    foreach ($group in @(Get-InsightsProperty $native 'freshnessGroups' @())) { foreach ($file in @($group.sources)) { $freshnessGroup[[string]$file] = $group } }
    foreach ($entry in @($native.sources)) {
        $file = [string]$entry.file
        $group = if ($freshnessGroup.ContainsKey($file)) { $freshnessGroup[$file] } else { $null }
        $limit = if ($group) { [double]$group.maxAgeHours } else { [double]$native.maxAgeHours }
        $warning = if ($group) { [double]$group.warningAgeHours } else { $null }
        $rules.Add([pscustomobject]@{ File = $file; Kind = 'Receipt'; LimitHours = $limit; WarningHours = $warning; Optional = $false; Group = $(if ($group) { [string]$group.name } else { 'Core' }) })
    }
    foreach ($group in @($intune.groups)) {
        foreach ($entry in @($group.files)) {
            $rules.Add([pscustomobject]@{ File = [string]$entry.file; Kind = 'IntuneReceipt'; LimitHours = (& $limitFor ([string]$entry.file)); WarningHours = $null; Optional = $true; Group = [string]$group.name })
        }
    }

    # Evaluation of each rule against DATA-LAST and the receipts.
    $receipts = @{}
    $intuneGroups = @{}
    foreach ($group in @($intune.groups)) {
        $receipt = Read-InsightsSwiReceipt -LatestFolderPath $LatestFolderPath -ReceiptName ([string]$group.receipt) -Cache $receipts
        $present = @($group.files | Where-Object { Test-Path -LiteralPath (Join-Path $LatestFolderPath $_.file) -PathType Leaf }).Count
        $intuneGroups[[string]$group.name] = [pscustomobject]@{ Receipt = $receipt; Present = $present; Expected = @($group.files).Count }
    }
    $evaluations = foreach ($rule in $rules) {
        $timestamp = $null; $problem = ''; $basis = ''; $optional = $rule.Optional
        switch ($rule.Kind) {
            'File' {
                $basis = 'DATA-LAST file time'
                $item = Get-Item -LiteralPath (Join-Path $LatestFolderPath $rule.File) -ErrorAction SilentlyContinue
                if ($null -eq $item -or $item.PSIsContainer) { $problem = 'Missing' }
                elseif ($item.Length -eq 0) { $problem = 'Empty' }
                else { $timestamp = [datetimeoffset]$item.LastWriteTimeUtc }
            }
            'Receipt' {
                $basis = 'collector receipt start'
                $name = if ($receiptByFile.ContainsKey($rule.File)) { $receiptByFile[$rule.File] } else { '' }
                $receipt = if ($name) { Read-InsightsSwiReceipt -LatestFolderPath $LatestFolderPath -ReceiptName $name -Cache $receipts } else { $null }
                if ($null -eq $receipt) { $problem = 'Missing receipt' }
                elseif ($receipt -is [string]) { $problem = 'Unreadable receipt' }
                elseif ([string](Get-InsightsProperty $receipt 'Status' '') -ne 'Completed' -or [int](Get-InsightsProperty $receipt 'Errors' 1) -ne 0) { $problem = 'Receipt not completed' }
                else {
                    $record = @(Get-InsightsProperty $receipt 'Files' @() | Where-Object { [string](Get-InsightsProperty $_ 'File' '') -eq $rule.File } | Select-Object -First 1)
                    $started = if ($record.Count -and (Get-InsightsProperty $record[0] 'StartedAtUtc' $null)) { Get-InsightsProperty $record[0] 'StartedAtUtc' $null } else { Get-InsightsProperty $receipt 'StartedAtUtc' $null }
                    if ($record.Count -eq 0) { $problem = 'File absent from receipt' }
                    elseif ($null -eq $started) { $problem = 'Receipt without start time' }
                    else { $timestamp = [datetimeoffset]::Parse((ConvertTo-InsightsSwiTimeText $started), [Globalization.CultureInfo]::InvariantCulture) }
                }
            }
            'IntuneReceipt' {
                $basis = 'collector receipt completion'
                $state = $intuneGroups[$rule.Group]
                if ($null -eq $state.Receipt -and $state.Present -eq 0) { $problem = 'Not collected'; $optional = $true }
                elseif ($null -eq $state.Receipt -or $state.Receipt -is [string] -or $state.Present -ne $state.Expected) { $problem = 'Incomplete group' }
                elseif ([string](Get-InsightsProperty $state.Receipt 'Status' '') -ne 'Completed' -or [int](Get-InsightsProperty $state.Receipt 'Errors' 1) -ne 0) { $problem = 'Receipt not completed' }
                else { $timestamp = [datetimeoffset]::Parse((ConvertTo-InsightsSwiTimeText (Get-InsightsProperty $state.Receipt 'CompletedAtUtc' '')), [Globalization.CultureInfo]::InvariantCulture) }
            }
        }
        if ($null -ne $timestamp -and $timestamp -gt $NowUtc.AddMinutes(5)) { $problem = 'Future timestamp'; $timestamp = $null }
        $expires = if ($null -ne $timestamp) { $timestamp.AddHours([double]$rule.LimitHours) } else { $null }
        $hoursLeft = if ($null -ne $expires) { ($expires - $NowUtc).TotalHours } else { $null }
        # Prepare keeps an optional source past its limit as the latest available data (stale, still used).
        $state = if ($problem) { $problem }
            elseif ($hoursLeft -lt 0) { if ($optional) { 'Old' } else { 'Expired' } }
            elseif ($null -ne $rule.WarningHours -and ($NowUtc - $timestamp).TotalHours -gt $rule.WarningHours) { 'Aging' }
            else { 'Valid' }
        [pscustomobject]@{
            File = $rule.File; Kind = $rule.Kind; Basis = $basis; LimitHours = $rule.LimitHours; Optional = [bool]$optional
            Producer = if ($producerByFile.ContainsKey($rule.File)) { $producerByFile[$rule.File] } else { '' }
            TimestampUtc = $timestamp; ExpiresAtUtc = $expires; HoursLeft = $hoursLeft; State = $state
            # An incomplete or unqualified Intune group is still refused although the group is optional.
            Blocks = ((-not $optional -and $state -notin @('Valid', 'Aging')) -or ($rule.Kind -eq 'IntuneReceipt' -and $state -in @('Incomplete group', 'Receipt not completed')))
        }
    }

    # One row per job: its scripts' inputs, the most urgent required one decides.
    $jobsByScript = @{}
    foreach ($job in @($JobsDocument.Jobs)) {
        $scriptName = [IO.Path]::GetFileName([string](Get-InsightsProperty $job 'ScriptPath' ''))
        if (-not $jobsByScript.ContainsKey($scriptName)) { $jobsByScript[$scriptName] = [Collections.Generic.List[string]]::new() }
        $jobsByScript[$scriptName].Add([string]$job.Name)
    }
    $rows = foreach ($job in @($JobsDocument.Jobs)) {
        $scriptName = [IO.Path]::GetFileName([string](Get-InsightsProperty $job 'ScriptPath' ''))
        $inputs = @($evaluations | Where-Object Producer -eq $scriptName)
        if ($inputs.Count -eq 0) {
            [pscustomobject]@{ Name = [string]$job.Name; Swi = 'Not used by SWI'; Blocks = $false; DaysLeft = $null; ExpiresAt = ''; Detail = 'No SWI input is produced by this job.'; Inputs = @() }
            continue
        }
        $required = @($inputs | Where-Object { -not $_.Optional })
        $pool = if ($required.Count) { $required } else { $inputs }
        $limiting = @($pool | Sort-Object @{ Expression = { if ($_.Blocks) { 0 } elseif ($_.State -notin @('Valid', 'Aging')) { 1 } else { 2 } } }, @{ Expression = { if ($null -ne $_.HoursLeft) { $_.HoursLeft } else { [double]::MinValue } } } | Select-Object -First 1)[0]
        $status = if ($limiting.State -in @('Valid', 'Aging')) {
                if ($limiting.HoursLeft -lt 24) { 'Expires < 24 h' } elseif ($limiting.State -eq 'Aging') { 'Aging (accepted)' } else { 'Valid' }
            }
            elseif ($limiting.Blocks) { '{0} - blocks SWI' -f $limiting.State }
            elseif ($limiting.State -eq 'Old') { 'Old - still used (optional)' }
            else { '{0} (optional)' -f $limiting.State }
        $expiresText = if ($null -ne $limiting.ExpiresAtUtc) { $limiting.ExpiresAtUtc.LocalDateTime.ToString('yyyy-MM-dd HH:mm') } else { '' }
        $shared = if ($jobsByScript[$scriptName].Count -gt 1) { (' Shared output of {0} ({1}).' -f $scriptName, ($jobsByScript[$scriptName] -join ', ')) } else { '' }
        $otherIssues = @($inputs | Where-Object { $_ -ne $limiting -and $_.State -notin @('Valid', 'Aging') } | ForEach-Object { '{0}: {1}{2}' -f $_.File, $_.State, $(if ($_.Optional) { ' (optional)' } else { '' }) })
        $detail = ('Limiting input: {0}; rule: {1} + {2} h{3}; expires {4}.{5}' -f $limiting.File, $limiting.Basis, $limiting.LimitHours, $(if ($limiting.Optional) { ' (optional input)' } else { '' }), $(if ($expiresText) { $expiresText } else { 'n/a' }), $shared)
        if ($otherIssues.Count) { $detail += ' Other inputs: ' + ($otherIssues -join '; ') + '.' }
        [pscustomobject]@{
            Name = [string]$job.Name; Swi = $status; Blocks = [bool]$limiting.Blocks
            DaysLeft = if ($null -ne $limiting.HoursLeft) { [math]::Round($limiting.HoursLeft / 24, 1) } else { $null }
            ExpiresAt = $expiresText; Detail = $detail; Inputs = $inputs
        }
    }
    [pscustomobject]@{
        Rows = @($rows); Inputs = @($evaluations)
        Unmapped = @($evaluations | Where-Object { -not $_.Producer -or -not $jobsByScript.ContainsKey($_.Producer) } | ForEach-Object File | Sort-Object -Unique)
        MaxSourceAgeHours = $settings.MaxSourceAgeHours; SettingsSource = $settings.Source
        LatestFolderPath = $LatestFolderPath; EvaluatedAtUtc = $NowUtc
    }
}

function Get-SmartM365OrchestratorSwiValidityForRoot {
    <# Resolves the deployed inputs as WorkplaceEvidence-Prepare does, then evaluates the validity. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)][string]$SmartM365RootPath, [Parameter(Mandatory)]$JobsDocument)
    $prepareFolder = Join-Path $SmartM365RootPath 'SmartInventory/PreparedEvidence'
    $prepareConfig = Join-Path $prepareFolder 'SmartM365-WorkplaceEvidence-Prepare.local.json'
    # Prepare's WorkplaceIntelligenceRootPath when it is a plain path, else the sibling SmartWorkplaceIntelligence.
    $product = Join-Path (Split-Path -Path $SmartM365RootPath -Parent) 'SmartWorkplaceIntelligence'
    try {
        if (Get-SmartM365JsonReadPath $prepareConfig -Optional) {
            $configured = [string](Get-InsightsProperty (Read-SmartM365JsonDocument -Path $prepareConfig).Document 'WorkplaceIntelligenceRootPath' '')
            if ($configured -and $configured -notmatch '\{\{|USE_GLOBAL') { $product = $configured }
        }
    }
    catch { Write-Verbose ("Prepare configuration unreadable; default SmartWorkplaceIntelligence location used: {0}" -f $_.Exception.Message) }
    Get-SmartM365OrchestratorSwiValidity -LatestFolderPath (Get-SmartM365OrchestratorLatestFolderPath -SharedDataFolderPath $SharedDataFolderPath) `
        -JobsDocument $JobsDocument -SwiConfigRoot (Join-Path $product 'config') `
        -RegistryPath (Join-Path $SmartM365RootPath 'Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt') `
        -PrepareConfigPath $prepareConfig -PrepareTemplatePath (Join-Path $prepareFolder 'SmartM365-WorkplaceEvidence-Prepare.local.json.template')
}

function ConvertTo-InsightsSwiTimeText {
    param([AllowNull()]$Value)
    if ($Value -is [datetime]) { return ([datetime]$Value).ToUniversalTime().ToString('o') }
    if ($Value -is [datetimeoffset]) { return ([datetimeoffset]$Value).UtcDateTime.ToString('o') }
    [string]$Value
}

Export-ModuleMember -Function @(
    'Get-SmartM365OrchestratorScheduleOccurrences', 'Get-SmartM365OrchestratorExpectedMaxAgeHours', 'Get-SmartM365OrchestratorRecentRuns',
    'Get-SmartM365OrchestratorJobHealth', 'Get-SmartM365OrchestratorDependents', 'Get-SmartM365OrchestratorDependencyReadiness',
    'Get-SmartM365OrchestratorOperations', 'Get-SmartM365OrchestratorServerMetrics', 'Get-SmartM365OrchestratorRecentPipelineRuns', 'New-SmartM365OrchestratorJobRunRequest',
    'Get-SmartM365OrchestratorLogEventCategory', 'Get-SmartM365OrchestratorLogEvents', 'Select-SmartM365OrchestratorLogEvents',
    'Get-SmartM365OrchestratorDashboard', 'Get-SmartM365OrchestratorDashboardData',
    'Get-SmartM365OrchestratorScriptVersion', 'Get-SmartM365OrchestratorLatestFolderPath', 'Get-SmartM365OrchestratorSwiValidity',
    'Get-SmartM365OrchestratorSwiValidityForRoot'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDT6/+gSwWbMvij
# Oa4KwYfxHHCe8TKQyMgx7Tv8StbOqaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBJp18A9O8bSTIPEacIvfE3CsfIgrN/BJuQ2S5iSnDPbMA0GCSqG
# SIb3DQEBAQUABIIBgFo1pos22YjLN3JRb+PCs9++07mGzV0xicUEHsCDUkc5lPp1
# XKIK964JJyHMb9xkq+K6VarhCy0B8JeVJohBLiYyeuZ0deYPGLNuxh2Susli+lCm
# MvpR13fdvpfRM6D6q/WzIu+zIQfEYi9PG/rIhN3sEfb+5FgRqiPdYjEHqSaz5slf
# 96qzJ/SU2VD6cXqJkjpVTAYXLwtHBVSYeBZ40f8v1NyI+aBeoQk/0u62sDMBr6Ej
# 0AN4tUTmx5BkCWD7zU4q3rgT/x98ehW150i3qxbesHLtg4W0ofVhqwDC7B4hmUWn
# BW4sMKcstMc88+QNhYs9KrGZNpQrpONUkq30ZEajeOZvJjO25M5XTzkTp4AU8VPL
# 9EW0wioaSnPSMHXOsI4kltlebHo0sj7muqla9C5fDfgESpXxs0GQJpkb6/fv4sbq
# hs3hIamZCXFkqjS+DlPW5jivHAFVwmUqQWKlv4VSQe8HY+FhHzZEeDOKjsgJdxLq
# UgqHFoB4h2kEJ4ge4aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgyMTA2
# MThaMC8GCSqGSIb3DQEJBDEiBCDrAkhXiR0FvxlOhJ/BG/tWN5JSdUzAQKIAa8DX
# C5pmxTANBgkqhkiG9w0BAQEFAASCAgB6uE3ozFzUUtcVtbAFFVHanlxsyJ9xWhU1
# G6PRyALd4ZihmRHb5iLcnyPuXcKZhXExyWgElW4fL7mUysCWJ0ZpLh9FCjzs9gf6
# JbEfuaafN/+IQfdFHwYbNWEh8VIARNKI2zSpS7jK53unI96mN87eLHWu5/0ZZYPJ
# uKnxS1pkBGo+CV/az7E/JxkkLQkbdaKa19T9o8VCDvt529QOmjQRVH99ifVEaGW1
# NaZz2e5ZmM4RPWKEPvq1a7YfVwlIVIXHAcN6tNObAhyyD9SfFTQoNdju2ulH25zN
# FFekqJo+eHBNJHUNe4CqrfWZBt9KoJkPOYEWjrxgjRjHWVauuw5Fv/dZxAcK11yv
# lcz9sE/MVDOQcpN8UFmIrOaS6EK9+tGpD2ruTb8E535O3dtxphMnMW50rISInCTr
# K5dc2knFKKHJgeBKNOrmLxL1TaQZPEMMz4D4C8BfDWE/MKIDAERvIpjoYaILJ4U1
# Hes0Nk4g4YVLrg26EQvbAvu9eIMI3m3LZ9OJi5zjpZLTqgpTH4QeJ6nVq4oFixKw
# xdBe2exI2QHZlJZhznV4FwudPBCbcs9qMullIniI3qy/2A6E2WvP7SVCnp0QfIFr
# vS1HjaKRHRQR5NFVSxM9G4Lt2RiYnn3l8voi4G1gh+rNVE9OSoRBaPo4wemDIWiw
# ybBNVZrN2Q==
# SIG # End signature block
