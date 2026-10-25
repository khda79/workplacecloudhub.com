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
    $mode = [string](Get-InsightsProperty $job 'DependencyMode' 'LatestOccurrence')
    $override = [int](Get-InsightsProperty $job 'DependencyMaxAgeHours' 0)
    foreach ($dependencyName in @(Get-InsightsProperty $job 'DependsOn' @() | ForEach-Object { [string]$_ })) {
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
        $cancellationRequested = $null -ne (Get-InsightsProperty $request 'Cancellation' $null)
        $cancelledCount = @($jobRows | Where-Object Status -eq 'Cancelled').Count
        $failedCount = @($jobRows | Where-Object { $_.Status -in $script:InsightsPipelineFailure }).Count
        if (-not $cancellationRequested) { $failedCount += $cancelledCount }
        $warningCount = @($jobRows | Where-Object { $_.Status -eq 'CompletedWithWarnings' }).Count
        $created = ConvertTo-InsightsUtc (Get-InsightsProperty $request 'CreatedAtUtc' $null)
        [pscustomobject]@{
            BatchId = $folder.Name; Selection = [string](Get-InsightsProperty $request 'Pipeline' '')
            Created = if ($null -ne $created) { $created.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
            Status = if ($pendingCount -gt 0) { if ($cancellationRequested) { 'Cancelling' } else { 'Running' } } elseif ($failedCount -gt 0) { 'Failed' } elseif ($cancellationRequested) { 'Cancelled' } elseif ($warningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
            Cancelled = $cancelledCount
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

function Get-SmartM365OrchestratorDashboard {
    <# Pure dashboard aggregation (banner, tiles, items to review, servers, next 24 h timeline). #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$JobsDocument, [Parameter(Mandatory)]$Operations,
        [AllowEmptyCollection()][object[]]$Health = @(), [AllowEmptyCollection()][object[]]$Runs = @(),
        [AllowNull()]$Plan = $null, [AllowEmptyCollection()][object[]]$Requests = @(), [AllowNull()]$LogEvents = $null,
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
    foreach ($incident in @($Operations.Incidents)) { Add-DashboardAction 'Warning' 'Peer monitoring' ([string]$incident.ObservedBy) ([string]$incident.Issue) 'Operations' }
    foreach ($request in @($Requests | Where-Object { $_.Status -in @('Running', 'Cancelling') })) {
        $created = [datetime]::MinValue
        if ([datetime]::TryParseExact([string]$request.Created, 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$created) -and ($Now - $created).TotalHours -gt $RequestStaleHours) {
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
        Timeline = [pscustomobject]@{ From = $Now; To = $to; Lanes = $lanes; Items = @($items | Sort-Object Start) }
        GeneratedAt = $Now
    }
}

function Get-SmartM365OrchestratorDashboardData {
    <# Reads the remaining read-only sources (runs, election plan) and builds the dashboard. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)]$JobsDocument, [Parameter(Mandatory)]$Operations,
        [AllowNull()][object[]]$Runs = $null, [AllowEmptyCollection()][object[]]$Requests = @(), [AllowNull()]$LogEvents = $null, [datetime]$Now = (Get-Date)
    )
    if ($null -eq $Runs) { $Runs = @(Get-SmartM365OrchestratorRecentRuns -SharedDataFolderPath $SharedDataFolderPath -Days 21 -Now $Now) }
    $health = @(Get-SmartM365OrchestratorJobHealth -JobsDocument $JobsDocument -Runs $Runs -Now $Now)
    $plan = Read-InsightsJson -Path (Join-Path $SharedDataFolderPath 'Election\Orchestrator-ElectionPlan.json')
    Get-SmartM365OrchestratorDashboard -JobsDocument $JobsDocument -Operations $Operations -Health $health -Runs $Runs -Plan $plan -Requests $Requests -LogEvents $LogEvents -Now $Now
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

Export-ModuleMember -Function @(
    'Get-SmartM365OrchestratorScheduleOccurrences', 'Get-SmartM365OrchestratorExpectedMaxAgeHours', 'Get-SmartM365OrchestratorRecentRuns',
    'Get-SmartM365OrchestratorJobHealth', 'Get-SmartM365OrchestratorDependents', 'Get-SmartM365OrchestratorDependencyReadiness',
    'Get-SmartM365OrchestratorOperations', 'Get-SmartM365OrchestratorRecentPipelineRuns', 'New-SmartM365OrchestratorJobRunRequest',
    'Get-SmartM365OrchestratorLogEventCategory', 'Get-SmartM365OrchestratorLogEvents', 'Select-SmartM365OrchestratorLogEvents',
    'Get-SmartM365OrchestratorDashboard', 'Get-SmartM365OrchestratorDashboardData'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB77+qaOm3KQgBJ
# RcTjKrTLmDNPGS7omLifkRMUMgnSp6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDXD6en404PMYdUj+3R2K8WC5lGrjFlzGTDsSu2b0QVBMA0GCSqG
# SIb3DQEBAQUABIIBgARocaWOyIZz3zlCX+McSCHMSt7TVO3/AoBVL1P9Z0BbjYdy
# WlXGKOI/nOZLezhEdPcYdAxeozduUg+Ry8LcxN+soJCZPfSukxGohYUald5uvjgk
# zt2SFdxvY0oZaRt28epa4CYjMVsUV/LdrbIx99O2kSqAaUJQm0uTXX8KOV2j/m5f
# RhPSorq2YwjE3/Ka98v8y2dgHJtzO1yHEdXhct5PYqYI8pCCu6Hlf2C32d++S0tj
# m7XzMjrnNwOHEdYf3die6MZDiP/nwXdT6VrGPicEmSSKxYGUogV4kJt70kQ4bhg6
# bzcNP5EKNLk4GHg3arSNdTmv0VkxlnpGqMf8zDpVK0cP71LLRMPBxStPJNJkdw5B
# Y0+gZ8wj0A9/6JkXUWrw76TLUK3UyBiQnTcEIM5JB+3xPprdN0MLZreSj0Pzy2xr
# MJdKMOz73s1vY1PImoXJ6Gu/YsNNfdZAVDuUoKvZ28DiFaX5dTMtgcbl/WC+VvpH
# wfBrqf6FeAfs736qYaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUyMDIx
# NTdaMC8GCSqGSIb3DQEJBDEiBCDFrEXWhEIG48aR8SOTTXm7okVfBnlAI1jJ8OEe
# kLmBhzANBgkqhkiG9w0BAQEFAASCAgCNAoCbWvnK/OTqGVXPKFQGEYksroxOm7af
# eAeYEj5HxMwq/XZ8k6OAo5EJWhS6CoJBN2MUUCxNCGEfikEYzYZUEkXJITIUyu0e
# 2fxnns9NTwmQEANuWwKOAaj1T2fDNAUvekkaXwuW5oEpgGfHt9mDAS/fK1Id06CD
# lhLJrFfTsN77sajvL1LCOLHp0UYEV4MwxezFeMP7ksnWgR2lbY/poA1uZGusywdd
# YqiXYJQOV4YLrJtXOIgYfIR+oGU+0E7QmJJje5/tiDNQ7BoJ++ajgAjinMF3t94p
# Gs8yUjWUACbKxPwDLj5mqmL7b6fEP2jBx6nBrpg31cDvdW0ywjncWeQZ5cp5cNud
# N74CxBmPvatAtI2G/juhMzLzn7drgf57YTh9hrOKsxWgRLZ5ZDN3zJGOrJbwruEV
# 73ZzPm7bf5R18rVFYNDUm/gdAk66oXLSAS7Iv93AZd+LhzcdnRs8ItoTsB5mVDxY
# fEFX7nn8nhuk2qy0ydGq7RhST6mYS/T7/hz21UHbTdZMLqXSCDVzGmJBTXnXvdfU
# PllO6axCH68nu38ewlmA16hP0HXS7/S3lQq0eofbhfWs72GoFP9G+7uVKuMIKPAl
# mdmeUvCLAXdMIc2nv4g7D9807F0ztNAAzCZ0qDGUqrcBYBJ2qxVHOkYNk/jVp8fS
# fuq1FaiVZw==
# SIG # End signature block
