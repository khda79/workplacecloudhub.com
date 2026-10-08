# Read-only operational insights for the SmartM365 Orchestrator GUI: live server state, job
# health, dependency readiness, pipeline requests and job run requests. No WPF code here, so
# every function can be tested headless against a copy of the shared Orchestrator folder.
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Pipeline.psm1') -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Maintenance.psm1') -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Distributed.psm1') -ErrorAction Stop

$script:InsightsTerminalSuccess = @('Success', 'CompletedWithWarnings')
$script:InsightsTerminalFailure = @('Failed', 'TimedOut', 'Interrupted')
# Same lists as SmartM365.Orchestrator.Pipeline.psm1 (Get-SmartM365OrchestratorPipelineRunStatus).
$script:InsightsPipelineFailure = @('Failed', 'TimedOut', 'Interrupted', 'BlockedDependencyFailed', 'BlockedDependencyTimeout', 'Rejected', 'MissingStatus')
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
            Maintenance = if ($maintenanceError) { 'Unknown' } elseif ($maintenance.Enabled) { 'ON' } else { 'OFF' }
            MaintenanceAcknowledgement = if ($maintenanceError) { 'Control unavailable' } else { [string](@($maintenanceServers | Where-Object Server -eq $serverName)[0].Status) }
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
        $cancellationRequested = $null -ne (Get-InsightsProperty $request 'Cancellation' $null)
        $created = ConvertTo-InsightsUtc (Get-InsightsProperty $request 'CreatedAtUtc' $null)
        [pscustomobject]@{
            BatchId = $folder.Name; Selection = [string](Get-InsightsProperty $request 'Pipeline' '')
            Created = if ($null -ne $created) { $created.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
            Status = if ($pendingCount -gt 0) { if ($cancellationRequested) { 'Cancelling' } else { 'Running' } } elseif ($failedCount -gt 0) { 'Failed' } elseif ($cancellationRequested) { 'Cancelled' } elseif ($warningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
            Cancelled = @($jobRows | Where-Object Status -eq 'Cancelled').Count
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDpZLDQf7y7lsFM
# SWvDrrw/8WyJKrA8J42yxImxDwsy9KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDF2Fm8LDfXVOg3pizbXln/
# s15k958hvw5clKa6UVtpszANBgkqhkiG9w0BAQEFAASCAYB3D5BxzwqP7aIEn3nz
# UjK8ZBkaFuUS3GK6hq4hLtC9dr5iEuj7vNxmVPVpu2otmG+zTsBTB0axyjcpfo31
# zRVDUx82mDjjixq7G2EujBGXxYMsQoB7N97S8Es7Nqcql0zEvT4SpT7tbi3pKrbT
# VeosA8FqHqPGuqM9EK5NRpek158DRbiX6KtXjHyIMuxGcyj87xLSikFddXYkbeCR
# 4xBsTfFJoh1knAXI3hE93YnzIUirXe3GDb4Gz7JEE7Dy93tvuay2jQJ4evpL1L5D
# imIo60NWufqHeBkqV+10SmI9GMIoOUFirPvCgYqjT+R/pvEMZsKKyX1LsgH0wOti
# pjq3MFoVtj7AsmeAuY6hYjr4eKn28O3eKRVZhyewaciiC4pZJiPcHLRC9q1bO1MX
# ykhHwIyZKppQsFgJI3UyRdi0xjgvhciqaCWOj8lOz+SQvfJZ0bMEIeC45WaO3ZsJ
# eXdfrzJlOQu+fLOLa6KclE3Ful/DJ6NgXf/XViGyOcF/jHU=
# SIG # End signature block
