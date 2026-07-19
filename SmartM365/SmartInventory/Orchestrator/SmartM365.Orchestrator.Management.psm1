Set-StrictMode -Version 2.0

$script:ValidCapabilities = @('SharedRuntime', 'Graph', 'EXO', 'AD', 'ExchangeOnPrem', 'TeamsPowerShell')
$script:ValidDays = @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')

function ConvertTo-SmartM365OrchestratorHashtable {
    param([AllowNull()]$InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $result = @{}
        foreach ($key in $InputObject.Keys) { $result[[string]$key] = ConvertTo-SmartM365OrchestratorHashtable $InputObject[$key] }
        return $result
    }
    if ($InputObject -is [pscustomobject]) {
        $result = @{}
        foreach ($property in $InputObject.PSObject.Properties) { $result[$property.Name] = ConvertTo-SmartM365OrchestratorHashtable $property.Value }
        return $result
    }
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        return @($InputObject | ForEach-Object { ConvertTo-SmartM365OrchestratorHashtable $_ })
    }
    return $InputObject
}

function Get-SmartM365OrchestratorConfigurationPaths {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)
    $configFolder = Join-Path $SharedDataFolderPath 'Config'
    [pscustomobject]@{
        SharedDataFolderPath = $SharedDataFolderPath
        ConfigFolderPath = $configFolder
        JobsPath = Join-Path $configFolder 'Orchestrator-Jobs.json'
        ClusterPath = Join-Path $configFolder 'Orchestrator-Cluster.json'
        VersionsFolderPath = Join-Path $configFolder 'Versions'
        AuditFolderPath = Join-Path $SharedDataFolderPath 'Audit'
        AuditPath = Join-Path $SharedDataFolderPath 'Audit\Orchestrator_ConfigChanges.csv'
        LockPath = Join-Path $configFolder 'Orchestrator-Configuration.lock'
    }
}

function Get-SmartM365OrchestratorFileHash {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    [string](Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
}

function Read-SmartM365OrchestratorJson {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "JSON configuration file not found: $Path" }
    Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -Depth 100 -ErrorAction Stop
}

function Write-SmartM365OrchestratorJsonAtomically {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$Document, [switch]$CreateNew)
    $folder = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $content = ($Document | ConvertTo-Json -Depth 100) + [Environment]::NewLine
    if ($CreateNew) {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($content)
        $stream = [IO.File]::Open($Path, 'CreateNew', 'Write', 'Read')
        try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
        return
    }
    $temporaryPath = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [IO.File]::WriteAllText($temporaryPath, $content, [Text.UTF8Encoding]::new($false))
        [IO.File]::Move($temporaryPath, $Path, $true)
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
}

function Test-SmartM365OrchestratorJobsDocument {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Document)
    $errors = [Collections.Generic.List[string]]::new()
    $warnings = [Collections.Generic.List[string]]::new()
    if (-not $Document.PSObject.Properties['Jobs'] -or $null -eq $Document.Jobs) {
        $errors.Add("The jobs document must contain a 'Jobs' array.")
        return [pscustomobject]@{ Valid = $false; Errors = @($errors); Warnings = @(); JobCount = 0 }
    }
    $seen = @{}; $dependencies = @{}
    foreach ($job in @($Document.Jobs)) {
        $name = if ($job.PSObject.Properties['Name']) { ([string]$job.Name).Trim() } else { '' }
        if (-not $name) { $errors.Add('A job has an empty Name.'); continue }
        if ($name -notmatch '^[A-Za-z0-9._-]+$') { $errors.Add("Job '$name': invalid Name.") }
        if ($seen.ContainsKey($name)) { $errors.Add("Duplicate job name: $name") } else { $seen[$name] = $true }
        if (-not $job.PSObject.Properties['ScriptPath'] -or -not [string]$job.ScriptPath) { $errors.Add("Job '$name': ScriptPath is required.") }
        $mode = if ($job.PSObject.Properties['AssignmentMode'] -and $job.AssignmentMode) { [string]$job.AssignmentMode } else { 'Legacy' }
        if ($mode -notin @('Legacy', 'Pinned', 'Elected', 'Manual')) { $errors.Add("Job '$name': invalid AssignmentMode.") }
        $allowed = if ($job.PSObject.Properties['AllowedServers']) { @($job.AllowedServers | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ }) } else { @() }
        if ($mode -eq 'Pinned' -and @($allowed).Count -ne 1) { $errors.Add("Job '$name': Pinned requires exactly one AllowedServers value.") }
        if ($mode -in @('Elected', 'Manual') -and @($allowed).Count -gt 0) { $warnings.Add("Job '$name': AllowedServers is ignored in $mode mode.") }
        $capabilities = if ($job.PSObject.Properties['RequiredCapabilities']) { @($job.RequiredCapabilities) } else { @() }
        foreach ($capability in $capabilities) { if ([string]$capability -notin $script:ValidCapabilities) { $errors.Add("Job '$name': unknown capability '$capability'.") } }
        $roles = if ($job.PSObject.Properties['RequiredGraphAppRoles']) { @($job.RequiredGraphAppRoles) } else { @() }
        if (@($roles).Count -gt 0 -and 'Graph' -notin $capabilities) { $errors.Add("Job '$name': RequiredGraphAppRoles requires Graph.") }
        if ($job.PSObject.Properties['EstimatedDurationMinutes'] -and [double]$job.EstimatedDurationMinutes -le 0) { $errors.Add("Job '$name': EstimatedDurationMinutes must be greater than zero.") }
        foreach ($propertyName in @('TimeoutMinutes', 'MaxRetries', 'RetryDelaySeconds', 'MinimumSuccessDurationSeconds', 'DependencyWaitTimeoutMinutes')) {
            if ($job.PSObject.Properties[$propertyName] -and [double]$job.$propertyName -lt 0) { $errors.Add("Job '$name': $propertyName cannot be negative.") }
        }
        if (-not $job.PSObject.Properties['Schedule'] -or $null -eq $job.Schedule) { $errors.Add("Job '$name': Schedule is required.") }
        else {
            $type = if ($job.Schedule.PSObject.Properties['Type']) { [string]$job.Schedule.Type } else { '' }
            if ($type -notin @('Daily', 'Weekly')) { $errors.Add("Job '$name': Schedule.Type must be Daily or Weekly.") }
            $times = if ($job.Schedule.PSObject.Properties['Times']) { @($job.Schedule.Times) } else { @() }
            if (@($times).Count -eq 0) { $errors.Add("Job '$name': Schedule.Times is empty.") }
            foreach ($timeText in $times) {
                $parsed = [timespan]::Zero
                if (-not [timespan]::TryParseExact([string]$timeText, 'hh\:mm', [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) { $errors.Add("Job '$name': invalid time '$timeText'.") }
            }
            if ($type -eq 'Weekly') {
                $days = if ($job.Schedule.PSObject.Properties['DaysOfWeek']) { @($job.Schedule.DaysOfWeek) } else { @() }
                if (@($days).Count -eq 0) { $errors.Add("Job '$name': Weekly requires DaysOfWeek.") }
                foreach ($day in $days) { if ([string]$day -notin $script:ValidDays) { $errors.Add("Job '$name': invalid day '$day'.") } }
            }
            $missed = if ($job.Schedule.PSObject.Properties['MissedRunPolicy'] -and $job.Schedule.MissedRunPolicy) { [string]$job.Schedule.MissedRunPolicy } else { 'RunOnce' }
            if ($missed -notin @('RunOnce', 'Skip')) { $errors.Add("Job '$name': invalid MissedRunPolicy.") }
        }
        $dependencies[$name] = if ($job.PSObject.Properties['DependsOn']) { @($job.DependsOn | ForEach-Object { [string]$_ }) } else { @() }
    }
    foreach ($name in @($dependencies.Keys)) { foreach ($dependency in @($dependencies[$name])) { if (-not $seen.ContainsKey($dependency)) { $errors.Add("Job '$name': unknown dependency '$dependency'.") } } }
    if ($errors.Count -eq 0) {
        $visiting = @{}; $visited = @{}
        function Test-DependencyNode {
            param([string]$Name)
            if ($visiting.ContainsKey($Name)) { return $false }
            if ($visited.ContainsKey($Name)) { return $true }
            $visiting[$Name] = $true
            foreach ($dependency in @($dependencies[$Name])) { if (-not (Test-DependencyNode $dependency)) { return $false } }
            $visiting.Remove($Name); $visited[$Name] = $true
            return $true
        }
        foreach ($name in @($dependencies.Keys)) { if (-not (Test-DependencyNode $name)) { $errors.Add('Dependency cycle detected in DependsOn definitions.'); break } }
    }
    [pscustomobject]@{ Valid = ($errors.Count -eq 0); Errors = @($errors); Warnings = @($warnings); JobCount = @($Document.Jobs).Count }
}

function Test-SmartM365OrchestratorClusterDocument {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Document)
    $errors = [Collections.Generic.List[string]]::new(); $warnings = [Collections.Generic.List[string]]::new()
    $servers = if ($Document.PSObject.Properties['ExpectedOrchestratorServers']) { @($Document.ExpectedOrchestratorServers | ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } | Where-Object { $_ }) } else { @() }
    if (@($servers).Count -eq 0) { $warnings.Add('ExpectedOrchestratorServers is empty.') }
    if (@($servers | Sort-Object -Unique).Count -ne @($servers).Count) { $errors.Add('ExpectedOrchestratorServers contains duplicates.') }
    if ($Document.PSObject.Properties['ElectionWeightsByServer'] -and $null -ne $Document.ElectionWeightsByServer) {
        foreach ($property in @($Document.ElectionWeightsByServer.PSObject.Properties)) {
            if ([double]$property.Value -le 0) { $errors.Add("Election weight for '$($property.Name)' must be greater than zero.") }
            if (@($servers).Count -gt 0 -and $property.Name.ToUpperInvariant() -notin $servers) { $warnings.Add("Weight exists for non-expected server '$($property.Name)'.") }
        }
    }
    if ($Document.PSObject.Properties['ServerJobPolicies'] -and $null -ne $Document.ServerJobPolicies) {
        foreach ($property in @($Document.ServerJobPolicies.PSObject.Properties)) {
            if ($null -eq $property.Value -or -not $property.Value.PSObject.Properties['OnlyJobsRequiring']) { $errors.Add("ServerJobPolicies.$($property.Name) must define OnlyJobsRequiring."); continue }
            $only = @($property.Value.OnlyJobsRequiring)
            if ($only.Count -eq 0) { $errors.Add("ServerJobPolicies.$($property.Name).OnlyJobsRequiring is empty.") }
            foreach ($capability in $only) { if ([string]$capability -notin $script:ValidCapabilities) { $errors.Add("Unknown policy capability '$capability'.") } }
        }
    }
    foreach ($name in @('PeerMonitoringCheckIntervalSeconds', 'PeerHeartbeatStaleMinutes', 'PeerMonitoringConfirmationChecks', 'PeerJobStartGraceMinutes', 'PeerAlertReminderMinutes', 'PeerAlertMailRetryMinutes')) {
        if ($Document.PSObject.Properties[$name] -and [int]$Document.$name -lt 1) { $errors.Add("$name must be greater than zero.") }
    }
    [pscustomobject]@{ Valid = ($errors.Count -eq 0); Errors = @($errors); Warnings = @($warnings); ServerCount = @($servers).Count }
}

function Test-SmartM365OrchestratorConfigurationConsistency {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$JobsDocument, [Parameter(Mandatory)]$ClusterDocument)
    $errors = [Collections.Generic.List[string]]::new(); $warnings = [Collections.Generic.List[string]]::new()
    $servers = @($ClusterDocument.ExpectedOrchestratorServers | ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } | Where-Object { $_ })
    foreach ($job in @($JobsDocument.Jobs)) {
        if ([string]$job.AssignmentMode -ne 'Pinned') { continue }
        $pinnedServer = @($job.AllowedServers | ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } | Where-Object { $_ })
        if ($pinnedServer.Count -eq 1 -and $pinnedServer[0] -notin $servers) { $errors.Add("Job '$($job.Name)': pinned server '$($pinnedServer[0])' is not in ExpectedOrchestratorServers.") }
    }
    if ($ClusterDocument.PSObject.Properties['ServerJobPolicies']) {
        foreach ($property in @($ClusterDocument.ServerJobPolicies.PSObject.Properties)) { if ($property.Name.ToUpperInvariant() -notin $servers) { $warnings.Add("Server policy exists for non-expected server '$($property.Name)'.") } }
    }
    [pscustomobject]@{ Valid = ($errors.Count -eq 0); Errors = @($errors); Warnings = @($warnings) }
}
function Get-SmartM365OrchestratorConfigurationSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)
    $paths = Get-SmartM365OrchestratorConfigurationPaths $SharedDataFolderPath
    [pscustomobject]@{ Paths = $paths; Jobs = Read-SmartM365OrchestratorJson $paths.JobsPath; Cluster = Read-SmartM365OrchestratorJson $paths.ClusterPath; JobsHash = Get-SmartM365OrchestratorFileHash $paths.JobsPath; ClusterHash = Get-SmartM365OrchestratorFileHash $paths.ClusterPath; ReadUtc = [datetime]::UtcNow }
}

function Initialize-SmartM365OrchestratorCentralConfiguration {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)][string]$BootstrapJobsPath, [Parameter(Mandatory)]$BootstrapClusterDocument)
    $paths = Get-SmartM365OrchestratorConfigurationPaths $SharedDataFolderPath
    foreach ($folder in @($paths.ConfigFolderPath, $paths.VersionsFolderPath, $paths.AuditFolderPath)) { if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null } }
    if (-not (Test-Path -LiteralPath $paths.JobsPath)) {
        try { Write-SmartM365OrchestratorJsonAtomically $paths.JobsPath (Read-SmartM365OrchestratorJson $BootstrapJobsPath) -CreateNew } catch [IO.IOException] { if (-not (Test-Path -LiteralPath $paths.JobsPath)) { throw } }
    }
    if (-not (Test-Path -LiteralPath $paths.ClusterPath)) {
        try { Write-SmartM365OrchestratorJsonAtomically $paths.ClusterPath $BootstrapClusterDocument -CreateNew } catch [IO.IOException] { if (-not (Test-Path -LiteralPath $paths.ClusterPath)) { throw } }
    }
    $snapshot = Get-SmartM365OrchestratorConfigurationSnapshot $SharedDataFolderPath
    $jobsValidation = Test-SmartM365OrchestratorJobsDocument $snapshot.Jobs; $clusterValidation = Test-SmartM365OrchestratorClusterDocument $snapshot.Cluster
    if (-not $jobsValidation.Valid) { throw "Central jobs configuration is invalid: $($jobsValidation.Errors -join '; ')" }
    if (-not $clusterValidation.Valid) { throw "Central cluster configuration is invalid: $($clusterValidation.Errors -join '; ')" }
    $snapshot
}

function Enter-ConfigurationLock {
    param([string]$Path, [int]$TimeoutSeconds)
    $deadline = [datetime]::UtcNow.AddSeconds([math]::Max(1, $TimeoutSeconds))
    do { try { return [IO.File]::Open($Path, 'CreateNew', 'ReadWrite', 'None') } catch [IO.IOException] { Start-Sleep -Milliseconds 200 } } while ([datetime]::UtcNow -lt $deadline)
    throw "Configuration is locked by another editor: $Path"
}

function Publish-SmartM365OrchestratorConfiguration {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)]$JobsDocument, [Parameter(Mandatory)]$ClusterDocument, [string]$ExpectedJobsHash = '', [string]$ExpectedClusterHash = '', [string]$ChangeSummary = 'Configuration updated', [int]$LockTimeoutSeconds = 15)
    $jobsValidation = Test-SmartM365OrchestratorJobsDocument $JobsDocument; $clusterValidation = Test-SmartM365OrchestratorClusterDocument $ClusterDocument
    $consistencyValidation = Test-SmartM365OrchestratorConfigurationConsistency $JobsDocument $ClusterDocument
    $errors = @($jobsValidation.Errors) + @($clusterValidation.Errors) + @($consistencyValidation.Errors)
    if (@($errors).Count -gt 0) { throw "Configuration validation failed: $($errors -join '; ')" }
    $paths = Get-SmartM365OrchestratorConfigurationPaths $SharedDataFolderPath; $lock = Enter-ConfigurationLock $paths.LockPath $LockTimeoutSeconds
    try {
        $currentJobsHash = Get-SmartM365OrchestratorFileHash $paths.JobsPath; $currentClusterHash = Get-SmartM365OrchestratorFileHash $paths.ClusterPath
        if ($ExpectedJobsHash -and $ExpectedJobsHash -ne $currentJobsHash) { throw 'The jobs configuration changed after it was loaded. Refresh before publishing.' }
        if ($ExpectedClusterHash -and $ExpectedClusterHash -ne $currentClusterHash) { throw 'The cluster configuration changed after it was loaded. Refresh before publishing.' }
        $versionId = '{0}_{1}' -f [datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ'), [guid]::NewGuid().ToString('N').Substring(0, 8)
        $versionFolder = Join-Path $paths.VersionsFolderPath $versionId; New-Item -ItemType Directory -Path $versionFolder -Force | Out-Null
        Copy-Item $paths.JobsPath (Join-Path $versionFolder 'Orchestrator-Jobs.before.json'); Copy-Item $paths.ClusterPath (Join-Path $versionFolder 'Orchestrator-Cluster.before.json')
        Write-SmartM365OrchestratorJsonAtomically $paths.JobsPath $JobsDocument; Write-SmartM365OrchestratorJsonAtomically $paths.ClusterPath $ClusterDocument
        Copy-Item $paths.JobsPath (Join-Path $versionFolder 'Orchestrator-Jobs.after.json'); Copy-Item $paths.ClusterPath (Join-Path $versionFolder 'Orchestrator-Cluster.after.json')
        $newJobsHash = Get-SmartM365OrchestratorFileHash $paths.JobsPath; $newClusterHash = Get-SmartM365OrchestratorFileHash $paths.ClusterPath
        $audit = [pscustomobject][ordered]@{ ChangedUtc = [datetime]::UtcNow.ToString('o'); ChangedBy = [Security.Principal.WindowsIdentity]::GetCurrent().Name; ChangedFromServer = $env:COMPUTERNAME; VersionId = $versionId; Summary = $ChangeSummary; PreviousJobsHash = $currentJobsHash; NewJobsHash = $newJobsHash; PreviousClusterHash = $currentClusterHash; NewClusterHash = $newClusterHash }
        if (Test-Path $paths.AuditPath) { $audit | Export-Csv $paths.AuditPath -NoTypeInformation -Append -Encoding utf8 } else { $audit | Export-Csv $paths.AuditPath -NoTypeInformation -Encoding utf8 }
        [pscustomobject]@{ VersionId = $versionId; VersionFolderPath = $versionFolder; JobsHash = $newJobsHash; ClusterHash = $newClusterHash; Warnings = @($jobsValidation.Warnings) + @($clusterValidation.Warnings) + @($consistencyValidation.Warnings) }
    }
    finally { if ($null -ne $lock) { $lock.Dispose() }; Remove-Item $paths.LockPath -Force -ErrorAction SilentlyContinue }
}

function Get-SmartM365OrchestratorConfigurationVersions {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)
    $paths = Get-SmartM365OrchestratorConfigurationPaths $SharedDataFolderPath
    if (-not (Test-Path $paths.VersionsFolderPath)) { return @() }
    @(Get-ChildItem $paths.VersionsFolderPath -Directory | Sort-Object Name -Descending | ForEach-Object { [pscustomobject]@{ VersionId = $_.Name; Created = $_.CreationTime; FolderPath = $_.FullName } })
}

function Restore-SmartM365OrchestratorConfigurationVersion {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)][string]$VersionFolderPath, [ValidateSet('Before', 'After')][string]$Snapshot = 'Before', [string]$ExpectedJobsHash = '', [string]$ExpectedClusterHash = '')
    $suffix = $Snapshot.ToLowerInvariant()
    $jobs = Read-SmartM365OrchestratorJson (Join-Path $VersionFolderPath "Orchestrator-Jobs.$suffix.json"); $cluster = Read-SmartM365OrchestratorJson (Join-Path $VersionFolderPath "Orchestrator-Cluster.$suffix.json")
    Publish-SmartM365OrchestratorConfiguration $SharedDataFolderPath $jobs $cluster $ExpectedJobsHash $ExpectedClusterHash "Rollback to $(Split-Path $VersionFolderPath -Leaf) ($Snapshot)"
}

function Get-SmartM365OrchestratorHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [datetime]$From = (Get-Date).AddDays(-7), [datetime]$To = (Get-Date), [string]$Server = '', [string]$JobName = '', [string]$Status = '')
    $result = [Collections.Generic.List[object]]::new()
    if (-not (Test-Path $SharedDataFolderPath)) { return @() }
    foreach ($serverFolder in @(Get-ChildItem $SharedDataFolderPath -Directory -ErrorAction SilentlyContinue)) {
        if ($Server -and $serverFolder.Name -ine $Server) { continue }
        $jobRunsFolder = Join-Path $serverFolder.FullName 'JobRuns'; if (-not (Test-Path $jobRunsFolder)) { continue }
        foreach ($csv in @(Get-ChildItem $jobRunsFolder -Filter 'Orchestrator_JobRuns_*.csv' -File -ErrorAction SilentlyContinue)) {
            foreach ($row in @(Import-Csv $csv.FullName -ErrorAction SilentlyContinue)) {
                $start = [datetime]::MinValue; if (-not [datetime]::TryParse([string]$row.StartTime, [ref]$start)) { continue }
                if ($start -lt $From -or $start -gt $To) { continue }; if ($JobName -and $row.JobName -ine $JobName) { continue }; if ($Status -and $row.Status -ine $Status) { continue }
                $result.Add([pscustomobject]@{ Server = $serverFolder.Name; JobName = [string]$row.JobName; ScheduledTime = [string]$row.ScheduledTime; StartTime = $start; EndTime = [string]$row.EndTime; DurationSec = [double]$row.DurationSec; ExitCode = [string]$row.ExitCode; Status = [string]$row.Status; RetryCount = [int]$row.RetryCount; LogPath = [string]$row.LogPath })
            }
        }
    }
    @($result | Sort-Object StartTime -Descending)
}

function Get-SmartM365OrchestratorServerStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)]$ClusterDocument)
    $assignments = @{}; $planPath = Join-Path $SharedDataFolderPath 'Election\Orchestrator-ElectionPlan.json'
    if (Test-Path $planPath) { try { $plan = Read-SmartM365OrchestratorJson $planPath; foreach ($assignment in @($plan.Assignments)) { $assignments[[string]$assignment.JobName] = [string]$assignment.OwnerServer } } catch {} }
    $weights = ConvertTo-SmartM365OrchestratorHashtable $ClusterDocument.ElectionWeightsByServer; $policies = ConvertTo-SmartM365OrchestratorHashtable $ClusterDocument.ServerJobPolicies
    $result = foreach ($server in @($ClusterDocument.ExpectedOrchestratorServers | Sort-Object -Unique)) {
        $serverName = ([string]$server).ToUpperInvariant(); $serverFolder = Join-Path $SharedDataFolderPath $server
        $heartbeatPath = Join-Path $serverFolder 'Orchestrator-Heartbeat.json'; $capabilitiesPath = Join-Path $serverFolder 'Orchestrator-Capabilities.json'; $heartbeatTime = [datetime]::MinValue
        if (Test-Path $heartbeatPath) { $heartbeat = Read-SmartM365OrchestratorJson $heartbeatPath; foreach ($name in @('TimestampUtc', 'HeartbeatUtc', 'UpdatedUtc')) { if ($heartbeat.PSObject.Properties[$name] -and [datetime]::TryParse([string]$heartbeat.$name, [ref]$heartbeatTime)) { break } } }
        $age = if ($heartbeatTime -eq [datetime]::MinValue) { [double]::PositiveInfinity } else { ([datetime]::UtcNow - $heartbeatTime.ToUniversalTime()).TotalMinutes }
        $capabilityText = ''; if (Test-Path $capabilitiesPath) { $capabilities = Read-SmartM365OrchestratorJson $capabilitiesPath; if ($capabilities.PSObject.Properties['ReadyCapabilities']) { $capabilityText = @($capabilities.ReadyCapabilities | Sort-Object -Unique) -join ', ' } }
        $assigned = @($assignments.Keys | Where-Object { $assignments[$_] -ieq $server })
        [pscustomobject]@{ Server = [string]$server; Online = ($age -le [double]$ClusterDocument.PeerHeartbeatStaleMinutes); HeartbeatUtc = if ($heartbeatTime -eq [datetime]::MinValue) { $null } else { $heartbeatTime.ToUniversalTime() }; HeartbeatAgeMinutes = if ([double]::IsPositiveInfinity($age)) { $null } else { [math]::Round($age, 1) }; Capabilities = $capabilityText; Weight = if ($weights.ContainsKey($serverName)) { [double]$weights[$serverName] } else { 1.0 }; Policy = if ($policies.ContainsKey($serverName)) { @($policies[$serverName].OnlyJobsRequiring) -join ', ' } else { '' }; AssignedJobs = $assigned.Count; AssignedJobNames = $assigned -join ', ' }
    }
    @($result)
}

Export-ModuleMember -Function @(
    'ConvertTo-SmartM365OrchestratorHashtable', 'Get-SmartM365OrchestratorConfigurationPaths', 'Get-SmartM365OrchestratorFileHash',
    'Read-SmartM365OrchestratorJson', 'Write-SmartM365OrchestratorJsonAtomically', 'Test-SmartM365OrchestratorJobsDocument',
    'Test-SmartM365OrchestratorClusterDocument', 'Test-SmartM365OrchestratorConfigurationConsistency', 'Get-SmartM365OrchestratorConfigurationSnapshot', 'Initialize-SmartM365OrchestratorCentralConfiguration',
    'Publish-SmartM365OrchestratorConfiguration', 'Get-SmartM365OrchestratorConfigurationVersions', 'Restore-SmartM365OrchestratorConfigurationVersion',
    'Get-SmartM365OrchestratorHistory', 'Get-SmartM365OrchestratorServerStatus'
)
