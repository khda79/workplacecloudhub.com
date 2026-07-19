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
    if ($InputObject.GetType() -eq [System.Management.Automation.PSCustomObject]) {
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
        $assignmentMode = if ($job.PSObject.Properties['AssignmentMode'] -and $job.AssignmentMode) { [string]$job.AssignmentMode } else { 'Legacy' }
        if ($assignmentMode -ne 'Pinned') { continue }
        $allowedServers = if ($job.PSObject.Properties['AllowedServers']) { @($job.AllowedServers) } else { @() }
        $pinnedServer = @($allowedServers | ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } | Where-Object { $_ })
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
        if (Test-Path $heartbeatPath) {
            $heartbeat = Read-SmartM365OrchestratorJson $heartbeatPath
            foreach ($name in @('Timestamp', 'TimestampUtc', 'HeartbeatUtc', 'UpdatedUtc')) {
                if (-not $heartbeat.PSObject.Properties[$name]) { continue }
                $value = $heartbeat.PSObject.Properties[$name].Value
                if ($value -is [datetime]) { $heartbeatTime = [datetime]$value; break }
                $parsedTime = [datetime]::MinValue
                if ([datetime]::TryParse([string]$value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsedTime) -or [datetime]::TryParse([string]$value, [ref]$parsedTime)) {
                    $heartbeatTime = $parsedTime
                    break
                }
            }
        }
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCSFTEnLKA5Od7
# zc+1WRBWDA+kvyC6Xj6v5hIwzYm1MqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIIZ+sDCZ0KNiibK5FYHKjIBkiWDR5aZ9PelIe3wHEYPzMA0GCSqG
# SIb3DQEBAQUABIIBgKINRNmn+gJjXJRkiISk3+STnampdjorbrs52B97K21YNK62
# tUybpT4T/G6Shou239C6OhoitqcTycmwgzyUd3ZPeIp5v5/iGS0CoJUEU1qyFjAB
# hfKS6LC7crJQKKywkSbzCs9Kao4UlWMbXo5cAh/soFDVLL14dT1ukg835XiuMXIV
# cJZNYWm1RrG2g/DpoUxOuds88ZFq/90fRsVXC6VfdTlgQY13HhZozembCuojsmNW
# KMyRQYt/gUBZnGA7Y/LF7kZwmH/OvzT4V5vRZUoaoj5c1UTrUdUUf61UKH8yMnjt
# 1NuAB1+fnhARDfEbMiHnwdurYT7OMcoAI8w93D8TP01kyUX1+Y4RvBMgKsVTv0lv
# dOoWI7poJrhYTsQTHQmWc+3cNl7kZPtdIzXqujhrCrpqNZ+20cRvlSSxB3nfNEqc
# p77+HfxUeNZ3beNTra1yywJB3APoELIo9ZWfBmSf6+UQb8dKVyZnPmHIwAjCBO9z
# yr2Fjp6NJW9hbOvplqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMDAw
# MTNaMC8GCSqGSIb3DQEJBDEiBCB6NYH1AChCcJ6x1ZYwogD1DCPM40qmSweVTJoT
# SRyYzTANBgkqhkiG9w0BAQEFAASCAgAafZtVef5VSn2ZHvEe2yfmjXPG+iaM2S3U
# ynhI6s/WLB8x7X4NodOmYM+/+ECmJjVnxW3kFe3EBrJGa2NKndZUCdfz3gzPAEGu
# MfMg9qimocRKJAkdfk9G0y3glwFFRVUJslB0KWyCIpOR3xlkxSSvocLZq/S7h2Ko
# JsYDqWfKZeTDsiD5wNiiKBwCsgAI07YcCIeFw5WzCS1iVBDcZtklUI6CgeA3ayr+
# RhtsiP1Sjwva7mGQ4Nj88Sa/sYpDZ6Tpfj72ukWiERgUHWtFtLc2/NT9RFQFbCRP
# 5F69f1Kj/daBR+K42WafFMiU8uXJSkxE7rQYdL6h3HGiaQwrcVKh/MFpkckVDY9h
# 63BFAAdv1SoN2WOeHTdqvCLOmO23SKTHSitzBxJL7bGgESwXSEx81UGLh2dSvVG8
# cpf3WWLDWD/frxB+ItOFn5RSzACrFqFld1rK3BaucxRTlbQg09r1MLUd1Yvnqpgs
# IXh7A9fEX5KT7FSwfqv0ctTeSdYCfiUDqqU4F1ilKKwecST+yI8gjTyIUYMFTcMU
# M0KsbRfwhSnEu4nNYQFZHZlIB1ulolIaAha3c/Sw/Zo7JD5P4EssbmtWgIseh9ie
# Q3WTyOcgt8OjQ5XCd3b2TD3sQmrPoq+XZpejZVzCLRXBIj/ijsFLPf2LfJTpLJd6
# AVd8ZswtOQ==
# SIG # End signature block
