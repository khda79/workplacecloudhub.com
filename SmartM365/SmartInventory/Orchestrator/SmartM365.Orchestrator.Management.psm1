Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop
$script:JsonVersionRoots = @{}

function Resolve-OrchestratorConfigurationJsonPath {
    param([Parameter(Mandatory)][string]$Path)
    Resolve-SmartM365OwnedJsonPath -Path $Path -Owner 'Orchestrator configuration' -Validate {
        param($document) if ($document -isnot [pscustomobject]) { throw 'Orchestrator configuration must be an object.' }
    }
}

function Convert-OrchestratorConfigurationVersions {
    param([Parameter(Mandatory)][string]$Root)
    if ((Get-SmartM365JsonTransportPolicy).Mode -ne 'JsonText' -or $script:JsonVersionRoots.ContainsKey($Root) -or -not (Test-Path -LiteralPath $Root)) { return }
    $legacyLeaves = @('Orchestrator-Jobs.before.json','Orchestrator-Cluster.before.json','Orchestrator-Jobs.after.json','Orchestrator-Cluster.after.json','Publication-Failed.json')
    foreach ($folder in @(Get-ChildItem -LiteralPath $Root -Directory -ErrorAction Stop | Where-Object { $_.Name -match '^\d{8}T\d{9}Z_[0-9a-f]{8}$' })) {
        foreach ($file in @(Get-ChildItem -LiteralPath $folder.FullName -File -ErrorAction Stop | Where-Object { $_.Name -in $legacyLeaves })) {
            $null = Resolve-OrchestratorConfigurationJsonPath $file.FullName
        }
    }
    $script:JsonVersionRoots[$Root] = $true
}

function Copy-OrchestratorConfigurationSnapshot {
    param([Parameter(Mandatory)][string]$Source,[Parameter(Mandatory)][string]$Destination)
    $sourceDocument = Read-SmartM365JsonDocument $Source
    $bytes = [IO.File]::ReadAllBytes($sourceDocument.Path)
    $receipt = Write-SmartM365JsonBytesAtomically -Path $Destination -Bytes $bytes -ExpectedSHA256 'ABSENT' -Validate { param($document) if ($document -isnot [pscustomobject]) { throw 'Configuration snapshot must be an object.' } }
    if ($receipt.SHA256 -ne $sourceDocument.SHA256) { throw 'Configuration source changed while saving version snapshot.' }
}

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
    Convert-OrchestratorConfigurationVersions -Root (Join-Path $configFolder 'Versions')
    [pscustomobject]@{
        SharedDataFolderPath = $SharedDataFolderPath
        ConfigFolderPath = $configFolder
        JobsPath = Resolve-OrchestratorConfigurationJsonPath (Join-Path $configFolder 'Orchestrator-Jobs.json')
        ClusterPath = Resolve-OrchestratorConfigurationJsonPath (Join-Path $configFolder 'Orchestrator-Cluster.json')
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
    if ($Path.EndsWith('.template', [StringComparison]::OrdinalIgnoreCase)) {
        return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -Depth 100 -ErrorAction Stop)
    }
    (Read-SmartM365JsonDocument -Path $Path).Document
}

function Get-SmartM365OrchestratorJsonContent {
    param([Parameter(Mandatory)]$Document)
    return ($Document | ConvertTo-Json -Depth 100) + [Environment]::NewLine
}

function Get-SmartM365OrchestratorContentHash {
    param([Parameter(Mandatory)][string]$Content)
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($Content)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToHexString($sha256.ComputeHash($bytes))
    }
    finally {
        $sha256.Dispose()
    }
}

function Move-SmartM365OrchestratorFileWithRetry {
    param(
        [Parameter(Mandatory)][string]$TemporaryPath,
        [Parameter(Mandatory)][string]$Path,
        [ValidateRange(0, 300)][int]$RetrySeconds = 10
    )

    $deadline = [datetime]::UtcNow.AddSeconds($RetrySeconds)
    $attempt = 0
    while ($true) {
        try {
            # Move-Item -Force follows the Windows/SMB replacement path already
            # used by the resident orchestrator for its shared JSON state.
            Move-Item -LiteralPath $TemporaryPath -Destination $Path -Force -ErrorAction Stop
            return
        }
        catch {
            if ([datetime]::UtcNow -ge $deadline) {
                $attributes = if (Test-Path -LiteralPath $Path -PathType Leaf) {
                    [string](Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue).Attributes
                }
                else {
                    'Missing'
                }
                $exceptionType = $_.Exception.GetType().FullName
                $hresult = '0x{0:X8}' -f ($_.Exception.HResult -band 0xffffffffL)
                $message = "Atomic replacement failed for '$Path' after $($attempt + 1) attempt(s). TargetAttributes=$attributes; ErrorType=$exceptionType; HResult=$hresult; Message=$($_.Exception.Message)"
                throw [IO.IOException]::new($message, $_.Exception)
            }
            $baseDelay = [math]::Min(2000, 100 * [math]::Pow(2, [math]::Min($attempt, 5)))
            $jitter = Get-Random -Minimum 0 -Maximum 251
            Start-Sleep -Milliseconds ([int]($baseDelay + $jitter))
            $attempt++
        }
    }
}

function Write-SmartM365OrchestratorTextAtomically {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Content,
        [ValidateRange(0, 300)][int]$RetrySeconds = 10
    )

    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    if ($Path -match '(?i)\.json(?:\.txt)?$') {
        $null = Write-SmartM365JsonBytesAtomically -Path $Path -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($Content)) -LockTimeoutSeconds $RetrySeconds -Validate { param($document) if ($document -isnot [pscustomobject]) { throw 'Orchestrator JSON must be an object.' } }
        return
    }
    $temporaryPath = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [IO.File]::WriteAllText($temporaryPath, $Content, [Text.UTF8Encoding]::new($false))
        Move-SmartM365OrchestratorFileWithRetry -TemporaryPath $temporaryPath -Path $Path -RetrySeconds $RetrySeconds
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
}

function Write-SmartM365OrchestratorJsonAtomically {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Document,
        [switch]$CreateNew,
        [ValidateRange(0, 300)][int]$RetrySeconds = 10
    )
    $folder = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $content = Get-SmartM365OrchestratorJsonContent -Document $Document
    if ($CreateNew) {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($content)
        $null = Write-SmartM365JsonBytesAtomically -Path $Path -Bytes $bytes -ExpectedSHA256 'ABSENT' -LockTimeoutSeconds $RetrySeconds -Validate { param($document) if ($document -isnot [pscustomobject]) { throw 'Orchestrator JSON must be an object.' } }
        return
    }
    Write-SmartM365OrchestratorTextAtomically -Path $Path -Content $content -RetrySeconds $RetrySeconds
}

function Write-SmartM365OrchestratorManagementLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )

    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null
    }
    $physicalLines = @($Message -split '\r?\n')
    foreach ($physicalLine in $physicalLines) {
        $line = '[{0}][{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'), $Level, $physicalLine
        Add-Content -LiteralPath $Path -Value $line -Encoding utf8 -ErrorAction Stop
    }
}

function Sync-SmartM365OrchestratorJobsManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$TemplatePath,
        [int]$LockTimeoutSeconds = 15
    )

    $template = Read-SmartM365OrchestratorJson -Path $TemplatePath
    $templateValidation = Test-SmartM365OrchestratorJobsDocument -Document $template
    if (-not $templateValidation.Valid) {
        throw "Jobs template is invalid: $($templateValidation.Errors -join '; ')"
    }

    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    $lockPath = Join-Path -Path $folder -ChildPath 'Orchestrator-Configuration.lock'
    $lock = Enter-ConfigurationLock -Path $lockPath -TimeoutSeconds $LockTimeoutSeconds
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            Write-SmartM365OrchestratorJsonAtomically -Path $Path -Document $template -CreateNew
            return [pscustomobject]@{
                Created = $true
                Updated = $true
                AddedJobNames = @($template.Jobs | ForEach-Object { [string]$_.Name })
                UpdatedJobNames = @()
            }
        }

        $document = Read-SmartM365OrchestratorJson -Path $Path
        $documentValidation = Test-SmartM365OrchestratorJobsDocument -Document $document
        if (-not $documentValidation.Valid) {
            throw "Existing jobs manifest is invalid and was not changed: $($documentValidation.Errors -join '; ')"
        }

        $knownNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($job in @($document.Jobs)) { [void]$knownNames.Add([string]$job.Name) }
        $missingJobs = @($template.Jobs | Where-Object { -not $knownNames.Contains([string]$_.Name) })

        $updatedJobNames = [Collections.Generic.List[string]]::new()
        $requiredExternalActionArgument = '-EnableConfiguredExternalActions'
        $templateJobsByName = @{}
        foreach ($templateJob in @($template.Jobs)) { $templateJobsByName[[string]$templateJob.Name] = $templateJob }
        foreach ($existingJob in @($document.Jobs)) {
            $jobName = [string]$existingJob.Name
            if (-not $templateJobsByName.ContainsKey($jobName)) { continue }
            $templateArguments = [string]$templateJobsByName[$jobName].Arguments
            $existingArguments = [string]$existingJob.Arguments
            $templateRequiresExternalActions = $templateArguments -match '(^|\s)-EnableConfiguredExternalActions(?=\s|$)'
            $existingEnablesExternalActions = $existingArguments -match '(^|\s)-EnableConfiguredExternalActions(?=\s|$)'
            if ($templateRequiresExternalActions -and -not $existingEnablesExternalActions) {
                $existingJob.Arguments = (@($existingArguments.Trim(), $requiredExternalActionArgument) | Where-Object { $_ }) -join ' '
                $updatedJobNames.Add($jobName) | Out-Null
            }
            # Add dependency-policy keys introduced by the template; never overwrite an existing value.
            foreach ($policyKey in @('DependencyMode', 'DependencyMaxAgeHours')) {
                if ($templateJobsByName[$jobName].PSObject.Properties[$policyKey] -and -not $existingJob.PSObject.Properties[$policyKey]) {
                    $existingJob | Add-Member -NotePropertyName $policyKey -NotePropertyValue $templateJobsByName[$jobName].$policyKey
                    if (-not $updatedJobNames.Contains($jobName)) { $updatedJobNames.Add($jobName) | Out-Null }
                }
            }
        }

        if ($missingJobs.Count -eq 0 -and $updatedJobNames.Count -eq 0) {
            return [pscustomobject]@{ Created = $false; Updated = $false; AddedJobNames = @(); UpdatedJobNames = @() }
        }

        $clonedMissingJobs = @($missingJobs | ForEach-Object {
            $_ | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
        })
        $document.Jobs = @($document.Jobs) + $clonedMissingJobs
        $mergedValidation = Test-SmartM365OrchestratorJobsDocument -Document $document
        if (-not $mergedValidation.Valid) {
            throw "Merged jobs manifest is invalid and was not written: $($mergedValidation.Errors -join '; ')"
        }
        Write-SmartM365OrchestratorJsonAtomically -Path $Path -Document $document
        [pscustomobject]@{
            Created = $false
            Updated = $true
            AddedJobNames = @($clonedMissingJobs | ForEach-Object { [string]$_.Name })
            UpdatedJobNames = @($updatedJobNames)
        }
    }
    finally {
        if ($null -ne $lock) { $lock.Dispose() }
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
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
        if ($job.PSObject.Properties['ConcurrencyKey'] -and
            -not [string]::IsNullOrWhiteSpace([string]$job.ConcurrencyKey) -and
            [string]$job.ConcurrencyKey -notmatch '^[A-Za-z0-9._-]+$') {
            $errors.Add("Job '$name': invalid ConcurrencyKey.")
        }
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
        if ($job.PSObject.Properties['DependencyMode'] -and $job.DependencyMode -and [string]$job.DependencyMode -notin @('LatestOccurrence', 'FreshSuccess')) { $errors.Add("Job '$name': DependencyMode must be LatestOccurrence or FreshSuccess.") }
        foreach ($propertyName in @('TimeoutMinutes', 'MaxRetries', 'RetryDelaySeconds', 'MinimumSuccessDurationSeconds', 'DependencyWaitTimeoutMinutes', 'DependencyMaxAgeHours')) {
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
    foreach ($name in @('PeerMonitoringCheckIntervalSeconds', 'PeerHeartbeatStaleMinutes', 'PeerMonitoringConfirmationChecks', 'PeerJobStartGraceMinutes', 'PeerRecycleGraceMinutes', 'PeerAlertReminderMinutes', 'PeerAlertMailRetryMinutes')) {
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
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][string]$BootstrapJobsPath,
        [Parameter(Mandatory)]$BootstrapClusterDocument,
        [string]$TemplateJobsPath = ''
    )
    $paths = Get-SmartM365OrchestratorConfigurationPaths $SharedDataFolderPath
    foreach ($folder in @($paths.ConfigFolderPath, $paths.VersionsFolderPath, $paths.AuditFolderPath)) { if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null } }
    $manifestSync = Sync-SmartM365OrchestratorJobsManifest -Path $paths.JobsPath -TemplatePath $BootstrapJobsPath
    if (-not [string]::IsNullOrWhiteSpace($TemplateJobsPath) -and
        (Resolve-Path -LiteralPath $TemplateJobsPath).Path -ne (Resolve-Path -LiteralPath $BootstrapJobsPath).Path) {
        $templateSync = Sync-SmartM365OrchestratorJobsManifest -Path $paths.JobsPath -TemplatePath $TemplateJobsPath
        $manifestSync = [pscustomobject]@{
            Created = [bool]$manifestSync.Created
            Updated = ([bool]$manifestSync.Updated -or [bool]$templateSync.Updated)
            AddedJobNames = @($manifestSync.AddedJobNames) + @($templateSync.AddedJobNames)
            UpdatedJobNames = @($manifestSync.UpdatedJobNames) + @($templateSync.UpdatedJobNames)
        }
    }
    if (-not (Test-Path -LiteralPath $paths.ClusterPath)) {
        try { Write-SmartM365OrchestratorJsonAtomically $paths.ClusterPath $BootstrapClusterDocument -CreateNew } catch [IO.IOException] { if (-not (Test-Path -LiteralPath $paths.ClusterPath)) { throw } }
    }
    $snapshot = Get-SmartM365OrchestratorConfigurationSnapshot $SharedDataFolderPath
    $jobsValidation = Test-SmartM365OrchestratorJobsDocument $snapshot.Jobs; $clusterValidation = Test-SmartM365OrchestratorClusterDocument $snapshot.Cluster
    if (-not $jobsValidation.Valid) { throw "Central jobs configuration is invalid: $($jobsValidation.Errors -join '; ')" }
    if (-not $clusterValidation.Valid) { throw "Central cluster configuration is invalid: $($clusterValidation.Errors -join '; ')" }
    $snapshot | Add-Member -NotePropertyName ManifestSync -NotePropertyValue $manifestSync
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
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)]$JobsDocument,
        [Parameter(Mandatory)]$ClusterDocument,
        [string]$ExpectedJobsHash = '',
        [string]$ExpectedClusterHash = '',
        [string]$ChangeSummary = 'Configuration updated',
        [int]$LockTimeoutSeconds = 15,
        [ValidateRange(0, 300)][int]$AtomicWriteRetrySeconds = 10
    )
    $jobsValidation = Test-SmartM365OrchestratorJobsDocument $JobsDocument; $clusterValidation = Test-SmartM365OrchestratorClusterDocument $ClusterDocument
    $consistencyValidation = Test-SmartM365OrchestratorConfigurationConsistency $JobsDocument $ClusterDocument
    $errors = @($jobsValidation.Errors) + @($clusterValidation.Errors) + @($consistencyValidation.Errors)
    if (@($errors).Count -gt 0) { throw "Configuration validation failed: $($errors -join '; ')" }
    $paths = Get-SmartM365OrchestratorConfigurationPaths $SharedDataFolderPath; $lock = Enter-ConfigurationLock $paths.LockPath $LockTimeoutSeconds
    try {
        $currentJobsHash = Get-SmartM365OrchestratorFileHash $paths.JobsPath; $currentClusterHash = Get-SmartM365OrchestratorFileHash $paths.ClusterPath
        if ($ExpectedJobsHash -and $ExpectedJobsHash -ne $currentJobsHash) { throw 'The jobs configuration changed after it was loaded. Refresh before publishing.' }
        if ($ExpectedClusterHash -and $ExpectedClusterHash -ne $currentClusterHash) { throw 'The cluster configuration changed after it was loaded. Refresh before publishing.' }
        $desiredJobsHash = Get-SmartM365OrchestratorContentHash -Content (Get-SmartM365OrchestratorJsonContent -Document $JobsDocument)
        $desiredClusterHash = Get-SmartM365OrchestratorContentHash -Content (Get-SmartM365OrchestratorJsonContent -Document $ClusterDocument)
        $jobsChanged = $desiredJobsHash -ne $currentJobsHash
        $clusterChanged = $desiredClusterHash -ne $currentClusterHash
        $versionId = '{0}_{1}' -f [datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ'), [guid]::NewGuid().ToString('N').Substring(0, 8)
        $versionFolder = Join-Path $paths.VersionsFolderPath $versionId; New-Item -ItemType Directory -Path $versionFolder -Force | Out-Null
        $jobsBeforePath = Resolve-OrchestratorConfigurationJsonPath (Join-Path $versionFolder 'Orchestrator-Jobs.before.json')
        $clusterBeforePath = Resolve-OrchestratorConfigurationJsonPath (Join-Path $versionFolder 'Orchestrator-Cluster.before.json')
        Copy-OrchestratorConfigurationSnapshot $paths.JobsPath $jobsBeforePath
        Copy-OrchestratorConfigurationSnapshot $paths.ClusterPath $clusterBeforePath
        $completedWrites = [Collections.Generic.List[string]]::new()
        $failedStage = ''
        try {
            if ($jobsChanged) {
                $failedStage = 'Orchestrator-Jobs.json'
                Write-SmartM365OrchestratorJsonAtomically -Path $paths.JobsPath -Document $JobsDocument -RetrySeconds $AtomicWriteRetrySeconds
                $completedWrites.Add('Jobs') | Out-Null
            }
            if ($clusterChanged) {
                $failedStage = 'Orchestrator-Cluster.json'
                Write-SmartM365OrchestratorJsonAtomically -Path $paths.ClusterPath -Document $ClusterDocument -RetrySeconds $AtomicWriteRetrySeconds
                $completedWrites.Add('Cluster') | Out-Null
            }
        }
        catch {
            $publicationError = $_
            $rollbackErrors = [Collections.Generic.List[string]]::new()
            foreach ($completedWrite in @($completedWrites.ToArray())) {
                try {
                    if ($completedWrite -eq 'Jobs') {
                        Write-SmartM365OrchestratorTextAtomically -Path $paths.JobsPath -Content (Get-Content -LiteralPath $jobsBeforePath -Raw -ErrorAction Stop) -RetrySeconds $AtomicWriteRetrySeconds
                    }
                    elseif ($completedWrite -eq 'Cluster') {
                        Write-SmartM365OrchestratorTextAtomically -Path $paths.ClusterPath -Content (Get-Content -LiteralPath $clusterBeforePath -Raw -ErrorAction Stop) -RetrySeconds $AtomicWriteRetrySeconds
                    }
                }
                catch {
                    $rollbackErrors.Add("$completedWrite rollback failed: $($_.Exception.Message)") | Out-Null
                }
            }

            $rollbackSucceeded = $rollbackErrors.Count -eq 0
            $failureRecord = [pscustomobject][ordered]@{
                FailedUtc = [datetime]::UtcNow.ToString('o')
                FailedBy = [Security.Principal.WindowsIdentity]::GetCurrent().Name
                FailedFromServer = $env:COMPUTERNAME
                VersionId = $versionId
                FailedStage = $failedStage
                CompletedWritesBeforeFailure = @($completedWrites)
                RollbackSucceeded = $rollbackSucceeded
                RollbackErrors = @($rollbackErrors)
                ErrorType = $publicationError.Exception.GetType().FullName
                ErrorMessage = $publicationError.Exception.Message
            }
            $failureRecordWriteError = ''
            try {
                $failurePath = Resolve-OrchestratorConfigurationJsonPath (Join-Path $versionFolder 'Publication-Failed.json')
                Write-SmartM365OrchestratorJsonAtomically -Path $failurePath -Document $failureRecord -CreateNew -RetrySeconds $AtomicWriteRetrySeconds
            }
            catch {
                $failureRecordWriteError = $_.Exception.Message
            }

            $failureRecordNote = if ($failureRecordWriteError) { "; FailureRecordError=$failureRecordWriteError" } else { '' }

            if (-not $rollbackSucceeded) {
                $details = $rollbackErrors -join ' | '
                throw [IO.IOException]::new("Configuration publication failed at '$failedStage' and rollback was incomplete. The shared configuration may be inconsistent. PublicationError=$($publicationError.Exception.Message); RollbackErrors=$details$failureRecordNote", $publicationError.Exception)
            }
            throw [IO.IOException]::new("Configuration publication failed at '$failedStage'. Any earlier file replacement was rolled back. $($publicationError.Exception.Message)$failureRecordNote", $publicationError.Exception)
        }
        Copy-OrchestratorConfigurationSnapshot $paths.JobsPath (Resolve-OrchestratorConfigurationJsonPath (Join-Path $versionFolder 'Orchestrator-Jobs.after.json'))
        Copy-OrchestratorConfigurationSnapshot $paths.ClusterPath (Resolve-OrchestratorConfigurationJsonPath (Join-Path $versionFolder 'Orchestrator-Cluster.after.json'))
        $newJobsHash = Get-SmartM365OrchestratorFileHash $paths.JobsPath; $newClusterHash = Get-SmartM365OrchestratorFileHash $paths.ClusterPath
        $audit = [pscustomobject][ordered]@{ ChangedUtc = [datetime]::UtcNow.ToString('o'); ChangedBy = [Security.Principal.WindowsIdentity]::GetCurrent().Name; ChangedFromServer = $env:COMPUTERNAME; VersionId = $versionId; Summary = $ChangeSummary; PreviousJobsHash = $currentJobsHash; NewJobsHash = $newJobsHash; PreviousClusterHash = $currentClusterHash; NewClusterHash = $newClusterHash }
        if (Test-Path $paths.AuditPath) { $audit | Export-Csv $paths.AuditPath -NoTypeInformation -Append -Encoding utf8 } else { $audit | Export-Csv $paths.AuditPath -NoTypeInformation -Encoding utf8 }
        [pscustomobject]@{ VersionId = $versionId; VersionFolderPath = $versionFolder; JobsHash = $newJobsHash; ClusterHash = $newClusterHash; JobsChanged = $jobsChanged; ClusterChanged = $clusterChanged; Warnings = @($jobsValidation.Warnings) + @($clusterValidation.Warnings) + @($consistencyValidation.Warnings) }
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

function Request-SmartM365OrchestratorRebalance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [string]$Reason = 'Requested from SmartM365 Orchestrator GUI',
        [ValidateRange(0, 300)][int]$AtomicWriteRetrySeconds = 10
    )

    $requestPath = Resolve-OrchestratorConfigurationJsonPath (Join-Path -Path $SharedDataFolderPath -ChildPath 'Election\Orchestrator-RebalanceRequest.json')
    $requestedBy = ''
    try { $requestedBy = [Security.Principal.WindowsIdentity]::GetCurrent().Name }
    catch { $requestedBy = [Environment]::UserName }
    $request = [pscustomobject][ordered]@{
        SchemaVersion = 1
        RequestId = [guid]::NewGuid().ToString('N')
        RequestedAtUtc = [datetime]::UtcNow.ToString('o')
        RequestedBy = $requestedBy
        RequestedFromServer = [Environment]::MachineName
        Reason = $Reason
    }
    Write-SmartM365OrchestratorJsonAtomically -Path $requestPath -Document $request -RetrySeconds $AtomicWriteRetrySeconds
    $request | Add-Member -NotePropertyName RequestPath -NotePropertyValue $requestPath
    return $request
}

function Get-SmartM365OrchestratorHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [datetime]$From = (Get-Date).AddDays(-7), [datetime]$To = (Get-Date), [string]$Server = '', [string]$JobName = '', [string]$Status = '')
    $result = [Collections.Generic.List[object]]::new()
    if (-not (Test-Path $SharedDataFolderPath)) { return @() }
    # The CSV is named when a run is recorded; allow one extra day for clock differences.
    $earliestFileDate = if ($From.Date -gt [datetime]::MinValue.AddDays(1)) { $From.Date.AddDays(-1) } else { [datetime]::MinValue }
    foreach ($serverFolder in @(Get-ChildItem $SharedDataFolderPath -Directory -ErrorAction SilentlyContinue)) {
        if ($Server -and $serverFolder.Name -ine $Server) { continue }
        $jobRunsFolder = Join-Path $serverFolder.FullName 'JobRuns'; if (-not (Test-Path $jobRunsFolder)) { continue }
        foreach ($csv in @(Get-ChildItem $jobRunsFolder -Filter 'Orchestrator_JobRuns_*.csv' -File -ErrorAction SilentlyContinue)) {
            $fileDateMatch = [regex]::Match($csv.BaseName, '^Orchestrator_JobRuns_(\d{8})$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($fileDateMatch.Success) {
                $fileDate = [datetime]::MinValue
                if ([datetime]::TryParseExact($fileDateMatch.Groups[1].Value, 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$fileDate) -and $fileDate -lt $earliestFileDate) { continue }
            }
            foreach ($row in @(Import-Csv $csv.FullName -ErrorAction SilentlyContinue)) {
                $start = [datetime]::MinValue
                $hasStart = [datetime]::TryParse([string]$row.StartTime, [ref]$start)
                $eventTime = $start
                if (-not $hasStart) {
                    # Skipped and blocked occurrences have no process start. Keep their actual
                    # StartTime empty, but include them in the history by scheduled time.
                    if (-not [datetime]::TryParse([string]$row.ScheduledTime, [ref]$eventTime) -and
                        -not [datetime]::TryParse([string]$row.EndTime, [ref]$eventTime)) { continue }
                }
                if ($eventTime -lt $From -or $eventTime -gt $To) { continue }; if ($JobName -and $row.JobName -ine $JobName) { continue }; if ($Status -and $row.Status -ine $Status) { continue }
                $result.Add([pscustomobject]@{ Server = $serverFolder.Name; JobName = [string]$row.JobName; ScheduledTime = [string]$row.ScheduledTime; StartTime = $(if ($hasStart) { $start } else { $null }); EventTime = $eventTime; EndTime = [string]$row.EndTime; DurationSec = [double]$row.DurationSec; ExitCode = [string]$row.ExitCode; Status = [string]$row.Status; RetryCount = [int]$row.RetryCount; LogPath = [string]$row.LogPath })
            }
        }
    }
    @($result | Sort-Object EventTime -Descending)
}

function Get-SmartM365OrchestratorServerStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)]$ClusterDocument)
    $assignments = @{}; $planPath = Join-Path $SharedDataFolderPath 'Election\Orchestrator-ElectionPlan.json'
    $selectedPlan = Get-SmartM365JsonReadPath $planPath -Optional
    if ($selectedPlan) { $planPath = $selectedPlan }
    if (Test-Path $planPath) { try { $plan = Read-SmartM365OrchestratorJson $planPath; foreach ($assignment in @($plan.Assignments)) { $assignments[[string]$assignment.JobName] = [string]$assignment.OwnerServer } } catch {} }
    $weights = ConvertTo-SmartM365OrchestratorHashtable $ClusterDocument.ElectionWeightsByServer; $policies = ConvertTo-SmartM365OrchestratorHashtable $ClusterDocument.ServerJobPolicies
    $result = foreach ($server in @($ClusterDocument.ExpectedOrchestratorServers | Sort-Object -Unique)) {
        $serverName = ([string]$server).ToUpperInvariant(); $serverFolder = Join-Path $SharedDataFolderPath $server
        $heartbeatPath = Join-Path $serverFolder 'Orchestrator-Heartbeat.json'; $capabilitiesPath = Join-Path $serverFolder 'Orchestrator-Capabilities.json'; $heartbeatTime = [datetime]::MinValue
        $selectedHeartbeat = Get-SmartM365JsonReadPath $heartbeatPath -Optional
        $selectedCapabilities = Get-SmartM365JsonReadPath $capabilitiesPath -Optional
        if ($selectedHeartbeat) { $heartbeatPath = $selectedHeartbeat }
        if ($selectedCapabilities) { $capabilitiesPath = $selectedCapabilities }
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
    'Read-SmartM365OrchestratorJson', 'Write-SmartM365OrchestratorJsonAtomically', 'Write-SmartM365OrchestratorManagementLog', 'Sync-SmartM365OrchestratorJobsManifest', 'Test-SmartM365OrchestratorJobsDocument',
    'Test-SmartM365OrchestratorClusterDocument', 'Test-SmartM365OrchestratorConfigurationConsistency', 'Get-SmartM365OrchestratorConfigurationSnapshot', 'Initialize-SmartM365OrchestratorCentralConfiguration',
    'Publish-SmartM365OrchestratorConfiguration', 'Get-SmartM365OrchestratorConfigurationVersions', 'Restore-SmartM365OrchestratorConfigurationVersion',
    'Request-SmartM365OrchestratorRebalance', 'Get-SmartM365OrchestratorHistory', 'Get-SmartM365OrchestratorServerStatus'
)

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBkFLZa/vEWRq4q
# +8sEDjdGMyUII8ezZBJ5M+J5zHE7eaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCfxCAlndKDljHUFJ6i8ZEY
# SkT8VyFcJJeNvLOYeyvxlTANBgkqhkiG9w0BAQEFAASCAYB8tSH3/25olIU8/DxV
# vIR1MtLNEO5NiWjrBWDQVnBZq1SxeipDiH2qvU7kOvMVfGodqOvr1qP3H2Ee/ze7
# ywWqQxh3jpi9P7BNh+opS64bmnKerhFbUj3HhK2P4o0TwrkrnkNbEaTJD43PtHBP
# T+r3HmI2KjgvxmYvN5dPMWc6TTzBn3hTJdVWADaNYJ4jwM9aFP/kEPhYuQ5zEn4s
# wueWtb/yD6R8FJUOHWgkTwFUm2QOrwDg3kty/LZQgTN3OFEHbAMv9NXpOMm4A58Y
# Zr5RSKSAFbcixCu6sKmgH/8H29oCwtt65Nqo1rapXQB6zieXQBodlx0AqRS/5SlQ
# EeNHxndBuIvvuFqSWWP58GVwd8O3hjR5CoGUNXR1997SLl51DqJSJd5d3zEtORA+
# zYQXDQIMTScrnnZNSeUPZ7494mfGfIUJTKsn3SjAZTKlqTwXWK+OIT2QSGG+ILmV
# M6PMWt8MiMAP0vKG2WdhVPsb+uVcEDCDEse7lmyzNzHcpbE=
# SIG # End signature block
