Set-StrictMode -Version 2.0

$script:PipelineTerminalStatuses = @(
    'Success',
    'CompletedWithWarnings',
    'Failed',
    'TimedOut',
    'Interrupted',
    'BlockedDependencyFailed',
    'BlockedDependencyTimeout',
    'Rejected',
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
        if ($BatchId -notmatch '^[A-Za-z0-9._-]+$') { throw "Invalid pipeline BatchId '$BatchId'." }
        $batchPath = Join-Path -Path $root -ChildPath $BatchId
        $result['BatchPath'] = $batchPath
        $result['RequestPath'] = Join-Path -Path $batchPath -ChildPath 'request.json'
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

    if ($JobName -notmatch '^[A-Za-z0-9._-]+$') { throw "Invalid pipeline job name '$JobName'." }
    $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId
    Join-Path -Path $paths.JobsFolderPath -ChildPath ($JobName + '.json')
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
    if ($CreateNew) {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($content)
        $stream = [IO.File]::Open($Path, 'CreateNew', 'Write', 'Read')
        try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
        return
    }

    $temporaryPath = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [IO.File]::WriteAllText($temporaryPath, $content, [Text.UTF8Encoding]::new($false))
        Move-SmartM365OrchestratorPipelineFileWithRetry -SourcePath $temporaryPath -DestinationPath $Path
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
}

function Read-SmartM365OrchestratorPipelineJson {
    param([Parameter(Mandatory)][string]$Path)
    Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -Depth 100 -ErrorAction Stop
}

function ConvertTo-SmartM365OrchestratorPipelineTimestampText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return ([datetime]$Value).ToUniversalTime().ToString('o') }
    [string]$Value
}

function Get-SmartM365OrchestratorPipelineSelection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$JobsDocument,
        [ValidateSet('Full', 'AD', 'Exchange', 'Exchange2016', 'M365', 'Intune')]
        [string]$Pipeline = 'Full'
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

    $selectedNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($job in $allJobs) {
        $enabled = $job.PSObject.Properties['Enabled'] -and [bool]$job.Enabled
        $assignmentMode = if ($job.PSObject.Properties['AssignmentMode'] -and $job.AssignmentMode) { [string]$job.AssignmentMode } else { 'Legacy' }
        $groupMatches = $Pipeline -eq 'Full' -or ([string]$job.Group -eq $Pipeline)
        if ($enabled -and $assignmentMode -ne 'Manual' -and $groupMatches) { [void]$selectedNames.Add([string]$job.Name) }
    }

    if ($selectedNames.Count -eq 0) { throw "Pipeline '$Pipeline' selected no enabled non-manual jobs." }

    $addedDependencyNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($name in @($selectedNames)) {
            $job = $byName[$name]
            $dependencies = if ($job.PSObject.Properties['DependsOn']) { @($job.DependsOn | ForEach-Object { [string]$_ }) } else { @() }
            foreach ($dependencyName in $dependencies) {
                if (-not $byName.ContainsKey($dependencyName)) { throw "Job '$name' depends on unknown job '$dependencyName'." }
                $dependency = $byName[$dependencyName]
                $dependencyEnabled = $dependency.PSObject.Properties['Enabled'] -and [bool]$dependency.Enabled
                $dependencyMode = if ($dependency.PSObject.Properties['AssignmentMode'] -and $dependency.AssignmentMode) { [string]$dependency.AssignmentMode } else { 'Legacy' }
                if (-not $dependencyEnabled -or $dependencyMode -eq 'Manual') {
                    throw "Job '$name' depends on '$dependencyName', which is disabled or manual and cannot participate in pipeline '$Pipeline'."
                }
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
            [pscustomobject][ordered]@{
                Name = [string]$job.Name
                Group = [string]$job.Group
                AssignmentMode = if ($job.PSObject.Properties['AssignmentMode'] -and $job.AssignmentMode) { [string]$job.AssignmentMode } else { 'Legacy' }
                DependsOn = if ($job.PSObject.Properties['DependsOn']) { @($job.DependsOn | ForEach-Object { [string]$_ }) } else { @() }
                RequiredCapabilities = if ($job.PSObject.Properties['RequiredCapabilities']) { @($job.RequiredCapabilities | ForEach-Object { [string]$_ }) } else { @() }
                IncludedAsDependency = $addedDependencyNames.Contains([string]$job.Name)
            }
        }
    )

    $excludedJobs = @(
        foreach ($job in $allJobs) {
            if ($selectedNames.Contains([string]$job.Name)) { continue }
            $enabled = $job.PSObject.Properties['Enabled'] -and [bool]$job.Enabled
            $assignmentMode = if ($job.PSObject.Properties['AssignmentMode'] -and $job.AssignmentMode) { [string]$job.AssignmentMode } else { 'Legacy' }
            $reason = if (-not $enabled) { 'Disabled' } elseif ($assignmentMode -eq 'Manual') { 'Manual' } else { 'OutsidePipelineGroup' }
            [pscustomobject]@{ Name = [string]$job.Name; Group = [string]$job.Group; Reason = $reason }
        }
    )

    [pscustomobject]@{
        Pipeline = $Pipeline
        SelectedJobs = $selectedJobs
        ExcludedJobs = $excludedJobs
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
    $overallStatus = if ($pending.Count -gt 0) { 'Running' } elseif ($failed.Count -gt 0) { 'Failed' } elseif ($warnings.Count -gt 0) { 'CompletedWithWarnings' } else { 'Success' }

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
            $requestPath = Join-Path -Path $folder.FullName -ChildPath 'request.json'
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
        [Parameter(Mandatory)][ValidateSet('Full', 'AD', 'Exchange', 'Exchange2016', 'M365', 'Intune')][string]$Pipeline,
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
        [Parameter(Mandatory)][ValidateSet('Pending', 'Starting', 'Running', 'RetryScheduled', 'Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted', 'BlockedDependencyFailed', 'BlockedDependencyTimeout', 'Rejected')][string]$Status,
        [int]$Attempt = 0,
        [string]$OwnerServer = '',
        [string]$Detail = '',
        [string]$NotBeforeUtc = '',
        [int]$LockTimeoutSeconds = 15
    )

    $statusPath = Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId -JobName $JobName
    if (-not (Test-Path -LiteralPath $statusPath -PathType Leaf)) { throw "Pipeline job status not found: $statusPath" }
    $lockPath = $statusPath + '.lock'
    $lock = Enter-SmartM365OrchestratorPipelineLock -Path $lockPath -TimeoutSeconds $LockTimeoutSeconds
    try {
        $document = Read-SmartM365OrchestratorPipelineJson -Path $statusPath
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

Export-ModuleMember -Function @(
    'Get-SmartM365OrchestratorPipelinePaths',
    'Get-SmartM365OrchestratorPipelineJobStatusPath',
    'Get-SmartM365OrchestratorPipelineSelection',
    'Get-SmartM365OrchestratorPipelineRunStatus',
    'Get-SmartM365OrchestratorActivePipelineRuns',
    'New-SmartM365OrchestratorPipelineRequest',
    'Set-SmartM365OrchestratorPipelineJobStatus'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDgIbQvqCXoKJkb
# acMCo/NjIG/12Z/cJMQep3H5cWeb+KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBteV0IZaBBbLqNA6O+MiPdbpvb0B2pxQw0Nla3aJD7SMA0GCSqG
# SIb3DQEBAQUABIIBgIiNqDcXV6h2pfGPwQG7NMfHEJxabmA7RwwxYFl9Wy9Za1NE
# WZRvYBrG69/ViRWojJAjiJQyYMT4gYVQ6f9n7+2Q3STZCIK2si2Qwd8v4doDPFsF
# ghsIo5HxHBLjRMU4tgk9CpWZF8mvV3tpjKLNqUXu4bOqGKpQKvD6/wRdmpl3er6K
# /oJXYo2lBO+DwdG5Xgf4uHVKukACaAHdAK/es2tKJv2OM226hGkaC31Df7UHCgu1
# AVL8H6hP2xKNl+IiOVew9T+z80/VY76LA+KqHXDgwx+YOL7Y/Rg1lnZKMhcFzzvl
# irkfTsjLHrz+SzqMUbJMHsln5iF/0aB78h8IDs3d3C8JpBrgEz3Db4qhqWat0La6
# HG/u5Hu9UIf5JTJqKawDlszHgxL9UiTq9AHrpkL+Gx0n5lqlQs+/3BMbhTNIVN9a
# ZBGPKSplhmH1ZvBLfxWb6iD2okN+T2MUUVSArwW4JIZURH4gh9PQ7AaPSCCX1Zdv
# ysgGOk6Fe4LQNYNkOqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjIxNjA3
# MzFaMC8GCSqGSIb3DQEJBDEiBCCb6ELK1706R4dEhVUORVzY3/4yoogJJjsxcB6K
# qFfu0TANBgkqhkiG9w0BAQEFAASCAgB6OQZOoiO8tVLPKKZlExg52rKT5NXsUv5o
# lt1KzEYRyXV0P/dn/EDjysEQAmd6S9TQbizF2TYDZsuMkIp5aQjFM/N3+2as1kiB
# 0RRvNxw9EuzJHIjs5hfbYLn8WbEbUecoVA2vAKWJ321NzHqDZypRqct0CYqtmMNX
# V7rbLwB1ltfDHQu5F1Rsx4GyTvXC0WS3QHeyMFpD/LI1RzZB7Pv0gZPMdGI+62MT
# 5co1+aKCdtxsGJq1I4HWtOUHseIvjtdPRlUzcCKTP/3IPsia0+3ZrOHd9Jr8QEjU
# a/zeLEXJbJHZOfm2enhGMQhqcA2CMJ/KyD96Yq+HKSqY+65/sRmEcqvNIbc3Qt5+
# sawbmD3Tn0doVkcldeVke/+UhW1Dc8wqyOLCkzwXgaiII6Nix0+6FRJUG4MajjAn
# JoJoA1v+6QewvqAi6Jx/1W7IlQwKBo4GCXX/9PLJpGdey3iHP1gWG+KrVcPJVCXX
# dOPVYga6VRIGv4mSsDuMpU9JyQ/OI9FIF1eFznoYQktn/Ow2NT/Jtf6jSgZwUEKS
# dN0OKcAkcfYEgqcTMt6I/DQoy2rDECTWz4lJ3b/l89Lq3IOPehYutf5MoVoCkx6U
# wRs9sThXO6C/dqQZsMW/FwZ8Ge1Z8hHbQC811uQC80L3Tbg7SYY18bC0SDppCwmM
# UZuGzemUbg==
# SIG # End signature block
