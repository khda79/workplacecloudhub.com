<#
.SYNOPSIS
Submits a distributed SmartInventory pipeline request to resident orchestrators.

.DESCRIPTION
Builds a plan from the effective central jobs manifest. ValidateOnly is read-only. Collect
publishes one atomic shared request that resident orchestrators consume while preserving
ownership, capability, dependency, claim, lock and concurrency controls.

.EXAMPLE
./SmartM365-Inventory-Pipeline.ps1 -Tenant prod -Pipeline Full -ValidateOnly

.EXAMPLE
./SmartM365-Inventory-Pipeline.ps1 -Tenant prod -Pipeline Full -Collect

.EXAMPLE
./SmartM365-Inventory-Pipeline.ps1 -Tenant prod -Job WorkplaceEvidence-Prepare -Collect
Runs only the named job(s) through the resident orchestrators. Dependencies are not rerun: the
orchestrator applies the job's scheduled dependency rule to them (for example FreshSuccess).
Add -IncludeDependencies to run the enabled dependencies in the same request.

.VERSION
1.1.0
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [ValidateSet('Full', 'AD', 'Exchange', 'Exchange2016', 'M365', 'Intune')]
    [string]$Pipeline = 'Full',
    [string[]]$Job = @(),
    [switch]$IncludeDependencies,
    [switch]$ValidateOnly,
    [switch]$Collect,
    [switch]$NoWait,
    [ValidateRange(1, 60)][int]$PollSeconds = 15,
    [ValidateRange(1, 720)][int]$WaitTimeoutHours = 168,
    [string]$JobsManifestPath = '',
    [string]$SharedDataFolderPath = ''
)

$ErrorActionPreference = 'Stop'
$startedAt = Get-Date
$scriptName = 'SmartM365-Inventory-Pipeline'
$smartM365Root = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$tenantContextPath = Join-Path -Path $smartM365Root -ChildPath 'Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) { throw "SmartM365 tenant context not found: $tenantContextPath" }
. $tenantContextPath
# Initialize-SmartM365TenantContext displays the startup banner.

function Write-Host {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)][object]$Object,
        [switch]$NoNewline,
        [object]$Separator,
        [ConsoleColor]$ForegroundColor,
        [ConsoleColor]$BackgroundColor
    )
    $optional = @{}
    foreach ($name in @('ForegroundColor', 'BackgroundColor', 'Separator')) {
        if ($PSBoundParameters.ContainsKey($name)) { $optional[$name] = $PSBoundParameters[$name] }
    }
    if ($NoNewline) { $optional.NoNewline = $true }
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    foreach ($line in ([string]$Object -split "`r`n|`n")) {
        $displayText = if ($line) { "[$stamp] $line" } else { '' }
        Microsoft.PowerShell.Utility\Write-Host $displayText @optional
    }
}

function Resolve-PipelineConfigValue {
    param([AllowNull()]$Value)
    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $property = $script:EffectiveConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $replacement = Resolve-PipelineConfigValue -Value $property.Value
            if ($null -eq $replacement) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$replacement)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    $resolved
}

function Get-PipelineConfigValue {
    param($LocalConfig, [string]$Name, $DefaultValue)
    $property = $LocalConfig.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        $text = if ($property.Value -is [string]) { $property.Value.Trim() } else { '' }
        if ($property.Value -isnot [string] -or ($text -and $text -notin @('__USE_GLOBAL__', 'USE_GLOBAL'))) {
            return Resolve-PipelineConfigValue -Value $property.Value
        }
    }
    $globalProperty = $script:EffectiveConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        return Resolve-PipelineConfigValue -Value $globalProperty.Value
    }
    $DefaultValue
}

function Complete-PipelineScript {
    param([ValidateSet('Success', 'Failed', 'CompletedWithWarnings')][string]$Status, [int]$Warnings = 0, [int]$Errors = 0)
    Write-SmartM365CompletionBanner -Status $Status -ScriptName $scriptName -StartedAt $startedAt -EndedAt (Get-Date) -WarningCount $Warnings -ErrorCount $Errors -GeneratedCsvFiles 0
}

$exitCode = 0
try {
    if ($ValidateOnly -and $Collect) { throw 'Choose either -ValidateOnly or -Collect, not both.' }
    if ($NoWait -and -not $Collect) { throw '-NoWait is valid only with -Collect.' }
    if (-not $ValidateOnly -and -not $Collect) { $ValidateOnly = $true }
    # A launcher passing a comma-separated list arrives as one string.
    $Job = @($Job | ForEach-Object { ([string]$_) -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($Job.Count -gt 0 -and $PSBoundParameters.ContainsKey('Pipeline')) { throw 'Choose either -Pipeline or -Job, not both.' }
    if ($IncludeDependencies -and $Job.Count -eq 0) { throw '-IncludeDependencies is valid only with -Job.' }
    if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'This script requires PowerShell 7 or later.' }

    $managementModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Orchestrator.Management.psm1'
    $pipelineModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Orchestrator.Pipeline.psm1'
    Import-Module -Name $managementModulePath -Force -ErrorAction Stop
    Import-Module -Name $pipelineModulePath -Force -ErrorAction Stop

    if ([string]::IsNullOrWhiteSpace($JobsManifestPath) -xor [string]::IsNullOrWhiteSpace($SharedDataFolderPath)) {
        throw 'Use -JobsManifestPath and -SharedDataFolderPath together, or omit both.'
    }

    if ([string]::IsNullOrWhiteSpace($JobsManifestPath)) {
        $script:EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
        $localPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Inventory-Orchestrator.local.json'; $localPath = Resolve-SmartM365JsonConfigurationPath -Path $localPath
        $templatePath = (Get-SmartM365JsonTemplateName -Path $localPath)
        $configPath = if (Test-Path -LiteralPath $localPath -PathType Leaf) { $localPath } else { $templatePath }
        if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) { throw "Orchestrator local configuration not found: $localPath" }
        $localConfig = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json -Depth 100
        $dataFolder = [string](Get-PipelineConfigValue -LocalConfig $localConfig -Name 'OrchestratorDataFolderPath' -DefaultValue (Join-Path $PSScriptRoot 'Output'))
        $SharedDataFolderPath = if ((Split-Path -Path $dataFolder -Leaf) -eq $env:COMPUTERNAME) { Split-Path -Path $dataFolder -Parent } else { $dataFolder }
        $centralEnabledValue = Get-PipelineConfigValue -LocalConfig $localConfig -Name 'CentralConfigurationEnabled' -DefaultValue $true
        $centralEnabled = if ($centralEnabledValue -is [bool]) { $centralEnabledValue } else { [string]$centralEnabledValue -in @('true', 'True', '1') }
        if ($centralEnabled) {
            $JobsManifestPath = (Get-SmartM365OrchestratorConfigurationPaths -SharedDataFolderPath $SharedDataFolderPath).JobsPath
        }
        else {
            $JobsManifestPath = Join-Path -Path $PSScriptRoot -ChildPath 'Orchestrator-Jobs.json'
        }
    }

    $SharedDataFolderPath = [IO.Path]::GetFullPath($SharedDataFolderPath)
    $JobsManifestPath = Get-SmartM365JsonReadPath ([IO.Path]::GetFullPath($JobsManifestPath))
    if (-not (Test-Path -LiteralPath $JobsManifestPath -PathType Leaf)) { throw "Effective jobs manifest not found: $JobsManifestPath" }
    if ($Collect -and -not (Test-Path -LiteralPath $SharedDataFolderPath -PathType Container)) { throw "Shared orchestrator data folder not found: $SharedDataFolderPath" }

    $jobsDocument = (Read-SmartM365JsonDocument $JobsManifestPath).Document
    $validation = Test-SmartM365OrchestratorJobsDocument -Document $jobsDocument
    if (-not $validation.Valid) { throw "Invalid jobs manifest: $($validation.Errors -join '; ')" }
    $selection = if ($Job.Count -gt 0) {
        Get-SmartM365OrchestratorPipelineSelection -JobsDocument $jobsDocument -JobName $Job -IncludeDependencies:$IncludeDependencies
    }
    else {
        Get-SmartM365OrchestratorPipelineSelection -JobsDocument $jobsDocument -Pipeline $Pipeline
    }
    $selectionName = [string]$selection.Pipeline

    $electionPlanPath = Join-Path -Path $SharedDataFolderPath -ChildPath 'Election\Orchestrator-ElectionPlan.json'
    $electionPlan = if (Get-SmartM365JsonReadPath $electionPlanPath -Optional) { (Read-SmartM365JsonDocument $electionPlanPath).Document } else { $null }
    $planRows = foreach ($plannedJob in @($selection.SelectedJobs)) {
        $source = @($jobsDocument.Jobs | Where-Object { [string]$_.Name -eq [string]$plannedJob.Name } | Select-Object -First 1)[0]
        $owner = switch ([string]$plannedJob.AssignmentMode) {
            'Elected' {
                if ($null -ne $electionPlan) { [string](@($electionPlan.Assignments | Where-Object { [string]$_.JobName -eq [string]$plannedJob.Name } | Select-Object -First 1).OwnerServer) } else { '' }
            }
            'Pinned' { [string](@($source.AllowedServers)[0]) }
            default { if (@($source.AllowedServers).Count -gt 0) { @($source.AllowedServers) -join ',' } else { 'Any eligible server' } }
        }
        [pscustomobject]@{ Job = $plannedJob.Name; Group = $plannedJob.Group; Mode = $plannedJob.AssignmentMode; Owner = $owner; Dependency = @($plannedJob.DependsOn) -join ','; External = @($plannedJob.ExternalDependencies).Count }
    }
    $missingOwners = @($planRows | Where-Object { $_.Mode -eq 'Elected' -and [string]::IsNullOrWhiteSpace($_.Owner) })

    Write-Host ("Pipeline plan: tenant={0}; pipeline={1}; selected={2}; excluded={3}; dependencyClosure={4}." -f $Tenant, $selectionName, $selection.SelectedCount, $selection.ExcludedCount, $selection.AddedDependencyCount) -ForegroundColor Cyan
    Write-Host ("Manifest: {0}" -f $JobsManifestPath) -ForegroundColor Gray
    Write-Host ("Shared data: {0}" -f $SharedDataFolderPath) -ForegroundColor Gray
    foreach ($row in $planRows) {
        $ownerText = if ([string]::IsNullOrWhiteSpace($row.Owner)) { '<unassigned>' } else { $row.Owner }
        $dependencyText = if ([string]::IsNullOrWhiteSpace($row.Dependency)) { '-' } else { $row.Dependency }
        Write-Host ("{0,-55} group={1,-12} mode={2,-7} owner={3} dependsOn={4}" -f $row.Job, $row.Group, $row.Mode, $ownerText, $dependencyText) -ForegroundColor Gray
        if ($row.External -gt 0) {
            Write-Host ("{0,-55} {1} dependency(ies) outside this request: the orchestrator applies the job's scheduled dependency rule (no stale data)." -f '', $row.External) -ForegroundColor Gray
        }
    }
    foreach ($ignored in @($selection.IgnoredDependencies)) {
        Write-Host ("Dependency ignored (disabled or manual, as in the schedule): {0} -> {1}" -f $ignored.Job, $ignored.Dependency) -ForegroundColor Yellow
    }
    if ($validation.Warnings.Count -gt 0) { foreach ($warning in $validation.Warnings) { Write-Host $warning -ForegroundColor Yellow } }
    if ($missingOwners.Count -gt 0) {
        $message = "No elected owner in the current shared plan for: $(@($missingOwners.Job) -join ', ')."
        if ($Collect) { throw $message }
        Write-Host $message -ForegroundColor Yellow
    }

    if ($ValidateOnly) {
        $warningCount = @($validation.Warnings).Count + $missingOwners.Count
        Complete-PipelineScript -Status $(if ($warningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }) -Warnings $warningCount
        exit $(if ($warningCount -gt 0) { 3 } else { 0 })
    }

    $request = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $SharedDataFolderPath -Tenant $Tenant -Pipeline $selectionName -Selection $selection -ManifestHash ((Get-FileHash -LiteralPath $JobsManifestPath -Algorithm SHA256).Hash) -RequestedBy ([Environment]::UserName) -RequestedFrom ([Environment]::MachineName)
    Write-Host ("Pipeline request submitted: BatchId={0}; request={1}" -f $request.BatchId, $request.RequestPath) -ForegroundColor Green
    if ($NoWait) {
        Complete-PipelineScript -Status Success
        exit 0
    }

    $deadline = [datetime]::UtcNow.AddHours($WaitTimeoutHours)
    $lastSignature = ''
    do {
        $status = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $SharedDataFolderPath -BatchId $request.BatchId
        $signature = (@($status.Jobs | Sort-Object JobName | ForEach-Object { "$($_.JobName)=$($_.Status)@$($_.OwnerServer)" }) -join ';')
        if ($signature -ne $lastSignature) {
            foreach ($jobStatus in @($status.Jobs | Sort-Object JobName)) {
                Write-Host ("Batch {0}: {1} -> {2} (owner={3})" -f $request.BatchId, $jobStatus.JobName, $jobStatus.Status, $jobStatus.OwnerServer) -ForegroundColor Gray
            }
            $lastSignature = $signature
        }
        if ($status.IsTerminal) { break }
        if ([datetime]::UtcNow -ge $deadline) { throw "Pipeline wait timed out after $WaitTimeoutHours hour(s). BatchId=$($request.BatchId)." }
        Start-Sleep -Seconds $PollSeconds
    } while ($true)

    Write-Host ("Pipeline completed: BatchId={0}; status={1}; failed={2}; warnings={3}." -f $status.BatchId, $status.OverallStatus, $status.FailedCount, $status.WarningCount) -ForegroundColor $(if ($status.OverallStatus -eq 'Success') { 'Green' } elseif ($status.OverallStatus -eq 'CompletedWithWarnings') { 'Yellow' } else { 'Red' })
    if ($status.OverallStatus -eq 'Success') { Complete-PipelineScript -Status Success; exit 0 }
    if ($status.OverallStatus -eq 'CompletedWithWarnings') { Complete-PipelineScript -Status CompletedWithWarnings -Warnings $status.WarningCount; exit 3 }
    Complete-PipelineScript -Status Failed -Errors $status.FailedCount
    exit 1
}
catch {
    Write-Host ("Pipeline error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Complete-PipelineScript -Status Failed -Errors 1
    $exitCode = 1
}
exit $exitCode
