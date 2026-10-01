<#
.SYNOPSIS
Previews or publishes the committed template schedule plan in the shared Orchestrator jobs configuration.

.DESCRIPTION
For every shared job that also exists in Orchestrator-Jobs.json.template, the planning fields are
taken from the template: Schedule, TimeoutMinutes, DependsOn, ConcurrencyKey, DependencyMode and
DependencyMaxAgeHours (removed when the template has none). Jobs listed in -RemoveJob are removed.
Every other field (Arguments, Enabled, AssignmentMode, AllowedServers, retries, capabilities,
launcher, estimated duration) and every cluster setting are preserved. Shared jobs absent from the
template are kept unchanged and reported; template jobs absent from the shared configuration are
reported (the resident orchestrator appends them at startup). Enabled differences are reported and
kept, unless -SyncEnabled is used. Without -Execute, the script is read-only.
Publication uses the Orchestrator management module's validation, optimistic hash checks,
configuration lock, before/after version snapshots and audit CSV.

.PARAMETER SharedDataFolderPath
Shared tenant Orchestrator root containing Config, Versions and Audit, for example
\\server\share\DATA-ALL\Orchestrator.

.PARAMETER TemplatePath
Jobs template to apply. Defaults to Orchestrator-Jobs.json.template next to this script.

.PARAMETER RemoveJob
Shared jobs to remove. Defaults to Exchange2016-Local-Mailboxes-Fast (same script and arguments
as Exchange2016-Local-Mailboxes-Inventory).

.PARAMETER SyncEnabled
Also takes the Enabled flag from the template for the jobs it contains. The changed flags are
shown in the preview like the other fields.

.PARAMETER Execute
Publishes the displayed changes. Omit this switch for preview only.

.VERSION
1.1
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SharedDataFolderPath,

    [string]$TemplatePath = '',

    [string[]]$RemoveJob = @('Exchange2016-Local-Mailboxes-Fast'),

    [switch]$SyncEnabled,

    [switch]$Execute
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$tenantContextPath = Join-Path -Path $PSScriptRoot -ChildPath '..\..\Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) { throw "SmartM365 tenant context not found: $tenantContextPath" }
. $tenantContextPath
Write-SmartM365StartupBanner

$managementModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Orchestrator.Management.psm1'
Import-Module -Name $managementModulePath -Force -ErrorAction Stop
if ([string]::IsNullOrWhiteSpace($TemplatePath)) { $TemplatePath = Join-Path -Path $PSScriptRoot -ChildPath 'Orchestrator-Jobs.json.template' }
$template = Get-Content -LiteralPath $TemplatePath -Raw | ConvertFrom-Json -Depth 100
$templateValidation = Test-SmartM365OrchestratorJobsDocument -Document $template
if (-not $templateValidation.Valid) { throw "The template is invalid: $($templateValidation.Errors -join '; ')" }
$templateJobs = @{}
foreach ($templateJob in @($template.Jobs)) { $templateJobs[[string]$templateJob.Name] = $templateJob }

$snapshot = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $SharedDataFolderPath
$jobsDocument = $snapshot.Jobs | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
$changes = [System.Collections.Generic.List[object]]::new()
$notes = [System.Collections.Generic.List[string]]::new()

function ConvertTo-PlanText {
    # Canonical, readable text: property order and empty-versus-absent lists never count as changes.
    param([Parameter(Mandatory = $true)][string]$Name, [AllowNull()]$Value)
    if ($Name -eq 'DependsOn') {
        $names = @($Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ } | Sort-Object)
        if ($names.Count -eq 0) { return '<none>' }
        return ($names -join ', ')
    }
    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) { return '<none>' }
    if ($Name -eq 'Schedule') {
        $days = if ($Value.PSObject.Properties['DaysOfWeek'] -and [string]$Value.Type -eq 'Weekly') { ' ' + ((@($Value.DaysOfWeek) | ForEach-Object { ([string]$_).Substring(0, 3) }) -join '/') } else { '' }
        $missed = if ($Value.PSObject.Properties['MissedRunPolicy'] -and $Value.MissedRunPolicy) { [string]$Value.MissedRunPolicy } else { 'RunOnce' }
        return ('{0}{1} {2} [{3}]' -f $Value.Type, $days, ((@($Value.Times) | Sort-Object) -join ','), $missed)
    }
    return [string]$Value
}

function Merge-PlanField {
    param([Parameter(Mandatory = $true)]$Job, [Parameter(Mandatory = $true)][string]$Name, [AllowNull()]$Value)
    $current = if ($Job.PSObject.Properties[$Name]) { $Job.$Name } else { $null }
    $before = ConvertTo-PlanText -Name $Name -Value $current
    $after = ConvertTo-PlanText -Name $Name -Value $Value
    if ($before -ceq $after) { return }
    if ($null -eq $Value) { [void]$Job.PSObject.Properties.Remove($Name) }
    elseif ($Job.PSObject.Properties[$Name]) { $Job.$Name = $Value }
    else { $Job | Add-Member -NotePropertyName $Name -NotePropertyValue $Value }
    $changes.Add([pscustomobject][ordered]@{ Job = [string]$Job.Name; Field = $Name; Before = $before; After = $after })
}

$keptJobs = foreach ($job in @($jobsDocument.Jobs)) {
    $name = [string]$job.Name
    if ($name -in $RemoveJob) {
        $changes.Add([pscustomobject][ordered]@{ Job = $name; Field = '<job>'; Before = 'present'; After = 'removed' })
        continue
    }
    if (-not $templateJobs.ContainsKey($name)) {
        $notes.Add("Kept unchanged (absent from the template): $name")
        $job
        continue
    }
    $templateJob = $templateJobs[$name]
    foreach ($field in @('Schedule', 'TimeoutMinutes', 'DependsOn', 'ConcurrencyKey', 'DependencyMode', 'DependencyMaxAgeHours')) {
        $value = if ($templateJob.PSObject.Properties[$field]) { $templateJob.$field } else { $null }
        if ($field -eq 'DependsOn' -and $null -eq $value) { $value = @() }
        Merge-PlanField -Job $job -Name $field -Value $value
    }
    if ($SyncEnabled) { Merge-PlanField -Job $job -Name 'Enabled' -Value ([bool]$templateJob.Enabled) }
    elseif ([bool]$job.Enabled -ne [bool]$templateJob.Enabled) {
        $notes.Add(("Enabled differs and is kept: {0} shared={1}, template={2}" -f $name, [bool]$job.Enabled, [bool]$templateJob.Enabled))
    }
    $job
}
$jobsDocument.Jobs = @($keptJobs)
foreach ($name in @($templateJobs.Keys | Where-Object { $_ -notin @($jobsDocument.Jobs.Name) -and $_ -notin $RemoveJob } | Sort-Object)) {
    $notes.Add("Template job absent from the shared configuration (appended by the orchestrator at startup): $name")
}

$validation = Test-SmartM365OrchestratorJobsDocument -Document $jobsDocument
$consistency = Test-SmartM365OrchestratorConfigurationConsistency -JobsDocument $jobsDocument -ClusterDocument $snapshot.Cluster
$errors = @($validation.Errors) + @($consistency.Errors)
if ($errors.Count -gt 0) { throw "The resulting configuration is invalid: $($errors -join '; ')" }

if ($changes.Count -eq 0) {
    Write-Output 'The shared jobs configuration already matches the template schedule plan.'
}
else {
    foreach ($group in @($changes | Group-Object Job | Sort-Object Name)) {
        Write-Output $group.Name
        foreach ($change in $group.Group) { Write-Output ("    {0,-22} {1}  ->  {2}" -f $change.Field, $change.Before, $change.After) }
    }
    Write-Output ''
    Write-Output ("{0} change(s) on {1} job(s)." -f $changes.Count, @($changes.Job | Sort-Object -Unique).Count)
}
foreach ($note in $notes) { Write-Output $note }
foreach ($warning in @($validation.Warnings) + @($consistency.Warnings)) { Write-Warning $warning }
if ($changes.Count -eq 0) { return }

if (-not $Execute) {
    Write-Warning 'Preview only. Re-run the same command with -Execute to publish these changes.'
    return
}

$result = Publish-SmartM365OrchestratorConfiguration `
    -SharedDataFolderPath $SharedDataFolderPath `
    -JobsDocument $jobsDocument `
    -ClusterDocument $snapshot.Cluster `
    -ExpectedJobsHash $snapshot.JobsHash `
    -ExpectedClusterHash $snapshot.ClusterHash `
    -ChangeSummary ("Apply template schedule plan: {0} change(s) on {1} job(s)" -f $changes.Count, @($changes.Job | Sort-Object -Unique).Count)

Write-Output ("Published configuration version {0}." -f $result.VersionId)
Write-Output ("New jobs hash: {0}" -f $result.JobsHash)
if (@($result.Warnings).Count -gt 0) {
    $result.Warnings | ForEach-Object { Write-Warning $_ }
}
