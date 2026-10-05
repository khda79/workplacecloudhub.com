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
1.2.0

.EXAMPLE
./SmartM365-Inventory-Pipeline.ps1 -Tenant test -Cancel -BatchId '<active-batch-id>' -Reason 'Controlled direct collection' -ValidateOnly
Read-only cancellation preview. Omit -ValidateOnly to cancel remaining jobs; running collectors continue.
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
    [switch]$Cancel,
    [string]$BatchId = '',
    [string]$Reason = '',
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
    if ($Cancel -and ($Collect -or $NoWait -or $Job.Count -gt 0 -or $IncludeDependencies -or $PSBoundParameters.ContainsKey('Pipeline'))) { throw '-Cancel cannot be combined with collection/selection options.' }
    if ($Cancel -and ([string]::IsNullOrWhiteSpace($BatchId) -or [string]::IsNullOrWhiteSpace($Reason))) { throw '-Cancel requires -BatchId and -Reason.' }
    if (-not $Cancel -and ($BatchId -or $Reason)) { throw '-BatchId and -Reason require -Cancel.' }
    if (-not $ValidateOnly -and -not $Collect -and -not $Cancel) { $ValidateOnly = $true }
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
    if ($Cancel) {
        $cancelStatus = Stop-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $SharedDataFolderPath -BatchId $BatchId -Tenant $Tenant -Reason $Reason -ValidateOnly:$ValidateOnly
        Write-Host ("Cancellation {0}: BatchId={1}; status={2}; cancelled={3}; remaining={4}; failed={5}. Running collectors are not stopped." -f $(if ($ValidateOnly) { 'validated only (no write)' } else { 'published' }), $BatchId, $cancelStatus.OverallStatus, $cancelStatus.CancelledCount, $cancelStatus.PendingCount, $cancelStatus.FailedCount)
        $cancelStatus.Jobs | Format-Table JobName, Status, OwnerServer, Detail -AutoSize
        Complete-PipelineScript -Status Success
        exit 0
    }
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
    if ($status.OverallStatus -eq 'Cancelled') { Complete-PipelineScript -Status CompletedWithWarnings -Warnings 1; exit 3 }
    Complete-PipelineScript -Status Failed -Errors $status.FailedCount
    exit 1
}
catch {
    Write-Host ("Pipeline error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Complete-PipelineScript -Status Failed -Errors 1
    $exitCode = 1
}
exit $exitCode

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDohCbgauZ4BUJ+
# i1TZItFM34PVIqAESeKUuCn7wrlv2qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC26WJv1SAjI6Jx5wc24AEjv7UlSP+XCuONHE729pXv9MA0GCSqG
# SIb3DQEBAQUABIIBgHsGSBvd2iFLw/Y+R3zqqIY1z+u1ZqeNCXPjShmNa8WU5aPM
# /y60KMqvC3PGWKeFkBx/aD++WKUXyrIC3RhMq6JzqjKm5E7wsuvHszHg6IeMKxxP
# k0J4wqc72TiH3PnrOpYlydM3PvIf3SNN1/VOOj6zF+k58e1jgfiAspONK1DQcKan
# IjiJEO/2Q31U1FmZ2ZVJr/7vrB+gaikdynDHGkBhYent7bvzu3OvS65kL0qgSi5q
# Nr9R/2vavKJ4OrZSVHQahVKPobomlbLg/WXF91Ts+IU8y//Y/L+r74uqaxEydPvA
# uGU/c0KikMN6xGxg9/1kplTKOXmIeqh7I0C7ICFRpfBPjW4K/VePBUjgbU5i9dit
# 5sJwaLxEp+S2dwGN5Ykl5FzHgQFjIRe6DZSF8h3Eb+J+GDuB2dqgif8Sj1SjhR5O
# 7PlHEZ1YdoFhpvHO9xiBOyfLFD0jnjzbm8NYE+uOu3G7/pkn/XsdAQsGoRjEcahC
# Dp5QnN/wx9qHCz3xGaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDUxODAy
# MzRaMC8GCSqGSIb3DQEJBDEiBCCkW+lWQuZi4m/iGJx0R/geDdyRxpC63R1NhNs1
# 8VBZ6DANBgkqhkiG9w0BAQEFAASCAgCoxVK3Kv/fXGiyX0KHCtE6sEULq0g41Tci
# 25jLdPrOrr+VwN2iF5dhrfaRJnMF7Hma1YGfDN9d64mb8FAp6MVL6q5dH5NeJ1To
# b4Btas6HLcabU/ITUq+QPVUQgj/euJ2/sdScu3IaWAzPd3/zg/GSxpUChxJbbcEB
# IyGgp2Dpmp3HlfDVwEyL2tF6FB9fyFEw9HWx4RPuHbHeKzdub654RSxcMqhuIbcR
# NXEcZtLGvqvHERmd5mzGwvBsp8Ff0ma5niBa6nKhhWxzs8pyfGyHORV5VlapAyGY
# /ne+YFUBeXkIjiIY3z8y5nKqfTV11dy/FuJxglhr4Hvk82/d1+3AxNX2jXRit5Tk
# YRLXFyzpdA1A+aKcSuZyeoRGuN9Ux6LZ4jy4dE9b9/w1M+p8anHRG8rTEL/lljjE
# pLR4FE2NkU+aj7PVW7ibQs7xJMYBXNDu/DpOORo1mp2DSte7ZTaXJ7cRAookQ6Lg
# mLfZT5ec32GG9zY0JoS08fxYftSwIoBgwNXaz/YUOgylFB73cvRQF4hUwDhXy+ra
# dapzR0rC89h9Jcus+XEah7CvReFXCUYRv7M9LMP4490DFMwAa+pbTCgPlIKK92Kn
# VBtJmDbI+VexL87/mibF+oJ0HGXyCZhrAjg484I9pJqpEhiX/o1FbNF2yWJzCYkp
# UNQKoxLOTQ==
# SIG # End signature block
