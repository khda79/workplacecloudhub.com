#Requires -Version 7.0
<#
.SYNOPSIS
Validate offline that the hand-written SmartInventory launchers stay aligned with the jobs template.
.VERSION
1.0.0
.NOTES
Every template job is either started through its LauncherPath, mapped to the manual launcher that
reproduces it, or explicitly covered only by the generated Launchers\Jobs launcher.
A LauncherPath launcher must forward the orchestrator Arguments (%*), follow the orchestrator
tenant and use the engine of PowerShellEdition. A mapped manual launcher must run the job script
with that engine and pass every template Argument. Reads files only: nothing is executed.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
$inventoryRoot = Join-Path $smartRoot 'SmartInventory'
$launchersRoot = Join-Path $inventoryRoot 'Launchers'
$checks = 0
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}

# Manual launcher (relative to Launchers) per job without LauncherPath. $null = only the generated
# Launchers\Jobs\Start-<JobName>.cmd reproduces the job.
$manualLaunchers = [ordered]@{
    'M365-ActiveUsers-Inventory' = 'Cloud\Start-SmartM365-ActiveUsers-Inventory.cmd'
    'M365-EntraDevices-Inventory' = 'Cloud\Start-SmartM365-EntraDevices-Inventory.cmd'
    'Intune-Devices-Inventory' = 'Cloud\Start-SmartM365-Devices-Inventory.cmd'
    'M365-VerifiedDomains-Inventory' = 'Cloud\Start-SmartM365-VerifiedDomains-Inventory.cmd'
    'EXO-AcceptedDomains-Inventory' = 'Cloud\Start-SmartM365-EXO-AcceptedDomains-Inventory.cmd'
    'M365-Licences-Inventory' = 'Cloud\Start-SmartM365-Licences-Inventory.cmd'
    'Intune-Devices-BIOS-Inventory' = 'Cloud\Start-SmartM365-Devices-BIOS-Inventory.cmd'
    'Intune-DeviceSystem-Inventory' = 'Cloud\Start-SmartM365-Device-System-Inventory.cmd'
    'Intune-Devices-Compliance-Inventory' = 'Cloud\Start-SmartM365-Devices-Compliance-Inventory.cmd'
    'Intune-Devices-UpgradeEligibility' = 'Cloud\Start-SmartM365-Devices-UpgradeEligibility.cmd'
    'EXO-Mailboxes-Inventory-Fast' = 'Cloud\Start-SmartM365-EXO-Mailboxes-Inventory.cmd'
    'EXO-Mailboxes-Inventory' = 'Cloud\Start-SmartM365-EXO-Mailboxes-Inventory-IncludeStats.cmd'
    'EXO-Mailboxes-Permissions' = $null
    'Intune-WinUpdate-Status' = 'Cloud\Start-SmartM365-WinUpdate_Status_From_Intune.cmd'
    'Intune-Windows11-Readiness-Issues' = 'Cloud\Start-SmartM365-Intune-Windows11-Readiness-Issues-Inventory.cmd'
    'M365-SecureScore-Inventory' = 'Cloud\Start-SmartM365-SecureScore-Inventory.cmd'
    'M365-AuthenticationMethodsRegistration-Inventory' = 'Cloud\Start-SmartM365-AuthenticationMethodsRegistration-Inventory.cmd'
    'M365-ConditionalAccess-Inventory' = 'Cloud\Start-SmartM365-ConditionalAccess-Inventory.cmd'
    'Intune-EndpointSecurityHealth-Inventory' = 'Cloud\Start-SmartM365-Intune-EndpointSecurityHealth-Inventory.cmd'
    'Intune-EndpointAnalytics-Inventory' = 'Cloud\Start-SmartM365-EndpointAnalytics-Inventory-IncludeStartupProcesses.cmd'
    'Intune-AutopatchAlerts-Inventory' = 'Cloud\Start-SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.cmd'
    'Intune-RBAC-GroupMembers' = 'Cloud\Start-SmartM365-Intune-RBAC-GroupMembers.cmd'
    'Intune-Remediations-Export' = 'Cloud\Start-SmartM365-Intune-ExportRemediationScripts.cmd'
    'M365-BackupProtectedMailboxes-Inventory' = 'Cloud\Start-SmartM365-M365-BackupProtectedMailboxes-Inventory-Interactive.cmd'
    'M365-BackupPolicyScope-Inventory' = 'Cloud\Start-SmartM365-M365-BackupPolicyScope-Inventory.cmd'
    'M365-BackupProtectedSitesAndDrives-Inventory' = 'Cloud\Start-SmartM365-M365-BackupProtectedSitesAndDrives-Inventory-Interactive.cmd'
    'M365-PowerBIFabricActivity-Inventory' = 'Cloud\Start-SmartM365-PowerBIFabricActivity-Inventory.cmd'
    'M365-Usage-Inventory' = 'Cloud\Start-SmartM365-M365UsageReports-Inventory.cmd'
    'M365-CopilotUsage-Inventory' = 'Cloud\Start-SmartM365-CopilotUsage-Inventory.cmd'
    'M365-TeamsPhonePstnUsage-Inventory' = 'Cloud\Start-SmartM365-TeamsPhonePstnUsage-Inventory-IncludePhoneAssignments.cmd'
    'M365-LicensePricing-Inventory' = 'Cloud\Start-SmartM365-LicensePricing-Inventory.cmd'
    'SmartFinOps-Workplace-Valuation' = $null
    'SmartFinOps-Workplace-Report' = $null
    'Exchange-HybridIdentity-Issues' = 'Cloud\Start-SmartM365-Exchange-HybridIdentity-Issues-Inventory.cmd'
    'Mailboxes-PermissionsByUser-Report' = 'Cloud\Start-SmartM365-Mailboxes-AllSources-PermissionsByUserReport.cmd'
    'WorkplaceEvidence-Prepare' = 'Cloud\Start-SmartM365-WorkplaceEvidence-Prepare.cmd'
    'M365-SyncHealth-Inventory' = 'Cloud\Start-SmartM365-AzureADConnect-SyncHealth-Inventory.cmd'
    'EXO-MigrationJobs-Inventory' = 'Cloud\Start-SmartM365-EXO-MigJob-Inventory.cmd'
    'EXO-Mailboxes-CalendarPermissions' = 'Cloud\Start-SmartM365-EXO-Mailboxes-CalPerm_Inventory.cmd'
    'Intune-DiscoveredApps-Inventory' = 'Cloud\Start-SmartM365-Intune-DiscoveredApps-Inventory.cmd'
    'M365-Teams-Inventory' = 'Cloud\Start-SmartM365-Teams-Inventory.cmd'
    'M365-SPO-Inventory' = 'Cloud\Start-SmartM365-SPO-Inventory.cmd'
    'Intune-WindowsAutopilot-Inventory' = 'Cloud\Start-SmartM365-WindowsAutopilot-Inventory.cmd'
    'M365-WorkplaceScope-Inventory' = $null
    'Intune-Configurations-Inventory' = 'Cloud\Start-SmartM365-IntuneConfigurations-Inventory.cmd'
    'Intune-ManagedApps-Inventory' = 'Cloud\Start-SmartM365-IntuneManagedApps-Inventory.cmd'
    'Intune-EnrollmentFailures-Inventory' = 'Cloud\Start-SmartM365-IntuneEnrollmentFailures-Inventory.cmd'
}

function Get-LauncherInvocation {
    # Returns the engine, script leaf, fixed arguments and forwarding of the launcher's final
    # PowerShell call, with the launcher's own set "NAME=value" variables expanded.
    param([string]$Path)
    $lines = @(Get-Content -LiteralPath $Path)
    $variables = @{}
    foreach ($line in $lines) {
        if ($line -match '^\s*set "([A-Za-z0-9_]+)=(.*)"\s*$') { $variables[$Matches[1]] = $Matches[2] }
    }
    $call = @($lines | Where-Object { $_ -match '-File\s+"' -and $_ -notmatch 'Unblock-File' } | Select-Object -Last 1)
    if ($call.Count -eq 0 -or $call[0] -notmatch '^\s*"(%[A-Z0-9_]+%)"\s.*-File\s+"([^"]+)"(.*)$') { throw "No PowerShell call found in launcher: $Path" }
    $engineVariable = $Matches[1]
    $scriptPath = $Matches[2]
    $rest = $Matches[3]
    for ($i = 0; $i -lt 5; $i++) { foreach ($name in $variables.Keys) { $scriptPath = $scriptPath.Replace("%$name%", $variables[$name]) } }
    $engine = switch ($engineVariable) { '%PWSH%' { 'PowerShell7' } '%POWERSHELL5%' { 'WindowsPowerShell' } default { 'Unknown' } }
    $tenant = if ($rest -match '-Tenant\s+"?([^\s"]+)"?') { $Matches[1] } else { '' }
    $fixed = (($rest -replace '%\*', ' ' -replace '-Tenant\s+"?[^\s"]+"?', ' ') -replace '\s+', ' ').Trim()
    [pscustomobject]@{
        Engine = $engine
        ScriptLeaf = ($scriptPath -split '[\\/]')[-1]
        Tenant = $tenant
        FixedArguments = $fixed
        ForwardsArguments = $rest -match '%\*'
    }
}

function Get-ArgumentGroup {
    # Splits "-Name value -Switch" into "-Name value", "-Switch".
    param([string]$Arguments)
    $groups = [System.Collections.Generic.List[string]]::new()
    foreach ($token in @(([string]$Arguments).Trim() -split '\s+' | Where-Object { $_ })) {
        if ($token.StartsWith('-') -or $groups.Count -eq 0) { $groups.Add($token) } else { $groups[$groups.Count - 1] += ' ' + $token }
    }
    return $groups
}

$template = Get-Content -LiteralPath (Join-Path $inventoryRoot 'Orchestrator/Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json -Depth 100
foreach ($job in @($template.Jobs)) {
    $name = [string]$job.Name
    $edition = if ($job.PSObject.Properties['PowerShellEdition'] -and $job.PowerShellEdition) { [string]$job.PowerShellEdition } else { 'PowerShell7' }
    $scriptLeaf = Split-Path -Path ([string]$job.ScriptPath) -Leaf
    $launcherPath = if ($job.PSObject.Properties['LauncherPath'] -and $job.LauncherPath) { [string]$job.LauncherPath } else { '' }
    if ($launcherPath) {
        # The orchestrator runs: cmd /c call "<LauncherPath>" <Arguments>, without -Tenant or -Connect.
        $path = Join-Path $inventoryRoot $launcherPath
        Assert-True (Test-Path -LiteralPath $path -PathType Leaf) "Job $name LauncherPath not found: $launcherPath"
        $invocation = Get-LauncherInvocation -Path $path
        Assert-True ($invocation.ScriptLeaf -eq $scriptLeaf) "Job $name launcher runs $($invocation.ScriptLeaf), not $scriptLeaf."
        Assert-True ($invocation.Engine -eq $edition) "Job $name launcher uses $($invocation.Engine), the job declares $edition."
        Assert-True $invocation.ForwardsArguments "Job $name launcher does not forward %*: orchestrator Arguments would be dropped."
        Assert-True ((Get-Content -LiteralPath $path -Raw) -match 'SMARTM365_ORCHESTRATOR_TENANT' -and $invocation.Tenant -ne 'prod') "Job $name launcher ignores the orchestrator tenant."
        $duplicated = @(Get-ArgumentGroup $job.Arguments | Where-Object { (' ' + $invocation.FixedArguments + ' ').Contains(' ' + ($_ -split ' ')[0] + ' ') })
        Assert-True ($duplicated.Count -eq 0) "Job $name passes $($duplicated -join ', ') twice (template Arguments and launcher)."
        Assert-True (-not $manualLaunchers.Contains($name)) "Job $name has a LauncherPath; remove it from the manual launcher map."
        continue
    }
    Assert-True ($manualLaunchers.Contains($name)) "Job $name has no LauncherPath and no manual launcher decision in this test."
    $manual = $manualLaunchers[$name]
    if ($null -eq $manual) { continue }
    $path = Join-Path $launchersRoot $manual
    Assert-True (Test-Path -LiteralPath $path -PathType Leaf) "Job $name manual launcher not found: $manual"
    $invocation = Get-LauncherInvocation -Path $path
    Assert-True ($invocation.ScriptLeaf -eq $scriptLeaf) "Job $name manual launcher runs $($invocation.ScriptLeaf), not $scriptLeaf."
    Assert-True ($invocation.Engine -eq $edition) "Job $name manual launcher uses $($invocation.Engine), the job declares $edition."
    foreach ($group in @(Get-ArgumentGroup $job.Arguments)) {
        Assert-True ((' ' + $invocation.FixedArguments + ' ').Contains(' ' + $group + ' ')) "Job $name manual launcher $manual does not pass '$group' (template Arguments: $($job.Arguments))."
    }
}
foreach ($name in $manualLaunchers.Keys) {
    Assert-True (@($template.Jobs | Where-Object Name -eq $name).Count -eq 1) "Manual launcher map names an unknown job: $name"
}
Write-Output ("PASS: {0} SmartInventory launcher alignment checks. No launcher or collector executed." -f $checks)
