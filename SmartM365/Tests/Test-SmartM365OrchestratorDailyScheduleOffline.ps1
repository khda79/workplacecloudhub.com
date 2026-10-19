#Requires -Version 7.0
<#
.SYNOPSIS
Validate the single daily inventory cycle and weekly exceptions offline.
.VERSION
1.1.0
.NOTES
Reads the repository template and an optional additional local manifest.
No tenant initialization, collectors, scheduling, authentication or publication.
#>
[CmdletBinding()]
param([string]$AdditionalManifestPath = '')
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
$orchestratorRoot = Join-Path $smartRoot 'SmartInventory/Orchestrator'
Import-Module (Join-Path $orchestratorRoot 'SmartM365.Orchestrator.Management.psm1') -Force
$checks = 0
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}
$daily = @{
        'M365-VerifiedDomains-Inventory' = '00:05'
        'AD-Inventory' = '00:10'
        'AD-HealthCheck' = '05:00'
        'M365-ActiveUsers-Inventory' = '00:15'
        'M365-Licences-Inventory' = '00:15'
        'M365-EntraDevices-Inventory' = '00:25'
        'M365-Teams-Inventory' = '03:30'
        'M365-SPO-Inventory' = '03:45'
        'Intune-Devices-Inventory' = '00:35'
        'Intune-Devices-Compliance-Inventory' = '01:30'
        'Intune-WindowsAutopilot-Inventory' = '04:00'
        'EXO-Mailboxes-Permissions' = '00:15'
        'EXO-Mailboxes-Inventory-Fast' = '00:15'
        'EXO-AcceptedDomains-Inventory' = '00:10'
        'Mailboxes-PermissionsByUser-Report' = '08:30'
        'Exchange2016-Local-Mailboxes-Inventory' = '00:35'
        'Exchange2016-ProxyAddresses-Check' = '03:15'
        'Exchange2016-Infrastructure-Inventory' = '00:00'
        'M365-Usage-Inventory' = '04:45'
        'Intune-DeviceSystem-Inventory' = '01:15'
        'Intune-Devices-BIOS-Inventory' = '00:55'
        'Intune-Devices-UpgradeEligibility' = '01:45'
        'Intune-RBAC-GroupMembers' = '04:15'
        'Intune-Remediations-Export' = '04:30'
        'Intune-AutopatchAlerts-Inventory' = '03:10'
        'Intune-Windows11-Readiness-Issues' = '07:00'
        'Intune-WinUpdate-Status' = '02:00'
        'Exchange-HybridIdentity-Issues' = '01:00'
        'M365-CopilotUsage-Inventory' = '05:00'
        'M365-TeamsPhonePstnUsage-Inventory' = '05:15'
        'Intune-EndpointAnalytics-Inventory' = '02:45'
        'EXO-QuarantineMessages-Report' = '08:15'
        'M365-SecureScore-Inventory' = '02:00'
        'M365-AuthenticationMethodsRegistration-Inventory' = '02:10'
        'M365-ConditionalAccess-Inventory' = '02:20'
        'Intune-EndpointSecurityHealth-Inventory' = '02:30'
        'M365-LicensePricing-Inventory' = '05:30'
        'SmartFinOps-Workplace-Valuation' = '05:45'
        'SmartFinOps-Workplace-Report' = '05:45'
        'WorkplaceEvidence-Prepare' = '07:45'
        'M365-WorkplaceScope-Inventory' = '18:00'
        'Intune-Configurations-Inventory' = '04:30'
        'Intune-ManagedApps-Inventory' = '05:00'
        'Intune-EnrollmentFailures-Inventory' = '05:30'
        'M365-SyncHealth-Inventory' = '00:20,02:20,04:20,06:20,08:20,10:20,12:20,14:20,16:20,18:20,20:20,22:20'
        'EXO-MigrationJobs-Inventory' = '06:00,12:00,18:00'
}
$weekly = @{
        'EXO-Mailboxes-Inventory' = 'Friday|20:00'
        'EXO-Mailboxes-CalendarPermissions' = 'Tuesday,Thursday|20:00'
        'Exchange2016-Mailboxes-CalendarPermissions' = 'Monday,Wednesday,Friday|04:00'
        'Intune-DiscoveredApps-Inventory' = 'Saturday|20:00'
}
$estimates = @{
        'M365-Licences-Inventory' = 90
        'AD-Inventory' = 120
        'Exchange2016-Local-Mailboxes-Inventory' = 150
        'M365-WorkplaceScope-Inventory' = 360
        'Intune-Devices-Compliance-Inventory' = 15
        'Intune-Configurations-Inventory' = 30
        'Intune-ManagedApps-Inventory' = 60
        'Intune-EnrollmentFailures-Inventory' = 10
        'Intune-DiscoveredApps-Inventory' = 1500
        'EXO-Mailboxes-Inventory-Fast' = 90
        'EXO-Mailboxes-Inventory' = 1200
        'EXO-Mailboxes-Permissions' = 150
        'EXO-Mailboxes-CalendarPermissions' = 480
}
# Template defaults for a new installation; a runtime manifest may enable them explicitly.
$disabledNames = @('M365-BackupProtectedMailboxes-Inventory',
    'M365-BackupPolicyScope-Inventory', 'M365-BackupProtectedSitesAndDrives-Inventory',
    'M365-PowerBIFabricActivity-Inventory')
# LatestOccurrence chain: Name = dependencies | DependencyWaitTimeoutMinutes.
$chain = @{
        'EXO-Mailboxes-Inventory-Fast' = 'AD-Inventory,M365-ActiveUsers-Inventory|1080'
        'EXO-Mailboxes-Permissions' = 'EXO-Mailboxes-Inventory-Fast|1380'
        'M365-Licences-Inventory' = 'EXO-Mailboxes-Inventory-Fast|1380'
        'SmartFinOps-Workplace-Valuation' = 'M365-Licences-Inventory,M365-LicensePricing-Inventory,AD-Inventory,M365-ActiveUsers-Inventory|1080'
        'SmartFinOps-Workplace-Report' = 'SmartFinOps-Workplace-Valuation|1110'
}
$templatePath = Join-Path $orchestratorRoot 'Orchestrator-Jobs.json.template'
$paths = @($templatePath)
if ($AdditionalManifestPath) { $paths += $AdditionalManifestPath }
foreach ($path in $paths) {
    $isTemplate = ($path -eq $templatePath)
    $newIntuneNames = @('Intune-Configurations-Inventory','Intune-ManagedApps-Inventory','Intune-EnrollmentFailures-Inventory')
    $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $document = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    $validation = Test-SmartM365OrchestratorJobsDocument -Document $document
    Assert-True $validation.Valid ('Invalid jobs manifest: ' + ($validation.Errors -join '; '))
    if ($isTemplate) { Assert-True ($document.Jobs.Count -eq 54) 'Expected 54 template jobs.' }
    Assert-True (@($document.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare').Count -eq 0) 'Retired CMDB preparation job is still present.'
    foreach ($name in @($daily.Keys | Where-Object { $isTemplate -or $_ -notin $newIntuneNames })) {
        $matches = @($document.Jobs | Where-Object Name -eq $name)
        Assert-True ($matches.Count -eq 1) "Missing or duplicate daily job: $name"
        $schedule = $matches[0].Schedule
        Assert-True ($schedule.Type -eq 'Daily') "Non-daily schedule: $name"
        Assert-True ((@($schedule.Times) -join ',') -ceq $daily[$name]) "Wrong daily occurrences: $name"
        Assert-True (@($schedule.DaysOfWeek).Count -eq 0) "Daily job retains weekly day restrictions: $name"
    }
    foreach ($name in $weekly.Keys) {
        $matches = @($document.Jobs | Where-Object Name -eq $name)
        Assert-True ($matches.Count -eq 1) "Missing or duplicate weekly job: $name"
        $schedule = $matches[0].Schedule
        Assert-True ($schedule.Type -eq 'Weekly') "Non-weekly schedule: $name"
        $actual = (@($schedule.DaysOfWeek) -join ',') + '|' + (@($schedule.Times) -join ',')
        Assert-True ($actual -ceq $weekly[$name]) "Wrong weekly occurrences: $name"
    }
    if ($isTemplate) {
        foreach ($name in $disabledNames) {
            $matches = @($document.Jobs | Where-Object Name -eq $name)
            Assert-True ($matches.Count -eq 1 -and -not $matches[0].Enabled) "Disabled template job was activated: $name"
        }
    }
    $templateOnlyEstimates = @($newIntuneNames) + 'M365-WorkplaceScope-Inventory' + 'Intune-Devices-Compliance-Inventory'
    foreach ($name in @($estimates.Keys | Where-Object { $isTemplate -or $_ -notin $templateOnlyEstimates })) {
        $job = @($document.Jobs | Where-Object Name -eq $name)[0]
        Assert-True ($job.EstimatedDurationMinutes -eq $estimates[$name]) "Unrealistic fallback estimate: $name"
    }
    $readiness = @($document.Jobs | Where-Object Name -eq 'Intune-Windows11-Readiness-Issues')[0]
    Assert-True ($readiness.DependencyMode -eq 'FreshSuccess' -and $readiness.DependencyMaxAgeHours -eq 48) 'Readiness lost its fresh-success dependency gate.'
    Assert-True ($readiness.Schedule.MissedRunPolicy -eq 'RunOnce') 'Missed readiness analysis is silently skipped.'
    $hybrid = @($document.Jobs | Where-Object Name -eq 'Exchange-HybridIdentity-Issues')[0]
    Assert-True ($hybrid.DependencyMode -eq 'LatestOccurrence' -and $hybrid.DependencyWaitTimeoutMinutes -eq 1380) 'Hybrid identity analysis no longer waits for the same night outputs.'
    foreach ($source in @('AD-Inventory','EXO-Mailboxes-Inventory-Fast','EXO-Mailboxes-Permissions','M365-Licences-Inventory','Exchange2016-Local-Mailboxes-Inventory')) {
        Assert-True (@($hybrid.DependsOn) -contains $source) "Hybrid identity analysis lost dependency: $source"
    }
    Assert-True ($hybrid.Schedule.MissedRunPolicy -eq 'RunOnce') 'Missed hybrid identity analysis is silently skipped.'
    foreach ($name in $chain.Keys) {
        $job = @($document.Jobs | Where-Object Name -eq $name)[0]
        $expected = $chain[$name].Split('|')
        Assert-True ((@($job.DependsOn) -join ',') -ceq $expected[0]) "Chained dependencies changed: $name"
        $mode = if ($job.PSObject.Properties['DependencyMode']) { [string]$job.DependencyMode } else { '' }
        Assert-True ($mode -in @('', 'LatestOccurrence')) "Chained job no longer uses LatestOccurrence: $name"
        Assert-True ([int]$job.DependencyWaitTimeoutMinutes -eq [int]$expected[1]) "Chained dependency wait changed: $name"
    }
    $valuation = @($document.Jobs | Where-Object Name -eq 'SmartFinOps-Workplace-Valuation')[0]
    $report = @($document.Jobs | Where-Object Name -eq 'SmartFinOps-Workplace-Report')[0]
    Assert-True ($valuation.ConcurrencyKey -eq 'SmartFinOps-Workplace' -and $report.ConcurrencyKey -eq $valuation.ConcurrencyKey) 'SmartFinOps jobs can run concurrently.'
    $prepare = @($document.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')[0]
    Assert-True ($prepare.DependencyMode -eq 'FreshSuccess' -and $prepare.DependencyMaxAgeHours -eq 48) 'Workplace preparation scheduler gate changed.'
    Assert-True (@($prepare.DependsOn) -contains 'EXO-Mailboxes-Inventory-Fast' -and
        @($prepare.DependsOn) -notcontains 'EXO-Mailboxes-Inventory' -and
        @($prepare.DependsOn) -contains 'M365-WorkplaceScope-Inventory' -and
        @($prepare.DependsOn) -contains 'Intune-DiscoveredApps-Inventory') 'Workplace preparation depends on weekly stats or lost a former CMDB producer.'
    $apps = @($document.Jobs | Where-Object Name -eq 'Intune-DiscoveredApps-Inventory')[0]
    Assert-True ([string]$apps.Arguments -ceq '-DeviceDetailMode All -FreshDeviceDetails') 'Weekly application acquisition is no longer complete and fresh.'
    $full = @($document.Jobs | Where-Object Name -eq 'EXO-Mailboxes-Inventory')[0]
    $fast = @($document.Jobs | Where-Object Name -eq 'EXO-Mailboxes-Inventory-Fast')[0]
    Assert-True ($full.ConcurrencyKey -eq 'EXOMailboxes' -and $fast.ConcurrencyKey -eq $full.ConcurrencyKey) 'Concurrent full and fast EXO acquisition can launch.'
    Assert-True ([string]::IsNullOrWhiteSpace([string]$fast.Arguments)) 'Daily EXO Fast is restricted or collects live stats.'
    $ad = @($document.Jobs | Where-Object Name -eq 'AD-Inventory')[0]
    $scope = @($document.Jobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')[0]
    Assert-True ($ad.ConcurrencyKey -eq 'AD-WorkplaceScope' -and $scope.ConcurrencyKey -eq $ad.ConcurrencyKey) 'AD and WorkplaceScope collections can run at the same time.'
    Assert-True ($ad.TimeoutMinutes -eq 240 -and $scope.TimeoutMinutes -eq 480) 'AD or WorkplaceScope timeout changed.'
    $sync = @($document.Jobs | Where-Object Name -eq 'M365-SyncHealth-Inventory')[0]
    Assert-True ($sync.Schedule.MissedRunPolicy -eq 'Skip') 'SyncHealth can accumulate catch-up monitoring runs.'
    if ($isTemplate) {
        foreach ($name in @('M365-WorkplaceScope-Inventory','Intune-Configurations-Inventory','Intune-ManagedApps-Inventory','Intune-EnrollmentFailures-Inventory','WorkplaceEvidence-Prepare')) {
            $job = @($document.Jobs | Where-Object Name -eq $name)[0]
            Assert-True ($job.Enabled -and -not $job.PSObject.Properties['RequiresExplicitActivation'] -and
                $job.AssignmentMode -eq 'Elected') "Template job activation differs: $name"
        }
        # The public template never names a server: no pinned job and no AllowedServers entry.
        foreach ($job in $document.Jobs) {
            Assert-True ([string]$job.AssignmentMode -ne 'Pinned') "Template job is pinned to a server: $($job.Name)"
            Assert-True ($null -ne $job.PSObject.Properties['AllowedServers'] -and @($job.AllowedServers).Count -eq 0) "Template job lists servers: $($job.Name)"
        }
    }
    Assert-True ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ceq $hash) 'Offline verification altered the manifest.'
}
$contract = Get-Content -LiteralPath (Join-Path $smartRoot 'SmartInventory/PreparedEvidence/cmdb-prepared-contract.json.txt') -Raw | ConvertFrom-Json
Assert-True ($contract.maxAgeHours -eq 48 -and $contract.maxCollectionSpanHours -eq 48) 'Core CMDB acquisition limits weakened.'
Assert-True ($contract.freshnessGroups[0].warningAgeHours -eq 168 -and $contract.freshnessGroups[0].maxAgeHours -eq 240) 'Weekly app acquisition limits changed.'
Write-Output "PASS: $checks daily schedule offline checks. No collectors, scheduling or live publication."
