#Requires -Version 7.0
<#
.SYNOPSIS
Validate the single daily inventory cycle and weekly exceptions offline.
.VERSION
1.2.9
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
        'Exchange2016-ProxyAddresses-Check' = '00:00'
        'Exchange2016-Infrastructure-Inventory' = '03:15'
        'M365-Usage-Inventory' = '04:45'
        'Intune-DeviceSystem-Inventory' = '01:15'
        'Intune-Devices-BIOS-Inventory' = '00:55'
        'Intune-Devices-UpgradeEligibility' = '01:45'
        'Intune-RBAC-GroupMembers' = '04:15'
        'Intune-Remediations-Export' = '04:30'
        'Intune-AutopatchAlerts-Inventory' = '03:10'
        'Intune-WinUpdate-Status' = '02:00'
        'M365-CopilotUsage-Inventory' = '05:00'
        'M365-TeamsPhonePstnUsage-Inventory' = '05:15'
        'Intune-EndpointAnalytics-Inventory' = '02:45'
        'EXO-QuarantineMessages-Report' = '08:15'
        'M365-SecureScore-Inventory' = '02:00'
        'M365-AuthenticationMethodsRegistration-Inventory' = '02:10'
        'M365-ConditionalAccess-Inventory' = '02:20'
        'Intune-EndpointSecurityHealth-Inventory' = '04:00'
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
        'EXO-Mailboxes-Inventory' = 'Saturday|06:00'
        'EXO-Mailboxes-CalendarPermissions' = 'Tuesday,Thursday|20:00'
        'Exchange2016-Mailboxes-CalendarPermissions' = 'Monday,Wednesday,Friday|04:00'
        'Intune-DiscoveredApps-Inventory' = 'Saturday|20:00'
        'Exchange2016-FreeBusy-Check' = 'Sunday|05:00'
}
$estimates = @{
        'M365-Licences-Inventory' = 90
        'AD-Inventory' = 120
        'Exchange2016-Local-Mailboxes-Inventory' = 150
        'Exchange2016-ProxyAddresses-Check' = 30
        'Exchange2016-Infrastructure-Inventory' = 90
        'Exchange2016-FreeBusy-Check' = 15
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
        # The weekly full run starts only after the same Saturday's Fast, so Fast never queues behind it.
        'EXO-Mailboxes-Inventory' = 'AD-Inventory,M365-ActiveUsers-Inventory,EXO-Mailboxes-Inventory-Fast|1080'
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
    if ($isTemplate) { Assert-True ($document.Jobs.Count -eq 55) 'Expected 55 template jobs.' }
    if ($isTemplate) {
        # SharePoint Server collectors are manual for now: -Force only, no dependency, no scheduled election.
        foreach ($name in @('SharePointOnPrem-Infrastructure-Inventory', 'SharePointOnPrem-Content-Inventory')) {
            $sharePointJob = @($document.Jobs | Where-Object Name -eq $name)
            Assert-True ($sharePointJob.Count -eq 1 -and $sharePointJob[0].AssignmentMode -eq 'Manual' -and @($sharePointJob[0].DependsOn).Count -eq 0 -and
                $sharePointJob[0].PowerShellEdition -eq 'WindowsPowerShell' -and (@($sharePointJob[0].RequiredCapabilities) -join ',') -ceq 'SharePointOnPrem,SharedRuntime' -and
                -not $sharePointJob[0].PSObject.Properties['ConcurrencyKey']) "SharePoint on-premises job is not manual and independent: $name"
            Assert-True (Test-Path -LiteralPath (Join-Path (Split-Path $orchestratorRoot -Parent) $sharePointJob[0].ScriptPath) -PathType Leaf) "SharePoint on-premises script not found: $name"
            Assert-True (@($document.Jobs | Where-Object { @($_.DependsOn) -contains $name }).Count -eq 0) "A job depends on the manual SharePoint job: $name"
        }
    }
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
    foreach ($retired in @('Exchange-HybridIdentity-Issues','Intune-Windows11-Readiness-Issues')) {
        Assert-True (@($document.Jobs | Where-Object Name -eq $retired).Count -eq 0) "Retired producer remains active: $retired"
        Assert-True (@($document.Jobs | Where-Object { @($_.DependsOn) -contains $retired }).Count -eq 0) "Dependency on retired producer: $retired"
    }
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
    # Every enabled producer of a CSV that preparation reads is a dependency, except the weekly full EXO
    # run (Saturday delay), the optional health check and the report that runs after preparation.
    foreach ($producer in @('Intune-DeviceSystem-Inventory','EXO-MigrationJobs-Inventory','Intune-Configurations-Inventory','Intune-ManagedApps-Inventory','Intune-EnrollmentFailures-Inventory')) {
        Assert-True (@($prepare.DependsOn) -contains $producer) "Workplace preparation lost source producer dependency: $producer"
    }
    foreach ($excluded in @('AD-HealthCheck','Mailboxes-PermissionsByUser-Report')) {
        Assert-True (@($prepare.DependsOn) -notcontains $excluded) "Workplace preparation waits for a non-blocking source: $excluded"
    }
    Assert-True (@($prepare.DependsOn | Sort-Object -Unique).Count -eq @($prepare.DependsOn).Count) 'Workplace preparation lists a dependency twice.'
    $apps = @($document.Jobs | Where-Object Name -eq 'Intune-DiscoveredApps-Inventory')[0]
    Assert-True ([string]$apps.Arguments -ceq '-DeviceDetailMode All -FreshDeviceDetails') 'Weekly application acquisition is no longer complete and fresh.'
    $full = @($document.Jobs | Where-Object Name -eq 'EXO-Mailboxes-Inventory')[0]
    $fast = @($document.Jobs | Where-Object Name -eq 'EXO-Mailboxes-Inventory-Fast')[0]
    Assert-True ($full.ConcurrencyKey -eq 'EXOMailboxes' -and $fast.ConcurrencyKey -eq $full.ConcurrencyKey) 'Concurrent full and fast EXO acquisition can launch.'
    Assert-True ([string]::IsNullOrWhiteSpace([string]$fast.Arguments)) 'Daily EXO Fast is restricted or collects live stats.'
    # The prepared artifact contract expects the service health, mailbox database path and Health Checker CSVs.
    $exchangeInfrastructure = @($document.Jobs | Where-Object Name -eq 'Exchange2016-Infrastructure-Inventory')[0]
    $infrastructureSwitches = @(([string]$exchangeInfrastructure.Arguments) -split '\s+' | Where-Object { $_ })
    foreach ($switch in @('-IncludeServicesHealth', '-IncludeMailboxDatabasePaths', '-IncludeHealthChecker')) {
        Assert-True ($infrastructureSwitches -contains $switch) "Exchange infrastructure no longer collects the prepared artifact enabled by $switch."
    }
    # Health Checker's default total budget (1320 minutes) plus the rest of the inventory must fit the job timeout.
    Assert-True ([int]$exchangeInfrastructure.TimeoutMinutes -ge 1380) 'Exchange infrastructure timeout no longer covers the Health Checker budget.'
    # Same concurrency key as the mailbox chain: it starts after the chain's estimated end.
    $chainEnd = [TimeSpan]::Parse('00:35').Add([TimeSpan]::FromMinutes(150))
    Assert-True ([TimeSpan]::Parse([string]@($exchangeInfrastructure.Schedule.Times)[0]) -ge $chainEnd) 'Exchange infrastructure (Health Checker) starts before the on-premises mailbox chain is expected to end.'
    # Free/busy runs on an Exchange server under Windows PowerShell 5.1; Exchange Online is opened by the
    # collector with the app certificate, so the pwsh Exchange Online probe is not a prerequisite.
    $freeBusy = @($document.Jobs | Where-Object Name -eq 'Exchange2016-FreeBusy-Check')[0]
    Assert-True ($freeBusy.PowerShellEdition -eq 'WindowsPowerShell' -and $freeBusy.ConcurrencyKey -eq 'ExchangeOnPremServer' -and
        (@($freeBusy.RequiredCapabilities) -join ',') -ceq 'AD,ExchangeOnPrem,SharedRuntime') 'Free/busy check requirements changed.'
    Assert-True ([string]::IsNullOrWhiteSpace([string]$freeBusy.Arguments) -and $freeBusy.LauncherPath -eq 'Launchers\OnPremises\Start-SmartM365-Exchange-Hybrid-FreeBusy-Inventory.cmd') 'Free/busy check launcher changed.'
    # The collector's default Free/Busy Checker limit (30 minutes) plus publication must fit the job timeout.
    Assert-True ([int]$freeBusy.TimeoutMinutes -ge 60 -and $freeBusy.Schedule.MissedRunPolicy -eq 'RunOnce') 'Free/busy check timeout or missed-run policy changed.'
    # Endpoint security exports get the same 1200-second budget as the manual launcher.
    $endpointSecurity = @($document.Jobs | Where-Object Name -eq 'Intune-EndpointSecurityHealth-Inventory')[0]
    $endpointSecurityArguments = @(([string]$endpointSecurity.Arguments) -split '\s+' | Where-Object { $_ })
    $exportTimeoutIndex = [array]::IndexOf($endpointSecurityArguments, '-ExportTimeoutSeconds')
    Assert-True ($exportTimeoutIndex -ge 0 -and $endpointSecurityArguments[$exportTimeoutIndex + 1] -eq '1200') 'Endpoint security exports are back to the shorter 600-second script default.'
    Assert-True ((3 * 1200 / 60) -lt [int]$endpointSecurity.TimeoutMinutes) 'Endpoint security export waits no longer fit inside the job timeout.'
    # The mailbox summary uses the same night's ProxyAddresses audit, which now runs first.
    $localMailboxes = @($document.Jobs | Where-Object Name -eq 'Exchange2016-Local-Mailboxes-Inventory')[0]
    Assert-True ((@($localMailboxes.DependsOn) -join ',') -ceq 'Exchange2016-ProxyAddresses-Check' -and $localMailboxes.DependencyMode -eq 'FreshSuccess') 'Local mailboxes no longer wait for a fresh ProxyAddresses check.'
    # One daily time: the 1380-minute timeout policy.
    foreach ($name in @('M365-Teams-Inventory', 'M365-SPO-Inventory', 'Intune-WindowsAutopilot-Inventory', 'Exchange2016-ProxyAddresses-Check', 'Exchange2016-Local-Mailboxes-Inventory')) {
        Assert-True (@($document.Jobs | Where-Object Name -eq $name)[0].TimeoutMinutes -eq 1380) "Daily job timeout is not 1380 minutes: $name"
    }
    # The template keeps the normal 48-hour freshness window; a server may widen it temporarily.
    if ($isTemplate) {
        Assert-True (@($document.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')[0].DependencyMaxAgeHours -eq 48) 'Template workplace preparation freshness window is not 48 hours.'
    }
    # Every one-time daily or weekly job follows the timeout policy unless it is a declared exception.
    $timeoutExceptions = Get-SmartM365OrchestratorTimeoutPolicyExceptions
    Assert-True ((@($timeoutExceptions.Keys) -join ',') -ceq 'AD-Inventory,M365-WorkplaceScope-Inventory,EXO-Mailboxes-Inventory,Exchange2016-FreeBusy-Check') 'Timeout policy exceptions changed.'
    foreach ($exceptionName in @($timeoutExceptions.Keys)) {
        Assert-True (@($document.Jobs | Where-Object Name -eq $exceptionName).Count -eq 1) "Timeout policy exception is not a manifest job: $exceptionName"
    }
    foreach ($job in @($document.Jobs | Where-Object { -not $timeoutExceptions.Contains([string]$_.Name) })) {
        if ($job.Schedule.Type -eq 'Daily' -and @($job.Schedule.Times).Count -eq 1) {
            Assert-True ([int]$job.TimeoutMinutes -eq 1380) "Daily-once job timeout is not 1380 minutes: $($job.Name)"
        }
        elseif ($job.Schedule.Type -eq 'Weekly' -and @($job.Schedule.DaysOfWeek).Count -eq 1 -and @($job.Schedule.Times).Count -eq 1) {
            Assert-True ([int]$job.TimeoutMinutes -eq 8640) "Weekly-once job timeout is not 8640 minutes: $($job.Name)"
        }
    }
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
