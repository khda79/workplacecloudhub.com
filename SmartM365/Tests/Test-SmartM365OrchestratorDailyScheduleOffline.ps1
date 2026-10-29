#Requires -Version 7.0
<#
.SYNOPSIS
Validate the single daily inventory cycle and weekly exceptions offline.
.VERSION
1.2.11
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
        'AD-Adaxes-Inventory' = '01:30'
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
    'M365-PowerBIFabricActivity-Inventory', 'AD-Adaxes-Inventory')
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
    if ($isTemplate) { Assert-True ($document.Jobs.Count -eq 56) 'Expected 56 template jobs.' }
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
        $expectedTime=if($isTemplate -and $name -eq 'M365-Licences-Inventory'){'04:45'}else{$daily[$name]}
        Assert-True ((@($schedule.Times) -join ',') -ceq $expectedTime) "Wrong daily occurrences: $name"
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
        if ($isTemplate -and $name -eq 'M365-Licences-Inventory') { $expected[0]='EXO-Mailboxes-Inventory-Fast,M365-Usage-Inventory,Intune-Devices-Inventory' }
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
    if ($isTemplate) { Assert-True ((@($prepare.LatestOccurrenceDependencies) -join ',') -ceq 'M365-Licences-Inventory') 'Prepare must wait for the latest licensing occurrence.' }
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA2Qwvt+/d/zrKa
# +CxetPxAmVrfeh/mhM00Z/OpSMZ186CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB83NsNQYaeoYfo+vCjmRvUYhKvPdzizHT9riLBMbtK2MA0GCSqG
# SIb3DQEBAQUABIIBgCNd3PWp1fzAAsDpaS+5ep4+/ikPt9Q6Zj40HDK3EL24/AgM
# OmhXzDEKqh4ApiX1f+OFAhDfMKeYRwXuROY1D8ujUZEV/A2PL2g3p70tWkNuUKsg
# VjApAS9Ko/PlaxxSCooqBsSBJSzcTkd8pLnSX6Qyzeqa70InLPI3YGDxjST3v6nn
# DMZ0Q1sMt1kvvRYDWPGg2r10FNDRKlQLqNkcZNezOH0UgzRhLe8aKi0jcLxj6oMH
# yTT6WiibuGwxhP5MbhQ2MbMIbZiM1e/eV6/VTpsSdTMpwEG6FaYa1g/WGNQhco6y
# +vI2fiHbbagib5OATDhOslim/p+2FQmhnmh32Hbo4QuOd5aeD4hZM6n8ujCuGuZg
# J/KcsSVRlbTh/tHqqPZD5fPdBrHBVbs0IVu8vs+OHQf9Rf4Fgy/NZA/fA8aJdsWy
# tbf253GFmN567M9oGmtc9PinUsROzIzi02bjGB4jgiDd7Ec27Hmumc4XPG1V8Fht
# YqZiC8Q9Sws5RbpftKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcyMTEx
# MDRaMC8GCSqGSIb3DQEJBDEiBCAlzxwh8829aijDteGGqyNZ9lgfc6W0orfQtcmq
# 3CERIDANBgkqhkiG9w0BAQEFAASCAgB0dTnkIXdIFqz43NqVHD82gur2w/FVUys5
# NTauI6WBLBaenDPPKQQkZllYF+lACR1s+oG8lWyEu2LKw3go3r88ZfdigqbsXnvL
# yvwU9PwRS/Vh7pS6AruznhD6ziFxbkoTa6QtbdgzShfDJPXyGHh8lWlewmhq8Nkf
# xK1H9m7QZDnHqMq9nCR0QdlUicbA1klvJijJnMZyJhVKDD4Lo9k0v7UL9I8Aaw3Q
# 7ykTu61ceRXg2XWTpVGvBpOraiyY60UwX6m6hS/VrpkWIpu2ZlUqXBFRSI1/2yB/
# Ya+PbT7LMIfVGJOiRo7OoNahfYIr0U4r0LvjcnBWcPCITifXbF1xJmNyxE9eN/2E
# fpLnjE1aj+YusyJ8VQBYJW9kRv12H6COD3VSQdmj9zW4y8prf+5T5KAddMDOiuwC
# Q2f2YATvSpWWgna68Wb8FQj8sew5dQvepZYAb3hEkGS0VyidMZqtUNSYB5dqMr7l
# PLNvCcN/JVrsTxTHUQnoObkyvElUzwp9KyDgumOwXTFrHY9gYstS9mkv5UPAfRE9
# g9JdK/NNNC8kfSTbcMTsDvBbh1ZppvQaf7LpM5U31cwmgmlXZdYlYcjI9e3dG9xN
# +7FRKJuRfTFYS8xfN+X0nSWaklbV4sA7McMe2OWt21RMQexRECFv+qM8U3E/fFlS
# D0+X4/OysQ==
# SIG # End signature block
