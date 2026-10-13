#Requires -Version 7.0
<#
.SYNOPSIS
Validate the single daily inventory cycle and weekly exceptions offline.
.VERSION
1.0.0
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
        'M365-Licences-Inventory' = '00:45'
        'M365-EntraDevices-Inventory' = '00:25'
        'M365-Teams-Inventory' = '03:30'
        'M365-SPO-Inventory' = '03:45'
        'Intune-Devices-Inventory' = '00:35'
        'Intune-Devices-Compliance-Inventory' = '01:30'
        'Intune-WindowsAutopilot-Inventory' = '04:00'
        'EXO-Mailboxes-Permissions' = '05:00'
        'EXO-Mailboxes-Inventory-Fast' = '03:15'
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
        'Exchange-HybridIdentity-Issues' = '08:00'
        'M365-CopilotUsage-Inventory' = '05:00'
        'M365-TeamsPhonePstnUsage-Inventory' = '05:15'
        'Intune-EndpointAnalytics-Inventory' = '02:45'
        'EXO-QuarantineMessages-Report' = '08:15'
        'M365-SecureScore-Inventory' = '02:00'
        'M365-AuthenticationMethodsRegistration-Inventory' = '02:10'
        'M365-ConditionalAccess-Inventory' = '02:20'
        'Intune-EndpointSecurityHealth-Inventory' = '02:30'
        'M365-LicensePricing-Inventory' = '05:30'
        'WorkplaceEvidence-Prepare' = '09:00'
        'M365-WorkplaceScope-Inventory' = '22:00'
        'Intune-Configurations-Inventory' = '04:30'
        'Intune-ManagedApps-Inventory' = '05:00'
        'Intune-EnrollmentFailures-Inventory' = '05:30'
        'CmdbEvidence-Prepare' = '09:30'
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
        'Intune-Configurations-Inventory' = 30
        'Intune-ManagedApps-Inventory' = 60
        'Intune-EnrollmentFailures-Inventory' = 10
        'Intune-DiscoveredApps-Inventory' = 1500
        'EXO-Mailboxes-Inventory-Fast' = 90
        'EXO-Mailboxes-Inventory' = 1200
        'EXO-Mailboxes-Permissions' = 150
        'EXO-Mailboxes-CalendarPermissions' = 480
}
$disabledNames = @('M365-BackupProtectedMailboxes-Inventory',
    'M365-BackupPolicyScope-Inventory', 'M365-BackupProtectedSitesAndDrives-Inventory',
    'M365-PowerBIFabricActivity-Inventory')
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
    if ($isTemplate) { Assert-True ($document.Jobs.Count -eq 53) 'Expected 53 template jobs.' }
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
    foreach ($name in $disabledNames) {
        $matches = @($document.Jobs | Where-Object Name -eq $name)
        Assert-True ($matches.Count -eq 1 -and -not $matches[0].Enabled) "Disabled job was activated: $name"
    }
    foreach ($name in @($estimates.Keys | Where-Object { $isTemplate -or $_ -notin (@($newIntuneNames) + 'M365-WorkplaceScope-Inventory') })) {
        $job = @($document.Jobs | Where-Object Name -eq $name)[0]
        Assert-True ($job.EstimatedDurationMinutes -eq $estimates[$name]) "Unrealistic fallback estimate: $name"
    }
    foreach ($name in @('Intune-Windows11-Readiness-Issues','Exchange-HybridIdentity-Issues')) {
        $job = @($document.Jobs | Where-Object Name -eq $name)[0]
        Assert-True ($job.DependencyMode -eq 'FreshSuccess' -and $job.DependencyMaxAgeHours -eq 48) "Analysis lost its fresh-success dependency gate: $name"
        Assert-True ($job.Schedule.MissedRunPolicy -eq 'RunOnce') "Missed daily analysis is silently skipped: $name"
    }
    $prepare = @($document.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare')[0]
    Assert-True ($prepare.Arguments -ceq '-Publish') 'CMDB job no longer publishes the verified prepared cohort.'
    Assert-True ($prepare.DependsOn.Count -eq 17 -and
        $prepare.DependsOn -contains 'EXO-Mailboxes-Inventory-Fast' -and
        $prepare.DependsOn -notcontains 'EXO-Mailboxes-Inventory' -and
        $prepare.DependsOn -notcontains 'EXO-Mailboxes-Permissions') 'CMDB acquisition depends on weekly stats or permissions-only.'
    Assert-True ($prepare.DependencyMode -eq 'FreshSuccess' -and $prepare.DependencyMaxAgeHours -eq 240) 'Weekly Apps scheduler gate changed.'
    $full = @($document.Jobs | Where-Object Name -eq 'EXO-Mailboxes-Inventory')[0]
    $fast = @($document.Jobs | Where-Object Name -eq 'EXO-Mailboxes-Inventory-Fast')[0]
    Assert-True ($full.ConcurrencyKey -eq 'EXOMailboxes' -and $fast.ConcurrencyKey -eq $full.ConcurrencyKey) 'Concurrent full and fast EXO acquisition can launch.'
    Assert-True ([string]::IsNullOrWhiteSpace([string]$fast.Arguments)) 'Daily EXO Fast is restricted or collects live stats.'
    $sync = @($document.Jobs | Where-Object Name -eq 'M365-SyncHealth-Inventory')[0]
    Assert-True ($sync.Schedule.MissedRunPolicy -eq 'Skip') 'SyncHealth can accumulate catch-up monitoring runs.'
    if ($isTemplate) {
        foreach ($name in @('M365-WorkplaceScope-Inventory','Intune-Configurations-Inventory','Intune-ManagedApps-Inventory','Intune-EnrollmentFailures-Inventory','CmdbEvidence-Prepare')) {
            $job = @($document.Jobs | Where-Object Name -eq $name)[0]
            Assert-True ($job.Enabled -and -not $job.PSObject.Properties['RequiresExplicitActivation'] -and
                $job.AssignmentMode -eq 'Elected' -and @($job.AllowedServers).Count -eq 0) "Template job activation differs: $name"
        }
    }
    Assert-True ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ceq $hash) 'Offline verification altered the manifest.'
}
$contract = Get-Content -LiteralPath (Join-Path $smartRoot 'SmartInventory/PreparedEvidence/cmdb-prepared-contract.json.txt') -Raw | ConvertFrom-Json
Assert-True ($contract.maxAgeHours -eq 48 -and $contract.maxCollectionSpanHours -eq 48) 'Core CMDB acquisition limits weakened.'
Assert-True ($contract.freshnessGroups[0].warningAgeHours -eq 168 -and $contract.freshnessGroups[0].maxAgeHours -eq 240) 'Weekly app acquisition limits changed.'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $smartRoot 'SmartInventory/PreparedEvidence/SmartM365-CmdbEvidence-Orchestrator.ps1'),
    [ref]$tokens, [ref]$parseErrors)
Assert-True ($parseErrors.Count -eq 0) 'CMDB activation entry point does not parse.'
foreach ($definition in @(@('PreparationTime','09:30'), @('ScopeTime','22:00'))) {
    $parameter = @($ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq $definition[0] })[0]
    Assert-True ($parameter.DefaultValue.SafeGetValue() -ceq $definition[1]) 'Explicit activation defaults differ from the daily plan.'
}
Write-Output "PASS: $checks daily schedule offline checks. No collectors, scheduling or live publication."


# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASKOK7AhceW0Ms
# BVQPuMFcAbAQ32FgdBycNfNvGXv0YqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA0MYUN2Htbb8akWKGCEFZj
# rjMvoOTRZf2FmlHJD5QVSDANBgkqhkiG9w0BAQEFAASCAYB2m/k1DGvGXbmX867a
# /U7dTiorO6OEHM9RD4Q4QqWADOdD/5vLJaJ0+LV7vhs9e0z97dEnUSHneplhave/
# e2NEx2lxgwrn9e+JPCrAVQBfpJQjEnO5fTu5eSfp9BjFG/XoNJBuSEbSmTcPJXTk
# PzFuSGgc1tkkx9JL4mLTKLK0zNPaI/FiVyBZP9EDhiiWnvh/N/AyRnCxyadP5KPu
# IXLMEqeUVaOyyS+J3zR99CfpM5kHmSjOCvdYdGiK4RS3+Cgh57gq1zXaMJjdv/jN
# ObhRR9efQyrywC8xlI1dAPX2AIgSIJ1zhlsu6Hna+whwOdwdM7QUN7CmxgzWZ28Q
# P/22nT4FGqtQ+yhBtW2R217V7n2ul/BSNiQYXy1qzrbMinyFyAXTPbDeWiNQInar
# NfEnaTcua1epGD5tEpHY+YOk1aCwunDfyMBFDL2GkvkD5JvRjGf409pw0hmetur1
# r3tSbmvA5Jwz9Aqf/5lShR/eAwQibL/Vhve7ipwrqwTC4ko=
# SIG # End signature block
