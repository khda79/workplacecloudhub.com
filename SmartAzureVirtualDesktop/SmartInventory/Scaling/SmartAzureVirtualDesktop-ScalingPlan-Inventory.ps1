<#
.SYNOPSIS
Exports Azure Virtual Desktop autoscale scaling plan inventory.

.DESCRIPTION
Connects with Az PowerShell, enumerates Azure Virtual Desktop host pools and scaling plans,
and exports CSV files for scaling plan configuration, schedules, host pool assignments, and
host pools without autoscale coverage.

The script is read-only. It does not create or update scaling plans.

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
Requires: Az.Accounts
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$TenantId,
    [string[]]$SubscriptionId,
    [string]$OutputRoot,
    [string]$LatestOutputRoot,
    [switch]$Connect,
    [switch]$UseDeviceCode,
    [string]$AvdApiVersion = '2024-04-03',
    [string]$ResourceApiVersion = '2021-04-01'
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    exit 1
}

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path -Path $d -ChildPath 'Config\SmartAzureVirtualDesktop-TenantContext.ps1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartAzureVirtualDesktop-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartAvdEffectiveConfig = Initialize-SmartAvdTenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ScriptLocalConfig = Get-SmartAvdScriptLocalConfig -ScriptPath $PSCommandPath
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    $TenantId = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'AzureTenantId' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') { $TenantId = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '' }
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') { $TenantId = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue '' }
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { $LatestOutputRoot = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue '' }

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartAvdOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'AzureVirtualDesktop\Scaling'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzureVirtualDesktop-ScalingPlan-Inventory_{0}.log" -f $runId)

Import-SmartAvdCoreModule -StartPath $PSScriptRoot
Set-SmartAvdCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-HostPoolArmPathFromReference {
    [CmdletBinding()]
    param([AllowNull()]$Reference)

    $value = Get-SmartAvdObjectPropertyValue -InputObject $Reference -PropertyName @('hostPoolArmPath', 'hostPoolResourceId', 'id')
    if ($null -eq $value) { return '' }
    return [string]$value
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAvdLog -Message "Starting SmartAzureVirtualDesktop scaling plan inventory. RunId=$runId"
    Import-SmartAvdRequiredModule -Name Az.Accounts
    $context = Connect-SmartAvdCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw 'No Azure subscriptions were found for the current identity and filters.' }

    $scalingPlanRows = @()
    $scheduleRows = @()
    $assignmentRows = @()
    $hostPoolCoverageRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAvdLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $hostPools = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.DesktopVirtualization/hostPools' -ApiVersion $ResourceApiVersion)
        $scalingPlans = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.DesktopVirtualization/scalingPlans' -ApiVersion $ResourceApiVersion)
        $assignedHostPoolIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        foreach ($scalingPlan in $scalingPlans) {
            $details = Get-SmartAvdResourceById -ResourceId $scalingPlan.id -ApiVersion $AvdApiVersion -Operation 'Scaling plan details'
            if ($null -eq $details) { $details = $scalingPlan }
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            $hostPoolReferences = @(Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('hostPoolReferences'))
            $scalingPlanName = Get-SmartAvdResourceNameFromId -ResourceId $scalingPlan.id

            $scalingPlanRows += [pscustomobject]@{
                RunId                  = $runId
                SubscriptionId         = $subscription.Id
                SubscriptionName       = $subscription.Name
                ResourceGroupName      = Get-SmartAvdResourceGroupNameFromId -ResourceId $scalingPlan.id
                ScalingPlanName        = $scalingPlanName
                Location               = $details.location
                HostPoolType           = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('hostPoolType')
                ScalingMethod          = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('scalingMethod')
                TimeZone               = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('timeZone')
                ExclusionTag           = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('exclusionTag')
                FriendlyName           = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('friendlyName')
                Description            = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('description')
                HostPoolReferenceCount = $hostPoolReferences.Count
                HostPoolReferencesJson = ConvertTo-SmartAvdCompactJson -Value $hostPoolReferences
                TagsJson               = ConvertTo-SmartAvdCompactJson -Value $details.tags
                ResourceId             = $scalingPlan.id
            }

            foreach ($reference in $hostPoolReferences) {
                $hostPoolArmPath = Get-HostPoolArmPathFromReference -Reference $reference
                if (-not [string]::IsNullOrWhiteSpace($hostPoolArmPath)) { [void]$assignedHostPoolIds.Add($hostPoolArmPath) }
                $assignmentRows += [pscustomobject]@{
                    RunId               = $runId
                    SubscriptionId      = $subscription.Id
                    SubscriptionName    = $subscription.Name
                    ScalingPlanName     = $scalingPlanName
                    HostPoolName        = Get-SmartAvdResourceNameFromId -ResourceId $hostPoolArmPath
                    HostPoolArmPath     = $hostPoolArmPath
                    ScalingPlanEnabled  = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $reference -PropertyName @('scalingPlanEnabled'))
                    AssignmentJson      = ConvertTo-SmartAvdCompactJson -Value $reference
                    ScalingPlanResourceId = $scalingPlan.id
                }
            }

            $pooledSchedules = @(Get-SmartAvdChildResources -ParentResourceId $scalingPlan.id -ChildPath 'pooledSchedules' -ApiVersion $AvdApiVersion -Operation "Pooled schedules for $scalingPlanName")
            foreach ($schedule in $pooledSchedules) {
                $scheduleProperties = Get-SmartAvdObjectPropertyValue -InputObject $schedule -PropertyName @('properties')
                $scheduleRows += [pscustomobject]@{
                    RunId                      = $runId
                    SubscriptionId             = $subscription.Id
                    SubscriptionName           = $subscription.Name
                    ScalingPlanName            = $scalingPlanName
                    ScheduleType               = 'Pooled'
                    ScheduleName               = Get-SmartAvdResourceNameFromId -ResourceId $schedule.id
                    DaysOfWeek                 = (@(Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('daysOfWeek')) -join ';')
                    RampUpStartTime            = "{0:D2}:{1:D2}" -f [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampUpStartTimeHour')), [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampUpStartTimeMinute'))
                    PeakStartTime              = "{0:D2}:{1:D2}" -f [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('peakStartTimeHour')), [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('peakStartTimeMinute'))
                    RampDownStartTime          = "{0:D2}:{1:D2}" -f [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownStartTimeHour')), [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownStartTimeMinute'))
                    OffPeakStartTime           = "{0:D2}:{1:D2}" -f [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('offPeakStartTimeHour')), [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('offPeakStartTimeMinute'))
                    RampUpMinimumHostsPct      = Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampUpMinimumHostsPct')
                    RampUpCapacityThresholdPct = Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampUpCapacityThresholdPct')
                    RampDownMinimumHostsPct    = Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownMinimumHostsPct')
                    RampDownCapacityThresholdPct = Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownCapacityThresholdPct')
                    RampDownForceLogoffUser    = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownForceLogoffUser'))
                    RampDownStopHostsWhen      = Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownStopHostsWhen')
                    ScheduleJson               = ConvertTo-SmartAvdCompactJson -Value $scheduleProperties
                    ResourceId                 = $schedule.id
                }
            }

            $personalSchedules = @(Get-SmartAvdChildResources -ParentResourceId $scalingPlan.id -ChildPath 'personalSchedules' -ApiVersion $AvdApiVersion -Operation "Personal schedules for $scalingPlanName")
            foreach ($schedule in $personalSchedules) {
                $scheduleProperties = Get-SmartAvdObjectPropertyValue -InputObject $schedule -PropertyName @('properties')
                $scheduleRows += [pscustomobject]@{
                    RunId                      = $runId
                    SubscriptionId             = $subscription.Id
                    SubscriptionName           = $subscription.Name
                    ScalingPlanName            = $scalingPlanName
                    ScheduleType               = 'Personal'
                    ScheduleName               = Get-SmartAvdResourceNameFromId -ResourceId $schedule.id
                    DaysOfWeek                 = (@(Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('daysOfWeek')) -join ';')
                    RampUpStartTime            = "{0:D2}:{1:D2}" -f [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampUpStartTimeHour')), [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampUpStartTimeMinute'))
                    PeakStartTime              = "{0:D2}:{1:D2}" -f [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('peakStartTimeHour')), [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('peakStartTimeMinute'))
                    RampDownStartTime          = "{0:D2}:{1:D2}" -f [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownStartTimeHour')), [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownStartTimeMinute'))
                    OffPeakStartTime           = "{0:D2}:{1:D2}" -f [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('offPeakStartTimeHour')), [int](Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('offPeakStartTimeMinute'))
                    RampUpAutoStartHost        = Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampUpAutoStartHost')
                    RampDownActionOnDisconnect = Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownActionOnDisconnect')
                    RampDownActionOnLogoff     = Get-SmartAvdObjectPropertyValue -InputObject $scheduleProperties -PropertyName @('rampDownActionOnLogoff')
                    ScheduleJson               = ConvertTo-SmartAvdCompactJson -Value $scheduleProperties
                    ResourceId                 = $schedule.id
                }
            }
        }

        foreach ($hostPool in $hostPools) {
            $details = Get-SmartAvdResourceById -ResourceId $hostPool.id -ApiVersion $AvdApiVersion -Operation 'Host pool details'
            if ($null -eq $details) { $details = $hostPool }
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            $isAssigned = $assignedHostPoolIds.Contains($hostPool.id)
            $hostPoolCoverageRows += [pscustomobject]@{
                RunId                = $runId
                SubscriptionId       = $subscription.Id
                SubscriptionName     = $subscription.Name
                ResourceGroupName    = Get-SmartAvdResourceGroupNameFromId -ResourceId $hostPool.id
                HostPoolName         = Get-SmartAvdResourceNameFromId -ResourceId $hostPool.id
                HostPoolType         = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('hostPoolType')
                MaxSessionLimit      = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('maxSessionLimit')
                LoadBalancerType     = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('loadBalancerType')
                StartVMOnConnect     = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('startVMOnConnect'))
                HasScalingPlan       = $isAssigned
                CoverageFinding      = if ($isAssigned) { '' } else { 'NoScalingPlanAssignment' }
                ResourceId           = $hostPool.id
            }
        }
    }

    Export-SmartAvdCsv -Name 'AVD_ScalingPlans' -Rows $scalingPlanRows
    Export-SmartAvdCsv -Name 'AVD_ScalingPlanSchedules' -Rows $scheduleRows
    Export-SmartAvdCsv -Name 'AVD_ScalingPlanAssignments' -Rows $assignmentRows
    Export-SmartAvdCsv -Name 'AVD_HostPoolScalingCoverage' -Rows $hostPoolCoverageRows

    Write-SmartAvdLog -Level SUCCESS -Message ("Completed scaling plan inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    if (Get-Command -Name Write-SmartAvdLog -ErrorAction SilentlyContinue) { Write-SmartAvdLog -Level ERROR -Message $_.Exception.Message }
    throw
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBJCJePqmiOLpHd
# wB2+ByFicA7gL6JQdvKPAKqXhLdsZqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCD99CvWC49z5KN0kRzkYocUMZQfJJD4GD6NWMRRJVJQFjANBgkqhkiG9w0B
# AQEFAASCAYAyy8+3nh77f65TpwI4vHqaR9rmsJLpCZUjeAxq8QQa6xdtzsiH0dw6
# kzLNdvsfflFj5UFQEPRd4KAG+o5kuogrgBNYkP2pjiO0qjqR5tuUaNY8CsUta2xk
# skb1HnhIF74uCaELp/BIa/Qokn4+ANIrafICe3Xw1nbRU02EK2+8j4oV2eHu6e86
# I99gnRfhjrG5H5/91enkA0Mvi8v8OqHDMOlbBQgk7H4Cj4LmHlI9uWLEXGM3DLIW
# 9ZGsmqIziSNvPTJ3RSyFnUD5DOkBb9qAbGQoHifo3AhfVR3Mba6wheHgF3c5kue2
# 7h0BnzNimK9h+aWKWV4YFv7H9LWoCq4UzPbEY/0/BPGBmMywgJh4CFr1LJ3ie/NJ
# 2TIDS4d/MoHamklWlXJUGVNdMZBavYsiG5MMF0MVZyUSXzqoZGoieQChzQq6t5ZT
# Y0hLwFiHD4WY0V4UXqaYHQklHhWJmThhzo0rRlRB6aUquu6haIZMQ/JDFxe4LAHQ
# OZ3vIk2nFbA=
# SIG # End signature block
