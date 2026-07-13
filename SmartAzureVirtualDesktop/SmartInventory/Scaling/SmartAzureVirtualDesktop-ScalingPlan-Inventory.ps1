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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCAZ5sKuxgITpMa
# Mj73W0BNzLb87O11ULWnmkvQBPmJiqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIO8G40VB6AXwhxlNsQGY46Lth+AvWZ6rHSGGccXn6Aq1MA0GCSqG
# SIb3DQEBAQUABIIBgCkowZHh2yJOIfCkCoSamaB71ZTL08ZfvDwUPLb8S6mJ5R/l
# hcc+c9AYPgHkfGXlbq78veopvWaYyQSdo9qyZh9VvRtWiYf1cvYDAozA2o92W9pr
# AXOgw9d1tH+yBajUgvQ9FraPw7TjofxE3jFTfa+Z2SF8exPtenHFLIOK2qpAD789
# 4s1J+pWOSP3lSvEUgZstegdwPYki4G5xWBNVgv/u13EJ/FZttJrDZcWLQ38Ts7xD
# d5wC7fTnZX9F0l5bqgJXX9JnsSuerp0OTCNpKpqvh+FOjy0ITLgYsKKJO+ayPHzD
# DfeAc4wn+TrM7i1+4qI/zotCh07dWYyCfuH2PPQUEf9pSJiWrFOUpy91bHtImTKX
# E2WxhbLjqWFu1+oXGDvYGr4FwH2T9orqMf2/7qcQbx5LEEOYt4FSBttIZ/XFB6tb
# 5oW8qEbt+8gT3usPCWpMBob5CqdYz4CU1VWdy/RVulboUjMjKFczjSVpuWBxBAuz
# wk9SdY18djlI/GLC6qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MTdaMC8GCSqGSIb3DQEJBDEiBCBE62c1v8TmYWjXxuhdvnC6BgTA8C66WL3qbijH
# Rm8NjTANBgkqhkiG9w0BAQEFAASCAgBVIbudVYdMdJRCc8+0/pVuLAN/vLFrVc1C
# /zEDWLDBdOSZIHk3Iv1LD31mtHN0uUOQ3sSkjkF8HFNF6WzN8VXmbho3dqwvFAny
# AqCRe8CSO7SW5RQwfBNPYwigrXPZhxsFkCjJUcdUVsH1GYDRXT1u4v5yf2Nfu96u
# Z93Pz3hAnqlnuDSzaqq7J1lzzlwY7aDy5xgLN43KjlBt9u7ERMyDw23G6ZQk7Jzf
# D/aBvQ5hJMOEphQRoE/BfpWVyTdA68IUAQ1QOm2BdQ1i0wB6A9GPYrfUgm2moLRT
# kr8Rc2pAekIbG22rSX2laFxZqp58oHd5kMRm3cofEc8JiPS0qHNGI29gKyQ+gctF
# R/fma23ij6Pu/JDsVr9IZvV71GjZgnML8JNyQcf5/fkFa5oaKNpM2k6M++mUSWoF
# 2U1zFDnYuxKq7PYzxAWArVqSyC8XOm20fqaZrfCtt1aa77B8ouWMoxG7vbliJYkh
# DryW2e/RHdvxbDR0EwCigAfpMhYqOEE0XBwKpqpy242iPo6cROnscc91ENptWx/O
# e+MzoE840dviAfilmqv5sXtZ3NUBwFwGTn9dJmf5nCsQUWlkEoMh1BAK0QUdcPJh
# cD3Ln1UGJuHbneMuw1rZ8EqpqX6ZoZIM1OD6KoPc/fVfgiWlGoN55oVNBBzIeX/V
# 4s0nKZFKuQ==
# SIG # End signature block
