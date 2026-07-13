<#
.SYNOPSIS
Exports Azure Virtual Desktop cost optimization signals.

.DESCRIPTION
Connects with Az PowerShell, enumerates Azure Virtual Desktop host pools, session hosts,
scaling plans, and related compute resources, then exports cost review candidates such as
inactive session hosts, unavailable hosts, host pools without autoscale, and unattached
managed disks in AVD resource groups.

The script is read-only. It does not stop, start, resize, or delete resources.

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
    [int]$InactiveSessionHostDays = 14,
    [string]$AvdApiVersion = '2024-04-03',
    [string]$ResourceApiVersion = '2021-04-01',
    [string]$ComputeApiVersion = '2024-07-01',
    [string]$DiskApiVersion = '2024-03-02'
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
$resolvedOutputRoots = Resolve-SmartAvdOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'AzureVirtualDesktop\Cost'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzureVirtualDesktop-CostOptimization-Inventory_{0}.log" -f $runId)

Import-SmartAvdCoreModule -StartPath $PSScriptRoot
Set-SmartAvdCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-VmPowerState {
    [CmdletBinding()]
    param([AllowNull()]$VmDetails)

    $statuses = @(Get-SmartAvdNestedPropertyValue -InputObject $VmDetails -Path @('properties', 'instanceView', 'statuses'))
    foreach ($status in $statuses) {
        $code = Get-SmartAvdObjectPropertyValue -InputObject $status -PropertyName @('code')
        if ($code -like 'PowerState/*') { return ($code -replace '^PowerState/', '') }
    }
    return ''
}

function Get-CostFinding {
    [CmdletBinding()]
    param(
        [string]$Status,
        [int]$Sessions,
        [AllowNull()]$LastHeartBeat,
        [AllowNull()]$AllowNewSession,
        [string]$VmPowerState,
        [int]$InactiveDays
    )

    $findings = @()
    if ($Status -and $Status -notin @('Available', 'NeedsAssistance')) { $findings += "SessionHostStatus=$Status" }
    if ($Status -eq 'NeedsAssistance') { $findings += 'NeedsAssistance' }
    if ($Sessions -eq 0) { $findings += 'NoActiveSessions' }
    if ($AllowNewSession -eq $false) { $findings += 'DrainMode' }
    if ($VmPowerState -eq 'running' -and $Sessions -eq 0) { $findings += 'RunningWithNoSessions' }
    if ($LastHeartBeat) {
        $ageDays = Get-SmartAvdAgeInDays -DateTimeValue $LastHeartBeat
        if ($ageDays -ne $null -and $ageDays -ge $InactiveDays) { $findings += ("HeartbeatOlderThan{0}Days" -f $InactiveDays) }
    }
    return ($findings | Select-Object -Unique) -join ';'
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAvdLog -Message "Starting SmartAzureVirtualDesktop cost optimization inventory. RunId=$runId"
    Import-SmartAvdRequiredModule -Name Az.Accounts
    $context = Connect-SmartAvdCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw 'No Azure subscriptions were found for the current identity and filters.' }

    $sessionHostCostRows = @()
    $hostPoolCostRows = @()
    $unattachedDiskRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAvdLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $avdResourceGroups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $assignedHostPoolIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $hostPools = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.DesktopVirtualization/hostPools' -ApiVersion $ResourceApiVersion)
        $scalingPlans = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.DesktopVirtualization/scalingPlans' -ApiVersion $ResourceApiVersion)

        foreach ($plan in $scalingPlans) {
            $details = Get-SmartAvdResourceById -ResourceId $plan.id -ApiVersion $AvdApiVersion -Operation 'Scaling plan details'
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            foreach ($reference in @(Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('hostPoolReferences'))) {
                $hostPoolArmPath = Get-SmartAvdObjectPropertyValue -InputObject $reference -PropertyName @('hostPoolArmPath', 'hostPoolResourceId', 'id')
                if (-not [string]::IsNullOrWhiteSpace([string]$hostPoolArmPath)) { [void]$assignedHostPoolIds.Add([string]$hostPoolArmPath) }
            }
        }

        foreach ($hostPool in $hostPools) {
            $resourceGroupName = Get-SmartAvdResourceGroupNameFromId -ResourceId $hostPool.id
            if (-not [string]::IsNullOrWhiteSpace($resourceGroupName)) { [void]$avdResourceGroups.Add($resourceGroupName) }
            $hostPoolName = Get-SmartAvdResourceNameFromId -ResourceId $hostPool.id
            $details = Get-SmartAvdResourceById -ResourceId $hostPool.id -ApiVersion $AvdApiVersion -Operation 'Host pool details'
            if ($null -eq $details) { $details = $hostPool }
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            $hasScalingPlan = $assignedHostPoolIds.Contains($hostPool.id)
            $hostPoolType = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('hostPoolType')
            $maxSessionLimit = [int](Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('maxSessionLimit'))
            $hosts = @(Get-SmartAvdChildResources -ParentResourceId $hostPool.id -ChildPath 'sessionHosts' -ApiVersion $AvdApiVersion -Operation "Session hosts for $hostPoolName")

            $runningNoSessionCount = 0
            $inactiveHostCount = 0
            $unavailableHostCount = 0
            $sessionCount = 0

            foreach ($sessionHost in $hosts) {
                $sessionHostProperties = Get-SmartAvdObjectPropertyValue -InputObject $sessionHost -PropertyName @('properties')
                $sessionHostName = Get-SmartAvdResourceNameFromId -ResourceId $sessionHost.id
                $vmResourceId = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('resourceId')
                $vmDetails = $null
                if (-not [string]::IsNullOrWhiteSpace([string]$vmResourceId)) {
                    $vmDetails = Get-SmartAvdResourceById -ResourceId ("{0}?`$expand=instanceView" -f $vmResourceId) -ApiVersion $ComputeApiVersion -Operation "VM instance view for $sessionHostName"
                }

                $vmProperties = Get-SmartAvdObjectPropertyValue -InputObject $vmDetails -PropertyName @('properties')
                $hardwareProfile = Get-SmartAvdObjectPropertyValue -InputObject $vmProperties -PropertyName @('hardwareProfile')
                $storageProfile = Get-SmartAvdObjectPropertyValue -InputObject $vmProperties -PropertyName @('storageProfile')
                $status = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('status')
                $sessions = [int](Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('sessions'))
                $lastHeartBeat = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('lastHeartBeat')
                $allowNewSession = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('allowNewSession')
                $vmPowerState = Get-VmPowerState -VmDetails $vmDetails
                $finding = Get-CostFinding -Status $status -Sessions $sessions -LastHeartBeat $lastHeartBeat -AllowNewSession $allowNewSession -VmPowerState $vmPowerState -InactiveDays $InactiveSessionHostDays
                $heartbeatAgeDays = Get-SmartAvdAgeInDays -DateTimeValue $lastHeartBeat

                $sessionCount += $sessions
                if ($vmPowerState -eq 'running' -and $sessions -eq 0) { $runningNoSessionCount++ }
                if ($heartbeatAgeDays -ne $null -and $heartbeatAgeDays -ge $InactiveSessionHostDays) { $inactiveHostCount++ }
                if ($status -and $status -notin @('Available', 'NeedsAssistance')) { $unavailableHostCount++ }

                $sessionHostCostRows += [pscustomobject]@{
                    RunId              = $runId
                    SubscriptionId     = $subscription.Id
                    SubscriptionName   = $subscription.Name
                    ResourceGroupName  = $resourceGroupName
                    HostPoolName       = $hostPoolName
                    HostPoolType       = $hostPoolType
                    SessionHostName    = $sessionHostName
                    CostFinding        = $finding
                    SessionHostStatus  = $status
                    AllowNewSession    = ConvertTo-SmartAvdBoolString -Value $allowNewSession
                    Sessions           = $sessions
                    LastHeartBeat      = $lastHeartBeat
                    HeartbeatAgeDays   = $heartbeatAgeDays
                    VmPowerState       = $vmPowerState
                    VmSize             = Get-SmartAvdObjectPropertyValue -InputObject $hardwareProfile -PropertyName @('vmSize')
                    OsDiskType         = Get-SmartAvdNestedPropertyValue -InputObject $storageProfile -Path @('osDisk', 'managedDisk', 'storageAccountType')
                    DataDiskCount      = @(Get-SmartAvdObjectPropertyValue -InputObject $storageProfile -PropertyName @('dataDisks')).Count
                    AgentVersion       = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('agentVersion')
                    VmResourceId       = $vmResourceId
                    ResourceId         = $sessionHost.id
                }
            }

            $capacity = if ($maxSessionLimit -gt 0) { $hosts.Count * $maxSessionLimit } else { 0 }
            $hostPoolCostRows += [pscustomobject]@{
                RunId                    = $runId
                SubscriptionId           = $subscription.Id
                SubscriptionName         = $subscription.Name
                ResourceGroupName        = $resourceGroupName
                HostPoolName             = $hostPoolName
                HostPoolType             = $hostPoolType
                HasScalingPlan           = $hasScalingPlan
                CostFinding              = if ($hasScalingPlan) { '' } else { 'NoScalingPlanAssignment' }
                SessionHostCount         = $hosts.Count
                RunningNoSessionHostCount = $runningNoSessionCount
                InactiveHostCount        = $inactiveHostCount
                UnavailableHostCount     = $unavailableHostCount
                CurrentSessionCount      = $sessionCount
                MaxSessionLimit          = $maxSessionLimit
                TheoreticalCapacity      = $capacity
                CapacityUsedPercent      = if ($capacity -gt 0) { [math]::Round(($sessionCount / $capacity) * 100, 2) } else { $null }
                LoadBalancerType         = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('loadBalancerType')
                StartVMOnConnect         = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('startVMOnConnect'))
                ResourceId               = $hostPool.id
            }
        }

        $disks = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.Compute/disks' -ApiVersion $ResourceApiVersion)
        $disks = @($disks | Where-Object { $avdResourceGroups.Contains((Get-SmartAvdResourceGroupNameFromId -ResourceId $_.id)) })
        foreach ($disk in $disks) {
            $details = Get-SmartAvdResourceById -ResourceId $disk.id -ApiVersion $DiskApiVersion -Operation 'Managed disk details'
            if ($null -eq $details) { $details = $disk }
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            $managedBy = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('managedBy')
            $diskState = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('diskState')
            if ([string]::IsNullOrWhiteSpace([string]$managedBy) -or $diskState -eq 'Unattached') {
                $unattachedDiskRows += [pscustomobject]@{
                    RunId              = $runId
                    SubscriptionId     = $subscription.Id
                    SubscriptionName   = $subscription.Name
                    ResourceGroupName  = Get-SmartAvdResourceGroupNameFromId -ResourceId $disk.id
                    DiskName           = Get-SmartAvdResourceNameFromId -ResourceId $disk.id
                    Location           = $details.location
                    DiskState          = $diskState
                    DiskSizeGb         = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('diskSizeGB')
                    SkuName            = Get-SmartAvdNestedPropertyValue -InputObject $details -Path @('sku', 'name')
                    OsType             = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('osType')
                    TimeCreated        = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('timeCreated')
                    ManagedBy          = $managedBy
                    CostFinding        = 'UnattachedDiskInAvdResourceGroup'
                    TagsJson           = ConvertTo-SmartAvdCompactJson -Value $details.tags
                    ResourceId         = $disk.id
                }
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                    = $runId
            SubscriptionId           = $subscription.Id
            SubscriptionName         = $subscription.Name
            TenantId                 = $subscription.TenantId
            HostPoolCount            = $hostPools.Count
            HostPoolsWithoutAutoscale = @($hostPoolCostRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.HasScalingPlan -eq $false }).Count
            SessionHostCount         = @($sessionHostCostRows | Where-Object { $_.SubscriptionId -eq $subscription.Id }).Count
            RunningNoSessionHostCount = @($sessionHostCostRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.CostFinding -like '*RunningWithNoSessions*' }).Count
            InactiveSessionHostCount = @($sessionHostCostRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.CostFinding -like '*HeartbeatOlderThan*' }).Count
            UnattachedDiskCount      = @($unattachedDiskRows | Where-Object { $_.SubscriptionId -eq $subscription.Id }).Count
        }
    }

    Export-SmartAvdCsv -Name 'AVD_Cost_SessionHostSignals' -Rows $sessionHostCostRows
    Export-SmartAvdCsv -Name 'AVD_Cost_HostPoolSignals' -Rows $hostPoolCostRows
    Export-SmartAvdCsv -Name 'AVD_Cost_UnattachedDisks' -Rows $unattachedDiskRows
    Export-SmartAvdCsv -Name 'AVD_Cost_Summary' -Rows $summaryRows

    Write-SmartAvdLog -Level SUCCESS -Message ("Completed cost optimization inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    if (Get-Command -Name Write-SmartAvdLog -ErrorAction SilentlyContinue) { Write-SmartAvdLog -Level ERROR -Message $_.Exception.Message }
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDXJcP37uZoxZtG
# 2zG5+x0cphXjf9H2uAdFH7GpZT2sY6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGF1BcAQ2fBucscWqAVLLdtOGwocyTJPvfWDL48kkug7MA0GCSqG
# SIb3DQEBAQUABIIBgJ77M7QGur09B04/hSKEdprFTMbqHuP0zYtP0CumS/yAmwEv
# 1eq7P5yEFBAvCXqNmyM8PQ3NEYpLu9OqYmkt6EvAyaXsY37Lb7dkSazzThDr5Gz0
# ILtSm5HiGMEMTxbK4/mC8jbrNdZ/5jBNvGaoaEYo9wur2j5/o78EyVA9DunNNbIW
# ig7vJRW0L+Fl60W2yT+bFtxNNN7rD5+9z2aZOR3vFxclTIYDeHww/UKxS9TjtRDP
# buzuf3S8OnAWy76CbwpUl5l55W9Q3qbWEGrYLzFIs3h5W8H/U4CM6+fx0AMwSZUv
# n9bUy1YAHfL8vR+H3KM/S3SO07QWaKc7CCWBbe4DXyHcvpNol4WwV5e+0qoSvTgP
# CQv30TGGkqAs4IgZsvWfRii5MvAo+bmbcTxz81fKI/LbWYgfwW073BPUybgvAq7R
# JPNZimvp0YMAqrt97OnlzJMYwFcOARUTSxb8NKA76DA0RqK5A7ZldMIoXa6oB+eo
# uHfmePJaGB/Gb/os9KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MTVaMC8GCSqGSIb3DQEJBDEiBCBAaQbLD9JijVd1I4YTmIMFOJamw0HQWNFT8TCD
# xL5DGDANBgkqhkiG9w0BAQEFAASCAgCSxOaKR/ix0FgYEUwucM36uJoDg5XE9pO5
# rFXBp04HbqgN6bj1Igbd8gMpLI7IcTYSuPsNf6b96dCI81nyhwn/cRAcaP4KSEq1
# YgBHyCcC9oQybUAt6fXnscHDuBk2Lk2DhsYMbPvH2cwiCPDvE0MaIQs+i73yS7zV
# RvZRzWthjDT2T3dE8CezOZTbIiPW47esNBrvBlzXQsn/jkmSHu12/0H0xRvKxFiR
# OyFZWkihLp704Sp04ErNP304MzsJu4r3U2ut6hpz3Sb8l5NZkE4pUSnULDBMCxWI
# AsBGqqjIBloCMzq7xSNnBuCg4Y3kWmP4Hy68lzNf3+xcgkeLc/IIw/rc53ZgwYR+
# OLizKi5LYf8iuk4/6vSQb1gv6D450O+dSmVNS6vNbli5zAL0o2Ayl87n5SvFmv4u
# ywVI8TpghmLv/tgARs8IRX5Jz/aMpyBcoVhPzPSd6njqf2tZn+8kKrwo+jnvL3Ks
# XYe3MJQArwS0x2NebFjnfPpOtdTk5YNbVVHEEy6I9NL/kEtRYF8cDraY+tQZu4/3
# xidWcb9q9qK63YQ07KUQU4mRrj6fw37NEylWQqESJ/+2L7SBaVV592cB3bcrJXQf
# MnFaHHVp9a2PAJSxr9X+TqAenC2dNFot1mn8bSI0r32LfbD/D9CJdDGfYchgKhXY
# stv6SBZ0SA==
# SIG # End signature block
