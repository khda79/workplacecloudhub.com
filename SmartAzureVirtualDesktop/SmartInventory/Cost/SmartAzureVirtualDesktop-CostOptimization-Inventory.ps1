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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDxrDpUzNpWfMZ1
# R0MYKF1fDlImam6OdSrRvArw8uRabKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAzcYTIr+UIV7ShHzTynfKkGJPquIfR5atXxvuGfh8ANTANBgkqhkiG9w0B
# AQEFAASCAYCakd5IIuIJgaVV+05JWdpIstx8UFbxjGjrc795VFh5L8uEXxOdiafQ
# sgqEGGmRCaTNxOY7aetu+NprfJND9Nz+z3DsIqy2Ox2sU5R5adiUsnYMCLi9gotl
# ka6vKiE98X3WJFrvzpHvfoBfAP4nQEGRczkPK8X/QI8MaGDiuHn7PGYNpKz9wFzq
# gbhYCZfLR3mbp70yx48vGaqCh2GpJ1eo4735weRcMAGteEFjQZGgZDueHj5ujiSO
# D7WvzcpsOumAsPYHUntV9j6yarAXAbiximXq+UtmOULWxDq+R3UpAliuhh6kGgx6
# 8BozSHTHPIIh1YgYMvv3lJmQJsZ6IQOKR32/k8s90kFDqgXAs8D+cjHBdE314zZc
# plqZvYhMZ54bllbByQuUkZBnTnR5fKOvF+mqgAQdQRMLNQWuiMIviB6GmSlhl1e7
# 57CfjgoPbPxy3d2mpDT20gjvEyIFQrASyezS8SXtx1W7E2pVJ19EYYQT4w4pywsU
# EaHxPkQM0nY=
# SIG # End signature block
