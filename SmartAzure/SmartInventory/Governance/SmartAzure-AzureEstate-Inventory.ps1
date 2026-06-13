<#
.SYNOPSIS
Exports a broad Azure Resource Manager estate inventory.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
management groups, subscriptions, Azure regions, resource groups, resources, resource locks,
resource providers, and a per-subscription summary.

The script is read-only. It does not modify Azure resources.

.PARAMETER TenantId
Optional tenant ID or domain used by Connect-AzAccount.

.PARAMETER Tenant
Local tenant profile key used to isolate SmartAzure output folders. Defaults to test.

.PARAMETER SubscriptionId
Optional list of subscription IDs to include. When omitted, all visible subscriptions are included.

.PARAMETER OutputRoot
Historical output root. A run-specific folder is created below this path.

.PARAMETER LatestOutputRoot
Latest CSV output root. Stable non-timestamped copies are written here.

.PARAMETER Connect
Forces an Azure sign-in before inventory.

.PARAMETER UseDeviceCode
Uses device code authentication for Connect-AzAccount.

.PARAMETER SkipManagementGroups
Skips management group inventory.

.PARAMETER SkipResources
Skips resource inventory.

.PARAMETER SkipLocks
Skips resource lock inventory.

.PARAMETER SkipProviders
Skips resource provider inventory.

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
Requires: Az.Accounts, Az.Resources
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
    [switch]$SkipManagementGroups,
    [switch]$SkipResources,
    [switch]$SkipLocks,
    [switch]$SkipProviders
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path -Path $d -ChildPath 'Config\SmartAzure-TenantContext.ps1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartAzure-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartAzureEffectiveConfig = Initialize-SmartAzureTenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ScriptLocalConfig = Get-SmartAzureScriptLocalConfig -ScriptPath $PSCommandPath
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    $TenantId = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'AzureTenantId' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') {
        $TenantId = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue ''
    }
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') {
        $TenantId = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue ''
    }
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue ''
}
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) {
    $LatestOutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue ''
}

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\Estate'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-AzureEstate-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function ConvertTo-TagString {
    [CmdletBinding()]
    param($Tags)

    if ($null -eq $Tags) { return '' }
    return ConvertTo-CompactJson -Value $Tags
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure Azure estate inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts
    Import-RequiredModule -Name Az.Resources

    $context = Connect-SmartAzureCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $allSubscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $allSubscriptions = @($allSubscriptions | Where-Object { $wanted.Contains($_.Id) })
    }

    if ($allSubscriptions.Count -eq 0) {
        throw "No Azure subscriptions were found for the current identity and filters."
    }

    Write-SmartAzureLog -Message ("Subscription count: {0}" -f $allSubscriptions.Count)

    $managementGroupRows = @()
    if (-not $SkipManagementGroups) {
        $managementGroupRows = @(Invoke-SafeInventoryBlock -Name 'Management group inventory' -ScriptBlock {
            $groups = @(Get-AzManagementGroup)
            foreach ($group in $groups) {
                [pscustomobject]@{
                    RunId             = $runId
                    ManagementGroupId = $group.Name
                    DisplayName       = $group.DisplayName
                    Type              = $group.Type
                    TenantId          = $group.TenantId
                    ParentId          = if ($group.ParentName) { $group.ParentName } else { '' }
                }
            }
        })
    }

    $subscriptionRows = @()
    $locationRows = @()
    $resourceGroupRows = @()
    $resourceRows = @()
    $lockRows = @()
    $providerRows = @()
    $summaryRows = @()

    foreach ($subscription in $allSubscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $subscriptionRows += [pscustomobject]@{
            RunId             = $runId
            SubscriptionId    = $subscription.Id
            SubscriptionName  = $subscription.Name
            TenantId          = $subscription.TenantId
            State             = $subscription.State
            AuthorizationSource = ($subscription.AuthorizationSource -join ';')
        }

        $locations = @(Invoke-SafeInventoryBlock -Name "Location inventory for $($subscription.Name)" -ScriptBlock { Get-AzLocation })
        foreach ($location in $locations) {
            $locationRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                Location         = $location.Location
                DisplayName      = $location.DisplayName
                GeographyGroup   = $location.GeographyGroup
                RegionCategory   = $location.RegionCategory
                RegionType       = $location.RegionType
                PhysicalLocation = $location.PhysicalLocation
                PairedRegion     = ConvertTo-CompactJson -Value $location.PairedRegion
            }
        }

        $resourceGroups = @(Invoke-SafeInventoryBlock -Name "Resource group inventory for $($subscription.Name)" -ScriptBlock { Get-AzResourceGroup })
        foreach ($resourceGroup in $resourceGroups) {
            $resourceGroupRows += [pscustomobject]@{
                RunId              = $runId
                SubscriptionId     = $subscription.Id
                SubscriptionName   = $subscription.Name
                ResourceGroupName  = $resourceGroup.ResourceGroupName
                Location           = $resourceGroup.Location
                ProvisioningState  = $resourceGroup.ProvisioningState
                ManagedBy          = $resourceGroup.ManagedBy
                TagsJson           = ConvertTo-TagString -Tags $resourceGroup.Tags
                ResourceId         = $resourceGroup.ResourceId
            }
        }

        $resources = @()
        if (-not $SkipResources) {
            $resources = @(Invoke-SafeInventoryBlock -Name "Resource inventory for $($subscription.Name)" -ScriptBlock { Get-AzResource })
            foreach ($resource in $resources) {
                $resourceRows += [pscustomobject]@{
                    RunId             = $runId
                    SubscriptionId    = $subscription.Id
                    SubscriptionName  = $subscription.Name
                    ResourceGroupName = $resource.ResourceGroupName
                    Name              = $resource.Name
                    ResourceType      = $resource.ResourceType
                    Location          = $resource.Location
                    Kind              = $resource.Kind
                    SkuName           = if ($resource.Sku) { $resource.Sku.Name } else { '' }
                    SkuTier           = if ($resource.Sku) { $resource.Sku.Tier } else { '' }
                    ManagedBy         = $resource.ManagedBy
                    IdentityType      = if ($resource.Identity) { $resource.Identity.Type } else { '' }
                    PlanName          = if ($resource.Plan) { $resource.Plan.Name } else { '' }
                    PlanProduct       = if ($resource.Plan) { $resource.Plan.Product } else { '' }
                    PlanPublisher     = if ($resource.Plan) { $resource.Plan.Publisher } else { '' }
                    TagsJson          = ConvertTo-TagString -Tags $resource.Tags
                    ResourceId        = $resource.ResourceId
                }
            }
        }

        $locks = @()
        if (-not $SkipLocks) {
            $locks = @(Invoke-SafeInventoryBlock -Name "Lock inventory for $($subscription.Name)" -ScriptBlock { Get-AzResourceLock })
        foreach ($lock in $locks) {
            $lockProperties = Get-ObjectPropertyValue -InputObject $lock -PropertyName @('Properties')
            $lockRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                LockName         = $lock.Name
                LockLevel        = if ($lockProperties) { Get-ObjectPropertyValue -InputObject $lockProperties -PropertyName @('Level') } else { Get-ObjectPropertyValue -InputObject $lock -PropertyName @('Level') }
                Notes            = if ($lockProperties) { Get-ObjectPropertyValue -InputObject $lockProperties -PropertyName @('Notes') } else { Get-ObjectPropertyValue -InputObject $lock -PropertyName @('Notes') }
                ResourceGroupName = $lock.ResourceGroupName
                ResourceType     = $lock.ResourceType
                ResourceName     = $lock.ResourceName
                LockId           = if ($lock.LockId) { $lock.LockId } else { $lock.Id }
            }
        }
        }

        $providers = @()
        if (-not $SkipProviders) {
            $providers = @(Invoke-SafeInventoryBlock -Name "Provider inventory for $($subscription.Name)" -ScriptBlock { Get-AzResourceProvider -ListAvailable })
            foreach ($provider in $providers) {
                $providerRows += [pscustomobject]@{
                    RunId              = $runId
                    SubscriptionId     = $subscription.Id
                    SubscriptionName   = $subscription.Name
                    ProviderNamespace  = $provider.ProviderNamespace
                    RegistrationState  = $provider.RegistrationState
                    ResourceTypesJson  = ConvertTo-CompactJson -Value $provider.ResourceTypes
                }
            }
        }

        $resourceTypeCounts = @{}
        foreach ($resource in $resources) {
            if ([string]::IsNullOrWhiteSpace($resource.ResourceType)) { continue }
            if (-not $resourceTypeCounts.ContainsKey($resource.ResourceType)) {
                $resourceTypeCounts[$resource.ResourceType] = 0
            }
            $resourceTypeCounts[$resource.ResourceType]++
        }

        $summaryRows += [pscustomobject]@{
            RunId              = $runId
            SubscriptionId     = $subscription.Id
            SubscriptionName   = $subscription.Name
            TenantId           = $subscription.TenantId
            State              = $subscription.State
            ResourceGroupCount = $resourceGroups.Count
            ResourceCount      = $resources.Count
            LockCount          = $locks.Count
            ProviderCount      = $providers.Count
            ResourceTypesJson  = ConvertTo-CompactJson -Value $resourceTypeCounts
        }
    }

    Export-SmartAzureCsv -Name 'Azure_ManagementGroups' -Rows $managementGroupRows
    Export-SmartAzureCsv -Name 'Azure_Subscriptions' -Rows $subscriptionRows
    Export-SmartAzureCsv -Name 'Azure_Locations' -Rows $locationRows
    Export-SmartAzureCsv -Name 'Azure_ResourceGroups' -Rows $resourceGroupRows
    Export-SmartAzureCsv -Name 'Azure_Resources' -Rows $resourceRows
    Export-SmartAzureCsv -Name 'Azure_ResourceLocks' -Rows $lockRows
    Export-SmartAzureCsv -Name 'Azure_ResourceProviders' -Rows $providerRows
    Export-SmartAzureCsv -Name 'Azure_Estate_Summary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed Azure estate inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}
