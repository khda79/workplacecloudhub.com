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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCChIIbHIj09ZLGe
# CyujsvfF9jd/adhljhV08vlonN+oYKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDhxpJdR9r/+jp4YDERybzFcIeVxR0B17u2ed96J6JMwzANBgkqhkiG9w0B
# AQEFAASCAYAnv0wKqFmlw05YxUcq79bknHXpgCxCENz+cLEVpytkT6RxLWK6N2Xg
# 1PVIC3KIx97N9gn0XY1NzyGraxO9bC0jQ+QDhumYZV4X/C3Y0jWGCdmR9dwZFedv
# HP+pcApPdOSv/Sr65Y28LsVjIk/wAQVLqgXd6v06K8SJwb6brtIrQIy70XUoPpj2
# 8Ezl1fAB5xCjQJHo4DNNRNnZDWAfyRQHqbXRbx6QFvNMFYTjKzTpY7FrRZBs9B7k
# 0djDoxeSp9TC94vnruJ4q/BglKoASsAqbRExbkXaYwXUwepgN6cFreuq3zPlqWvB
# uBYFXhk/0ZKpwzsbpTGtieDj2PIlFMlxplAvQdTaA/2oFGQkmn/cXlFbkLFOhHml
# anX5DZjpcDzcF7p7qU+JaabQDiCtEH5TwuWjO+3jWm2/hNctNqqzZC670kprHdUp
# 1YjG7DvtHkpK3ntNuyxsH/cNVod2t/YpGLoOncpcGE8aAPVsJ9AwsP3wZomcft5h
# VsIGdBRQeaQ=
# SIG # End signature block
