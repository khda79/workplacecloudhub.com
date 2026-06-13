<#
.SYNOPSIS
Exports Azure Storage account security posture inventory.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
storage account public network access, anonymous blob access, shared key access, HTTPS only,
minimum TLS version, firewall settings, and blob service properties.

The script is read-only and uses Azure Resource Manager REST calls through Invoke-AzRestMethod.

.PARAMETER Tenant
Local tenant profile key used to isolate SmartAzure output folders. Defaults to test.

.PARAMETER TenantId
Optional tenant ID or domain used by Connect-AzAccount. If omitted, the SmartAzure tenant
profile AzureTenantId, TenantId, or OrgDomain is used when available.

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
    [string]$StorageApiVersion = '2023-01-01'
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
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { $LatestOutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue '' }

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\Storage'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-StorageSecurity-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-ResourceGroupFromId {
    [CmdletBinding()]
    param([string]$ResourceId)
    if ($ResourceId -match '/resourceGroups/([^/]+)') { return $matches[1] }
    return ''
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure Storage security inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts

    $context = Connect-SmartAzureCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw "No Azure subscriptions were found for the current identity and filters." }

    $accountRows = @()
    $blobServiceRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null
        $encodedSubscriptionId = [uri]::EscapeDataString($subscription.Id)
        $accounts = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Storage/storageAccounts?api-version={1}" -f $encodedSubscriptionId, $StorageApiVersion) -Operation "Storage accounts for $($subscription.Name)"

        foreach ($account in $accounts) {
            $properties = Get-ObjectPropertyValue -InputObject $account -PropertyName @('properties')
            $network = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('networkAcls')
            $accountRows += [pscustomobject]@{
                RunId                   = $runId
                SubscriptionId          = $subscription.Id
                SubscriptionName        = $subscription.Name
                ResourceGroup           = Get-ResourceGroupFromId -ResourceId $account.id
                StorageAccountName      = $account.name
                Location                = $account.location
                Kind                    = $account.kind
                SkuName                 = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $account -PropertyName @('sku')) -PropertyName @('name')
                PublicNetworkAccess     = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('publicNetworkAccess')
                AllowBlobPublicAccess   = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('allowBlobPublicAccess')
                AllowSharedKeyAccess    = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('allowSharedKeyAccess')
                SupportsHttpsTrafficOnly = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('supportsHttpsTrafficOnly')
                MinimumTlsVersion       = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('minimumTlsVersion')
                DefaultAction           = Get-ObjectPropertyValue -InputObject $network -PropertyName @('defaultAction')
                Bypass                  = Get-ObjectPropertyValue -InputObject $network -PropertyName @('bypass')
                IpRulesCount            = @((Get-ObjectPropertyValue -InputObject $network -PropertyName @('ipRules'))).Count
                VirtualNetworkRulesCount = @((Get-ObjectPropertyValue -InputObject $network -PropertyName @('virtualNetworkRules'))).Count
                PrivateEndpointConnectionCount = @((Get-ObjectPropertyValue -InputObject $properties -PropertyName @('privateEndpointConnections'))).Count
                AccessTier              = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('accessTier')
                CreationTime            = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('creationTime')
                TagsJson                = ConvertTo-CompactJson -Value $account.tags
                ResourceId              = $account.id
            }

            $blobUri = "https://management.azure.com{0}/blobServices/default?api-version={1}" -f $account.id, $StorageApiVersion
            $blobServices = Invoke-SmartAzureArmGetPaged -Uri $blobUri -Operation "Blob service properties for $($account.name)"
            foreach ($blobService in $blobServices) {
                $blobProperties = Get-ObjectPropertyValue -InputObject $blobService -PropertyName @('properties')
                $deleteRetention = Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('deleteRetentionPolicy')
                $containerDeleteRetention = Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('containerDeleteRetentionPolicy')
                $blobServiceRows += [pscustomobject]@{
                    RunId                         = $runId
                    SubscriptionId                = $subscription.Id
                    SubscriptionName              = $subscription.Name
                    ResourceGroup                 = Get-ResourceGroupFromId -ResourceId $account.id
                    StorageAccountName            = $account.name
                    IsVersioningEnabled           = Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('isVersioningEnabled')
                    ChangeFeedEnabled             = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('changeFeed')) -PropertyName @('enabled')
                    BlobDeleteRetentionEnabled    = Get-ObjectPropertyValue -InputObject $deleteRetention -PropertyName @('enabled')
                    BlobDeleteRetentionDays       = Get-ObjectPropertyValue -InputObject $deleteRetention -PropertyName @('days')
                    ContainerDeleteRetentionEnabled = Get-ObjectPropertyValue -InputObject $containerDeleteRetention -PropertyName @('enabled')
                    ContainerDeleteRetentionDays  = Get-ObjectPropertyValue -InputObject $containerDeleteRetention -PropertyName @('days')
                    CorsRulesJson                 = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('cors')) -PropertyName @('corsRules'))
                    PropertiesJson                = ConvertTo-CompactJson -Value $blobProperties
                    ResourceId                    = $blobService.id
                }
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                         = $runId
            SubscriptionId                = $subscription.Id
            SubscriptionName              = $subscription.Name
            TenantId                      = $subscription.TenantId
            StorageAccountCount           = $accounts.Count
            PublicNetworkEnabledCount     = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.PublicNetworkAccess -eq 'Enabled' }).Count
            BlobPublicAccessAllowedCount  = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.AllowBlobPublicAccess -eq $true }).Count
            SharedKeyAllowedCount         = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.AllowSharedKeyAccess -ne $false }).Count
            NonHttpsOnlyCount             = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.SupportsHttpsTrafficOnly -ne $true }).Count
            FirewallAllowByDefaultCount   = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.DefaultAction -eq 'Allow' }).Count
        }
    }

    Export-SmartAzureCsv -Name 'Azure_Storage_AccountsSecurity' -Rows $accountRows
    Export-SmartAzureCsv -Name 'Azure_Storage_BlobServiceSecurity' -Rows $blobServiceRows
    Export-SmartAzureCsv -Name 'Azure_Storage_Summary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed Storage security inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}
