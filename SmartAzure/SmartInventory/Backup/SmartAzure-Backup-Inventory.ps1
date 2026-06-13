<#
.SYNOPSIS
Exports Azure Recovery Services backup posture inventory.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
Recovery Services vaults, vault backup/security settings, protected backup items, Azure VMs
with detected backup protection, Azure VMs without detected Recovery Services protection, and
subscription summaries.

The script is read-only and uses Azure Resource Manager REST calls through Invoke-AzRestMethod.

.PARAMETER Tenant
Local tenant profile key used to isolate SmartAzure output folders. Defaults to test.

.PARAMETER TenantId
Optional tenant ID or domain used by Connect-AzAccount. If omitted, the SmartAzure tenant
profile AzureTenantId, TenantId, or OrgDomain is used when available.

.PARAMETER SubscriptionId
Optional list of subscription IDs to include. When omitted, all visible subscriptions are included.

.PARAMETER SkipProtectedItems
Skips backup protected item enumeration.

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
    [switch]$SkipProtectedItems,
    [string]$OutputRoot,
    [string]$LatestOutputRoot,
    [switch]$Connect,
    [switch]$UseDeviceCode,
    [string]$RecoveryServicesApiVersion = '2023-02-01',
    [string]$ComputeApiVersion = '2023-07-01'
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
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\Backup'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-Backup-Inventory_{0}.log" -f $runId)

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

    Write-SmartAzureLog -Message "Starting SmartAzure Backup inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts

    $context = Connect-SmartAzureCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw "No Azure subscriptions were found for the current identity and filters." }

    $vaultRows = @()
    $vaultSettingRows = @()
    $protectedItemRows = @()
    $protectedVmRows = @()
    $unprotectedVmRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null
        $encodedSubscriptionId = [uri]::EscapeDataString($subscription.Id)

        $vaults = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.RecoveryServices/vaults?api-version={1}" -f $encodedSubscriptionId, $RecoveryServicesApiVersion) -Operation "Recovery Services vaults for $($subscription.Name)"
        $vms = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Compute/virtualMachines?api-version={1}" -f $encodedSubscriptionId, $ComputeApiVersion) -Operation "Virtual machines for $($subscription.Name)"
        $protectedVmIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $subscriptionProtectedItems = @()

        foreach ($vault in $vaults) {
            $properties = Get-ObjectPropertyValue -InputObject $vault -PropertyName @('properties')
            $vaultRg = Get-ResourceGroupFromId -ResourceId $vault.id
            $vaultRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                ResourceGroup    = $vaultRg
                VaultName        = $vault.name
                Location         = $vault.location
                SkuName          = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $vault -PropertyName @('sku')) -PropertyName @('name')
                ProvisioningState = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('provisioningState')
                PublicNetworkAccess = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('publicNetworkAccess')
                SoftDeleteStateJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('securitySettings')) -PropertyName @('softDeleteSettings'))
                ImmutabilityStateJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('securitySettings')) -PropertyName @('immutabilitySettings'))
                TagsJson         = ConvertTo-CompactJson -Value $vault.tags
                PropertiesJson   = ConvertTo-CompactJson -Value $properties
                ResourceId       = $vault.id
            }

            $settingsUri = "https://management.azure.com/subscriptions/{0}/resourceGroups/{1}/providers/Microsoft.RecoveryServices/vaults/{2}/backupconfig/vaultconfig?api-version={3}" -f $encodedSubscriptionId, [uri]::EscapeDataString($vaultRg), [uri]::EscapeDataString($vault.name), $RecoveryServicesApiVersion
            $settings = Invoke-SmartAzureArmGetPaged -Uri $settingsUri -Operation "Backup settings for vault $($vault.name)"
            foreach ($setting in $settings) {
                $settingProperties = Get-ObjectPropertyValue -InputObject $setting -PropertyName @('properties')
                $vaultSettingRows += [pscustomobject]@{
                    RunId            = $runId
                    SubscriptionId   = $subscription.Id
                    SubscriptionName = $subscription.Name
                    ResourceGroup    = $vaultRg
                    VaultName        = $vault.name
                    SoftDeleteFeatureState = Get-ObjectPropertyValue -InputObject $settingProperties -PropertyName @('softDeleteFeatureState')
                    EnhancedSecurityState = Get-ObjectPropertyValue -InputObject $settingProperties -PropertyName @('enhancedSecurityState')
                    StorageType      = Get-ObjectPropertyValue -InputObject $settingProperties -PropertyName @('storageType')
                    StorageTypeState = Get-ObjectPropertyValue -InputObject $settingProperties -PropertyName @('storageTypeState')
                    ResourceId       = $setting.id
                    PropertiesJson   = ConvertTo-CompactJson -Value $settingProperties
                }
            }

            if (-not $SkipProtectedItems) {
                $itemsUri = "https://management.azure.com/subscriptions/{0}/resourceGroups/{1}/providers/Microsoft.RecoveryServices/vaults/{2}/backupProtectedItems?api-version={3}" -f $encodedSubscriptionId, [uri]::EscapeDataString($vaultRg), [uri]::EscapeDataString($vault.name), $RecoveryServicesApiVersion
                $items = Invoke-SmartAzureArmGetPaged -Uri $itemsUri -Operation "Protected items for vault $($vault.name)"
                $subscriptionProtectedItems += @($items)
                foreach ($item in $items) {
                    $itemProperties = Get-ObjectPropertyValue -InputObject $item -PropertyName @('properties')
                    $sourceResourceId = Get-ObjectPropertyValue -InputObject $itemProperties -PropertyName @('sourceResourceId', 'protectedItemDataSourceId')
                    if (-not [string]::IsNullOrWhiteSpace($sourceResourceId)) { [void]$protectedVmIds.Add($sourceResourceId) }
                    $row = [pscustomobject]@{
                        RunId                 = $runId
                        SubscriptionId        = $subscription.Id
                        SubscriptionName      = $subscription.Name
                        VaultResourceGroup    = $vaultRg
                        VaultName             = $vault.name
                        ProtectedItemName     = $item.name
                        WorkloadType          = Get-ObjectPropertyValue -InputObject $itemProperties -PropertyName @('workloadType')
                        BackupManagementType  = Get-ObjectPropertyValue -InputObject $itemProperties -PropertyName @('backupManagementType')
                        ProtectionState       = Get-ObjectPropertyValue -InputObject $itemProperties -PropertyName @('protectionState')
                        PolicyName            = Get-ObjectPropertyValue -InputObject $itemProperties -PropertyName @('policyName')
                        LastBackupStatus      = Get-ObjectPropertyValue -InputObject $itemProperties -PropertyName @('lastBackupStatus')
                        LastBackupTime        = Get-ObjectPropertyValue -InputObject $itemProperties -PropertyName @('lastBackupTime')
                        SourceResourceId      = $sourceResourceId
                        ResourceId            = $item.id
                    }
                    $protectedItemRows += $row
                    if (-not [string]::IsNullOrWhiteSpace($sourceResourceId) -and $sourceResourceId -match '/providers/Microsoft.Compute/virtualMachines/') {
                        $protectedVmRows += $row
                    }
                }
            }
        }

        foreach ($vm in $vms) {
            if (-not $protectedVmIds.Contains($vm.id)) {
                $vmProperties = Get-ObjectPropertyValue -InputObject $vm -PropertyName @('properties')
                $unprotectedVmRows += [pscustomobject]@{
                    RunId            = $runId
                    SubscriptionId   = $subscription.Id
                    SubscriptionName = $subscription.Name
                    ResourceGroup    = Get-ResourceGroupFromId -ResourceId $vm.id
                    VmName           = $vm.name
                    Location         = $vm.location
                    PowerStateJson   = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $vmProperties -PropertyName @('instanceView')) -PropertyName @('statuses'))
                    TagsJson         = ConvertTo-CompactJson -Value $vm.tags
                    ResourceId       = $vm.id
                }
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                 = $runId
            SubscriptionId        = $subscription.Id
            SubscriptionName      = $subscription.Name
            TenantId              = $subscription.TenantId
            VaultCount            = $vaults.Count
            VmCount               = $vms.Count
            ProtectedItemCount    = $subscriptionProtectedItems.Count
            ProtectedVmCount      = @($protectedVmRows | Where-Object { $_.SubscriptionId -eq $subscription.Id }).Count
            UnprotectedVmCount    = @($unprotectedVmRows | Where-Object { $_.SubscriptionId -eq $subscription.Id }).Count
        }
    }

    Export-SmartAzureCsv -Name 'Azure_Backup_RecoveryServicesVaults' -Rows $vaultRows
    Export-SmartAzureCsv -Name 'Azure_Backup_VaultSettings' -Rows $vaultSettingRows
    Export-SmartAzureCsv -Name 'Azure_Backup_ProtectedItems' -Rows $protectedItemRows
    Export-SmartAzureCsv -Name 'Azure_Backup_ProtectedVMs' -Rows $protectedVmRows
    Export-SmartAzureCsv -Name 'Azure_Backup_UnprotectedVMs' -Rows $unprotectedVmRows
    Export-SmartAzureCsv -Name 'Azure_Backup_Summary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed Backup inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}
