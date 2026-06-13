<#
.SYNOPSIS
Exports Azure Monitor diagnostic settings for Azure Virtual Desktop resources.

.DESCRIPTION
Connects with Az PowerShell, enumerates Azure Virtual Desktop host pools, workspaces,
application groups, and scaling plans, then exports diagnostic settings and resources
without diagnostic settings.

The script is read-only. It does not create or update diagnostic settings.

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
    [string]$DiagnosticSettingsApiVersion = '2021-05-01-preview',
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
$resolvedOutputRoots = Resolve-SmartAvdOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'AzureVirtualDesktop\Diagnostics'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzureVirtualDesktop-Diagnostics-Inventory_{0}.log" -f $runId)

Import-SmartAvdCoreModule -StartPath $PSScriptRoot
Set-SmartAvdCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-AvdResourceTypeLabel {
    [CmdletBinding()]
    param([string]$ResourceType)

    switch -Regex ($ResourceType) {
        'hostPools$' { return 'HostPool' }
        'workspaces$' { return 'Workspace' }
        'applicationGroups$' { return 'ApplicationGroup' }
        'scalingPlans$' { return 'ScalingPlan' }
        default { return $ResourceType }
    }
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAvdLog -Message "Starting SmartAzureVirtualDesktop diagnostics inventory. RunId=$runId"
    Import-SmartAvdRequiredModule -Name Az.Accounts
    $context = Connect-SmartAvdCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw 'No Azure subscriptions were found for the current identity and filters.' }

    $diagnosticRows = @()
    $missingRows = @()
    $summaryRows = @()
    $avdResourceTypes = @(
        'Microsoft.DesktopVirtualization/hostPools',
        'Microsoft.DesktopVirtualization/workspaces',
        'Microsoft.DesktopVirtualization/applicationGroups',
        'Microsoft.DesktopVirtualization/scalingPlans'
    )

    foreach ($subscription in $subscriptions) {
        Write-SmartAvdLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null
        $subscriptionResourceCount = 0
        $subscriptionConfiguredCount = 0

        foreach ($resourceType in $avdResourceTypes) {
            $resources = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType $resourceType -ApiVersion $ResourceApiVersion)
            foreach ($resource in $resources) {
                $subscriptionResourceCount++
                $typeLabel = Get-AvdResourceTypeLabel -ResourceType $resourceType
                $resourceName = Get-SmartAvdResourceNameFromId -ResourceId $resource.id
                $uri = "https://management.azure.com{0}/providers/microsoft.insights/diagnosticSettings?api-version={1}" -f $resource.id, $DiagnosticSettingsApiVersion
                $settings = @(Invoke-SmartAvdArmGetPaged -Uri $uri -Operation "Diagnostic settings for $resourceName")

                if ($settings.Count -eq 0) {
                    $missingRows += [pscustomobject]@{
                        RunId             = $runId
                        SubscriptionId    = $subscription.Id
                        SubscriptionName  = $subscription.Name
                        ResourceGroupName = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                        ResourceType      = $typeLabel
                        ResourceName      = $resourceName
                        Location          = $resource.location
                        Finding           = 'NoDiagnosticSettings'
                        ResourceId        = $resource.id
                    }
                    continue
                }

                $subscriptionConfiguredCount++
                foreach ($setting in $settings) {
                    $properties = Get-SmartAvdObjectPropertyValue -InputObject $setting -PropertyName @('properties')
                    $logs = @(Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('logs'))
                    $metrics = @(Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('metrics'))
                    $enabledLogs = @($logs | Where-Object { (Get-SmartAvdObjectPropertyValue -InputObject $_ -PropertyName @('enabled')) -eq $true })
                    $enabledMetrics = @($metrics | Where-Object { (Get-SmartAvdObjectPropertyValue -InputObject $_ -PropertyName @('enabled')) -eq $true })

                    $diagnosticRows += [pscustomobject]@{
                        RunId                      = $runId
                        SubscriptionId             = $subscription.Id
                        SubscriptionName           = $subscription.Name
                        ResourceGroupName          = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                        ResourceType               = $typeLabel
                        ResourceName               = $resourceName
                        DiagnosticSettingName      = Get-SmartAvdResourceNameFromId -ResourceId $setting.id
                        EnabledLogCount            = $enabledLogs.Count
                        EnabledMetricCount         = $enabledMetrics.Count
                        EnabledLogCategories       = (@($enabledLogs | ForEach-Object { Get-SmartAvdObjectPropertyValue -InputObject $_ -PropertyName @('category', 'categoryGroup') }) -join ';')
                        EnabledMetricCategories    = (@($enabledMetrics | ForEach-Object { Get-SmartAvdObjectPropertyValue -InputObject $_ -PropertyName @('category') }) -join ';')
                        WorkspaceId                = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('workspaceId')
                        StorageAccountId           = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('storageAccountId')
                        EventHubAuthorizationRuleId = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('eventHubAuthorizationRuleId')
                        EventHubName               = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('eventHubName')
                        MarketplacePartnerId       = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('marketplacePartnerId')
                        LogAnalyticsDestinationType = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('logAnalyticsDestinationType')
                        LogsJson                   = ConvertTo-SmartAvdCompactJson -Value $logs
                        MetricsJson                = ConvertTo-SmartAvdCompactJson -Value $metrics
                        ResourceId                 = $resource.id
                        DiagnosticSettingId        = $setting.id
                    }
                }
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                         = $runId
            SubscriptionId                = $subscription.Id
            SubscriptionName              = $subscription.Name
            TenantId                      = $subscription.TenantId
            AvdResourceCount              = $subscriptionResourceCount
            ResourceWithDiagnosticsCount  = $subscriptionConfiguredCount
            ResourceWithoutDiagnosticsCount = $subscriptionResourceCount - $subscriptionConfiguredCount
            DiagnosticsCoveragePercent    = if ($subscriptionResourceCount -gt 0) { [math]::Round(($subscriptionConfiguredCount / $subscriptionResourceCount) * 100, 2) } else { $null }
        }
    }

    Export-SmartAvdCsv -Name 'AVD_DiagnosticSettings' -Rows $diagnosticRows
    Export-SmartAvdCsv -Name 'AVD_DiagnosticsMissing' -Rows $missingRows
    Export-SmartAvdCsv -Name 'AVD_DiagnosticsSummary' -Rows $summaryRows

    Write-SmartAvdLog -Level SUCCESS -Message ("Completed diagnostics inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    if (Get-Command -Name Write-SmartAvdLog -ErrorAction SilentlyContinue) { Write-SmartAvdLog -Level ERROR -Message $_.Exception.Message }
    throw
}
