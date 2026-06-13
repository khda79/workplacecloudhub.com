<#
.SYNOPSIS
Exports Azure network exposure signals.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
public IP addresses, inbound NSG allow rules with Internet-like source prefixes, load balancers,
application gateways, private endpoints, and a subscription-level summary.

The script is read-only. It does not modify Azure networking resources.

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

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
Requires: Az.Accounts, Az.Network
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$TenantId,
    [string[]]$SubscriptionId,
    [string]$OutputRoot,
    [string]$LatestOutputRoot,
    [switch]$Connect,
    [switch]$UseDeviceCode
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
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\Network'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-NetworkExposure-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Join-RuleValues {
    [CmdletBinding()]
    param(
        [AllowNull()]$SingleValue,
        [AllowNull()]$MultipleValue
    )

    $values = @()
    if ($null -ne $MultipleValue) { $values += @($MultipleValue) }
    if ($null -ne $SingleValue) { $values += @($SingleValue) }
    return (($values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique) -join ';')
}

function Test-InternetSourcePrefix {
    [CmdletBinding()]
    param([string]$Source)

    if ([string]::IsNullOrWhiteSpace($Source)) { return $false }

    $parts = $Source -split ';'
    foreach ($part in $parts) {
        $value = $part.Trim()
        if ($value -in @('*', 'Internet', '0.0.0.0/0', '::/0', 'Any')) {
            return $true
        }
    }

    return $false
}

function Test-AdminPort {
    [CmdletBinding()]
    param([string]$DestinationPort)

    if ([string]::IsNullOrWhiteSpace($DestinationPort)) { return $false }
    if ($DestinationPort -in @('*', '22', '3389', '5985', '5986')) { return $true }

    foreach ($part in ($DestinationPort -split ';')) {
        $value = $part.Trim()
        if ($value -in @('22', '3389', '5985', '5986')) { return $true }
        if ($value -match '^(\d+)-(\d+)$') {
            $start = [int]$matches[1]
            $end = [int]$matches[2]
            foreach ($port in @(22, 3389, 5985, 5986)) {
                if ($port -ge $start -and $port -le $end) { return $true }
            }
        }
    }

    return $false
}

function Get-ResourceNameFromId {
    [CmdletBinding()]
    param([string]$ResourceId)

    if ([string]::IsNullOrWhiteSpace($ResourceId)) { return '' }
    return ($ResourceId.TrimEnd('/') -split '/')[-1]
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure network exposure inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts
    Import-RequiredModule -Name Az.Network

    $context = Connect-SmartAzureCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }

    if ($subscriptions.Count -eq 0) {
        throw "No Azure subscriptions were found for the current identity and filters."
    }

    $publicIpRows = @()
    $internetInboundRuleRows = @()
    $loadBalancerRows = @()
    $applicationGatewayRows = @()
    $privateEndpointRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $publicIps = @(Invoke-SafeInventoryBlock -Name "Public IP inventory for $($subscription.Name)" -ScriptBlock { Get-AzPublicIpAddress })
        foreach ($publicIp in $publicIps) {
            $ipConfig = Get-ObjectPropertyValue -InputObject $publicIp -PropertyName @('IpConfiguration')
            $attachedId = if ($ipConfig) { $ipConfig.Id } else { '' }
            $publicIpRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                ResourceGroupName = $publicIp.ResourceGroupName
                Name             = $publicIp.Name
                Location         = $publicIp.Location
                IpAddress        = $publicIp.IpAddress
                AllocationMethod = $publicIp.PublicIpAllocationMethod
                IpVersion        = $publicIp.PublicIpAddressVersion
                SkuName          = if ($publicIp.Sku) { $publicIp.Sku.Name } else { '' }
                AttachedResourceName = Get-ResourceNameFromId -ResourceId $attachedId
                AttachedResourceId = $attachedId
                DnsName          = if ($publicIp.DnsSettings) { $publicIp.DnsSettings.DomainNameLabel } else { '' }
                Fqdn             = if ($publicIp.DnsSettings) { $publicIp.DnsSettings.Fqdn } else { '' }
                TagsJson         = ConvertTo-CompactJson -Value $publicIp.Tag
                ResourceId       = $publicIp.Id
            }
        }

        $nsgs = @(Invoke-SafeInventoryBlock -Name "NSG inventory for $($subscription.Name)" -ScriptBlock { Get-AzNetworkSecurityGroup })
        foreach ($nsg in $nsgs) {
            $allRules = @()
            if ($nsg.SecurityRules) { $allRules += @($nsg.SecurityRules) }
            if ($nsg.DefaultSecurityRules) { $allRules += @($nsg.DefaultSecurityRules) }

            foreach ($rule in $allRules) {
                $source = Join-RuleValues -SingleValue $rule.SourceAddressPrefix -MultipleValue $rule.SourceAddressPrefixes
                $destinationPort = Join-RuleValues -SingleValue $rule.DestinationPortRange -MultipleValue $rule.DestinationPortRanges
                if ($rule.Direction -ne 'Inbound' -or $rule.Access -ne 'Allow') { continue }
                if (-not (Test-InternetSourcePrefix -Source $source)) { continue }

                $internetInboundRuleRows += [pscustomobject]@{
                    RunId              = $runId
                    SubscriptionId     = $subscription.Id
                    SubscriptionName   = $subscription.Name
                    ResourceGroupName  = $nsg.ResourceGroupName
                    NetworkSecurityGroupName = $nsg.Name
                    RuleName           = $rule.Name
                    Priority           = $rule.Priority
                    Direction          = $rule.Direction
                    Access             = $rule.Access
                    Protocol           = $rule.Protocol
                    SourceAddress      = $source
                    SourcePort         = Join-RuleValues -SingleValue $rule.SourcePortRange -MultipleValue $rule.SourcePortRanges
                    DestinationAddress = Join-RuleValues -SingleValue $rule.DestinationAddressPrefix -MultipleValue $rule.DestinationAddressPrefixes
                    DestinationPort    = $destinationPort
                    IsAdminPort        = Test-AdminPort -DestinationPort $destinationPort
                    Description        = $rule.Description
                    IsDefaultRule      = $rule.Name -like 'Allow*' -and $rule.Priority -ge 65000
                    NetworkSecurityGroupId = $nsg.Id
                }
            }
        }

        $loadBalancers = @(Invoke-SafeInventoryBlock -Name "Load balancer inventory for $($subscription.Name)" -ScriptBlock { Get-AzLoadBalancer })
        foreach ($loadBalancer in $loadBalancers) {
            $loadBalancerRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                ResourceGroupName = $loadBalancer.ResourceGroupName
                Name             = $loadBalancer.Name
                Location         = $loadBalancer.Location
                SkuName          = if ($loadBalancer.Sku) { $loadBalancer.Sku.Name } else { '' }
                FrontendIpConfigCount = @($loadBalancer.FrontendIpConfigurations).Count
                BackendPoolCount = @($loadBalancer.BackendAddressPools).Count
                LoadBalancingRuleCount = @($loadBalancer.LoadBalancingRules).Count
                InboundNatRuleCount = @($loadBalancer.InboundNatRules).Count
                PublicFrontendIdsJson = ConvertTo-CompactJson -Value @($loadBalancer.FrontendIpConfigurations | Where-Object { $_.PublicIpAddress } | ForEach-Object { $_.PublicIpAddress.Id })
                ResourceId       = $loadBalancer.Id
            }
        }

        $applicationGateways = @(Invoke-SafeInventoryBlock -Name "Application gateway inventory for $($subscription.Name)" -ScriptBlock { Get-AzApplicationGateway })
        foreach ($applicationGateway in $applicationGateways) {
            $applicationGatewayRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                ResourceGroupName = $applicationGateway.ResourceGroupName
                Name             = $applicationGateway.Name
                Location         = $applicationGateway.Location
                SkuName          = if ($applicationGateway.Sku) { $applicationGateway.Sku.Name } else { '' }
                SkuTier          = if ($applicationGateway.Sku) { $applicationGateway.Sku.Tier } else { '' }
                FrontendIpConfigCount = @($applicationGateway.FrontendIPConfigurations).Count
                PublicFrontendIdsJson = ConvertTo-CompactJson -Value @($applicationGateway.FrontendIPConfigurations | Where-Object { $_.PublicIPAddress } | ForEach-Object { $_.PublicIPAddress.Id })
                ListenerCount    = @($applicationGateway.HttpListeners).Count
                RuleCount        = @($applicationGateway.RequestRoutingRules).Count
                WafEnabled       = if ($applicationGateway.WebApplicationFirewallConfiguration) { $applicationGateway.WebApplicationFirewallConfiguration.Enabled } else { $null }
                FirewallPolicyId = if ($applicationGateway.FirewallPolicy) { $applicationGateway.FirewallPolicy.Id } else { '' }
                ResourceId       = $applicationGateway.Id
            }
        }

        $privateEndpoints = @(Invoke-SafeInventoryBlock -Name "Private endpoint inventory for $($subscription.Name)" -ScriptBlock { Get-AzPrivateEndpoint })
        foreach ($privateEndpoint in $privateEndpoints) {
            $privateEndpointRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                ResourceGroupName = $privateEndpoint.ResourceGroupName
                Name             = $privateEndpoint.Name
                Location         = $privateEndpoint.Location
                SubnetId         = if ($privateEndpoint.Subnet) { $privateEndpoint.Subnet.Id } else { '' }
                PrivateLinkServiceConnectionsJson = ConvertTo-CompactJson -Value $privateEndpoint.PrivateLinkServiceConnections
                ManualPrivateLinkServiceConnectionsJson = ConvertTo-CompactJson -Value $privateEndpoint.ManualPrivateLinkServiceConnections
                NetworkInterfacesJson = ConvertTo-CompactJson -Value $privateEndpoint.NetworkInterfaces
                ResourceId       = $privateEndpoint.Id
            }
        }

        $adminExposureCount = @($internetInboundRuleRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.IsAdminPort -eq $true }).Count
        $summaryRows += [pscustomobject]@{
            RunId                       = $runId
            SubscriptionId              = $subscription.Id
            SubscriptionName            = $subscription.Name
            TenantId                    = $subscription.TenantId
            PublicIpCount               = $publicIps.Count
            InternetInboundNsgRuleCount = @($internetInboundRuleRows | Where-Object { $_.SubscriptionId -eq $subscription.Id }).Count
            AdminPortExposureRuleCount  = $adminExposureCount
            LoadBalancerCount           = $loadBalancers.Count
            ApplicationGatewayCount     = $applicationGateways.Count
            PrivateEndpointCount        = $privateEndpoints.Count
        }
    }

    Export-SmartAzureCsv -Name 'Azure_Network_PublicIPs' -Rows $publicIpRows
    Export-SmartAzureCsv -Name 'Azure_Network_InternetInboundNSGRules' -Rows $internetInboundRuleRows
    Export-SmartAzureCsv -Name 'Azure_Network_LoadBalancers' -Rows $loadBalancerRows
    Export-SmartAzureCsv -Name 'Azure_Network_ApplicationGateways' -Rows $applicationGatewayRows
    Export-SmartAzureCsv -Name 'Azure_Network_PrivateEndpoints' -Rows $privateEndpointRows
    Export-SmartAzureCsv -Name 'Azure_Network_ExposureSummary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed network exposure inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}
