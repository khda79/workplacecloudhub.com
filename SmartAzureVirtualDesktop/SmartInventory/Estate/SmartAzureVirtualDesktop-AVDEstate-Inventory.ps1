<#
.SYNOPSIS
Exports an Azure Virtual Desktop estate inventory.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
Azure Virtual Desktop host pools, workspaces, application groups, applications, desktops,
session hosts, scaling plans, private endpoint connections, and subscription summaries.

The script is read-only. It does not modify Azure resources and does not export AVD
registration token values.

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
$resolvedOutputRoots = Resolve-SmartAvdOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'AzureVirtualDesktop\Estate'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzureVirtualDesktop-AVDEstate-Inventory_{0}.log" -f $runId)

Import-SmartAvdCoreModule -StartPath $PSScriptRoot
Set-SmartAvdCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAvdLog -Message "Starting SmartAzureVirtualDesktop AVD estate inventory. RunId=$runId"
    Import-SmartAvdRequiredModule -Name Az.Accounts
    $context = Connect-SmartAvdCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw 'No Azure subscriptions were found for the current identity and filters.' }

    $hostPoolRows = @()
    $workspaceRows = @()
    $applicationGroupRows = @()
    $applicationRows = @()
    $desktopRows = @()
    $sessionHostRows = @()
    $scalingPlanRows = @()
    $privateEndpointRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAvdLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $hostPools = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.DesktopVirtualization/hostPools' -ApiVersion $ResourceApiVersion)
        $workspaces = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.DesktopVirtualization/workspaces' -ApiVersion $ResourceApiVersion)
        $applicationGroups = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.DesktopVirtualization/applicationGroups' -ApiVersion $ResourceApiVersion)
        $scalingPlans = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.DesktopVirtualization/scalingPlans' -ApiVersion $ResourceApiVersion)

        foreach ($resource in $hostPools) {
            $details = Get-SmartAvdResourceById -ResourceId $resource.id -ApiVersion $AvdApiVersion -Operation 'Host pool details'
            if ($null -eq $details) { $details = $resource }
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            $agentUpdate = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('agentUpdate')
            $registrationInfo = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('registrationInfo')
            $hostPoolName = Get-SmartAvdResourceNameFromId -ResourceId $resource.id

            $hostPoolRows += [pscustomobject]@{
                RunId                            = $runId
                SubscriptionId                   = $subscription.Id
                SubscriptionName                 = $subscription.Name
                TenantId                         = $subscription.TenantId
                ResourceGroupName                = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                HostPoolName                     = $hostPoolName
                Location                         = $details.location
                HostPoolType                     = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('hostPoolType')
                PersonalDesktopAssignmentType    = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('personalDesktopAssignmentType')
                PreferredAppGroupType            = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('preferredAppGroupType')
                LoadBalancerType                 = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('loadBalancerType')
                MaxSessionLimit                  = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('maxSessionLimit')
                StartVMOnConnect                 = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('startVMOnConnect'))
                ValidationEnvironment            = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('validationEnvironment'))
                PublicNetworkAccess              = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('publicNetworkAccess')
                Ring                             = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('ring')
                FriendlyName                     = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('friendlyName')
                Description                      = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('description')
                CustomRdpProperty                = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('customRdpProperty')
                AgentUpdateType                  = Get-SmartAvdObjectPropertyValue -InputObject $agentUpdate -PropertyName @('type')
                AgentUpdateUseSessionHostTimezone = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $agentUpdate -PropertyName @('useSessionHostLocalTime'))
                RegistrationInfoExpirationTime   = Get-SmartAvdObjectPropertyValue -InputObject $registrationInfo -PropertyName @('expirationTime')
                SkuName                          = Get-SmartAvdNestedPropertyValue -InputObject $details -Path @('sku', 'name')
                TagsJson                         = ConvertTo-SmartAvdCompactJson -Value $details.tags
                ResourceId                       = $resource.id
            }

            $sessionHosts = @(Get-SmartAvdChildResources -ParentResourceId $resource.id -ChildPath 'sessionHosts' -ApiVersion $AvdApiVersion -Operation "Session hosts for $hostPoolName")
            foreach ($sessionHost in $sessionHosts) {
                $sessionHostProperties = Get-SmartAvdObjectPropertyValue -InputObject $sessionHost -PropertyName @('properties')
                $sessionHostRows += [pscustomobject]@{
                    RunId              = $runId
                    SubscriptionId     = $subscription.Id
                    SubscriptionName   = $subscription.Name
                    ResourceGroupName  = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                    HostPoolName       = $hostPoolName
                    SessionHostName    = Get-SmartAvdResourceNameFromId -ResourceId $sessionHost.id
                    Status             = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('status')
                    UpdateState        = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('updateState')
                    AllowNewSession    = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('allowNewSession'))
                    AssignedUser       = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('assignedUser')
                    Sessions           = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('sessions')
                    LastHeartBeat      = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('lastHeartBeat')
                    AgentVersion       = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('agentVersion')
                    OSVersion          = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('osVersion')
                    StatusTimestamp    = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('statusTimestamp')
                    VmResourceId       = Get-SmartAvdObjectPropertyValue -InputObject $sessionHostProperties -PropertyName @('resourceId')
                    ResourceId         = $sessionHost.id
                }
            }

            $privateConnections = @(Get-SmartAvdChildResources -ParentResourceId $resource.id -ChildPath 'privateEndpointConnections' -ApiVersion $AvdApiVersion -Operation "Private endpoint connections for $hostPoolName")
            foreach ($connection in $privateConnections) {
                $connectionProperties = Get-SmartAvdObjectPropertyValue -InputObject $connection -PropertyName @('properties')
                $state = Get-SmartAvdObjectPropertyValue -InputObject $connectionProperties -PropertyName @('privateLinkServiceConnectionState')
                $privateEndpointRows += [pscustomobject]@{
                    RunId              = $runId
                    SubscriptionId     = $subscription.Id
                    SubscriptionName   = $subscription.Name
                    ParentType         = 'HostPool'
                    ParentName         = $hostPoolName
                    ConnectionName     = Get-SmartAvdResourceNameFromId -ResourceId $connection.id
                    ProvisioningState  = Get-SmartAvdObjectPropertyValue -InputObject $connectionProperties -PropertyName @('provisioningState')
                    PrivateEndpointId  = Get-SmartAvdNestedPropertyValue -InputObject $connectionProperties -Path @('privateEndpoint', 'id')
                    ConnectionStatus   = Get-SmartAvdObjectPropertyValue -InputObject $state -PropertyName @('status')
                    Description        = Get-SmartAvdObjectPropertyValue -InputObject $state -PropertyName @('description')
                    ActionsRequired    = Get-SmartAvdObjectPropertyValue -InputObject $state -PropertyName @('actionsRequired')
                    ResourceId         = $connection.id
                }
            }
        }

        foreach ($resource in $workspaces) {
            $details = Get-SmartAvdResourceById -ResourceId $resource.id -ApiVersion $AvdApiVersion -Operation 'Workspace details'
            if ($null -eq $details) { $details = $resource }
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            $workspaceName = Get-SmartAvdResourceNameFromId -ResourceId $resource.id
            $appGroupReferences = @(Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('applicationGroupReferences'))
            $workspaceRows += [pscustomobject]@{
                RunId                    = $runId
                SubscriptionId           = $subscription.Id
                SubscriptionName         = $subscription.Name
                TenantId                 = $subscription.TenantId
                ResourceGroupName        = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                WorkspaceName            = $workspaceName
                Location                 = $details.location
                FriendlyName             = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('friendlyName')
                Description              = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('description')
                PublicNetworkAccess      = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('publicNetworkAccess')
                ApplicationGroupCount    = $appGroupReferences.Count
                ApplicationGroupRefsJson = ConvertTo-SmartAvdCompactJson -Value $appGroupReferences
                TagsJson                 = ConvertTo-SmartAvdCompactJson -Value $details.tags
                ResourceId               = $resource.id
            }

            $privateConnections = @(Get-SmartAvdChildResources -ParentResourceId $resource.id -ChildPath 'privateEndpointConnections' -ApiVersion $AvdApiVersion -Operation "Private endpoint connections for $workspaceName")
            foreach ($connection in $privateConnections) {
                $connectionProperties = Get-SmartAvdObjectPropertyValue -InputObject $connection -PropertyName @('properties')
                $state = Get-SmartAvdObjectPropertyValue -InputObject $connectionProperties -PropertyName @('privateLinkServiceConnectionState')
                $privateEndpointRows += [pscustomobject]@{
                    RunId              = $runId
                    SubscriptionId     = $subscription.Id
                    SubscriptionName   = $subscription.Name
                    ParentType         = 'Workspace'
                    ParentName         = $workspaceName
                    ConnectionName     = Get-SmartAvdResourceNameFromId -ResourceId $connection.id
                    ProvisioningState  = Get-SmartAvdObjectPropertyValue -InputObject $connectionProperties -PropertyName @('provisioningState')
                    PrivateEndpointId  = Get-SmartAvdNestedPropertyValue -InputObject $connectionProperties -Path @('privateEndpoint', 'id')
                    ConnectionStatus   = Get-SmartAvdObjectPropertyValue -InputObject $state -PropertyName @('status')
                    Description        = Get-SmartAvdObjectPropertyValue -InputObject $state -PropertyName @('description')
                    ActionsRequired    = Get-SmartAvdObjectPropertyValue -InputObject $state -PropertyName @('actionsRequired')
                    ResourceId         = $connection.id
                }
            }
        }

        foreach ($resource in $applicationGroups) {
            $details = Get-SmartAvdResourceById -ResourceId $resource.id -ApiVersion $AvdApiVersion -Operation 'Application group details'
            if ($null -eq $details) { $details = $resource }
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            $applicationGroupName = Get-SmartAvdResourceNameFromId -ResourceId $resource.id
            $hostPoolArmPath = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('hostPoolArmPath')
            $applicationGroupRows += [pscustomobject]@{
                RunId                = $runId
                SubscriptionId       = $subscription.Id
                SubscriptionName     = $subscription.Name
                ResourceGroupName    = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                ApplicationGroupName = $applicationGroupName
                Location             = $details.location
                ApplicationGroupType = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('applicationGroupType')
                HostPoolName         = Get-SmartAvdResourceNameFromId -ResourceId $hostPoolArmPath
                HostPoolArmPath      = $hostPoolArmPath
                FriendlyName         = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('friendlyName')
                Description          = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('description')
                TagsJson             = ConvertTo-SmartAvdCompactJson -Value $details.tags
                ResourceId           = $resource.id
            }

            $applications = @(Get-SmartAvdChildResources -ParentResourceId $resource.id -ChildPath 'applications' -ApiVersion $AvdApiVersion -Operation "Applications for $applicationGroupName")
            foreach ($application in $applications) {
                $applicationProperties = Get-SmartAvdObjectPropertyValue -InputObject $application -PropertyName @('properties')
                $applicationRows += [pscustomobject]@{
                    RunId                = $runId
                    SubscriptionId       = $subscription.Id
                    SubscriptionName     = $subscription.Name
                    ResourceGroupName    = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                    ApplicationGroupName = $applicationGroupName
                    ApplicationName      = Get-SmartAvdResourceNameFromId -ResourceId $application.id
                    FriendlyName         = Get-SmartAvdObjectPropertyValue -InputObject $applicationProperties -PropertyName @('friendlyName')
                    Description          = Get-SmartAvdObjectPropertyValue -InputObject $applicationProperties -PropertyName @('description')
                    FilePath             = Get-SmartAvdObjectPropertyValue -InputObject $applicationProperties -PropertyName @('filePath')
                    CommandLineSetting   = Get-SmartAvdObjectPropertyValue -InputObject $applicationProperties -PropertyName @('commandLineSetting')
                    ShowInPortal         = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $applicationProperties -PropertyName @('showInPortal'))
                    IconPath             = Get-SmartAvdObjectPropertyValue -InputObject $applicationProperties -PropertyName @('iconPath')
                    IconIndex            = Get-SmartAvdObjectPropertyValue -InputObject $applicationProperties -PropertyName @('iconIndex')
                    ResourceId           = $application.id
                }
            }

            $desktops = @(Get-SmartAvdChildResources -ParentResourceId $resource.id -ChildPath 'desktops' -ApiVersion $AvdApiVersion -Operation "Desktops for $applicationGroupName")
            foreach ($desktop in $desktops) {
                $desktopProperties = Get-SmartAvdObjectPropertyValue -InputObject $desktop -PropertyName @('properties')
                $desktopRows += [pscustomobject]@{
                    RunId                = $runId
                    SubscriptionId       = $subscription.Id
                    SubscriptionName     = $subscription.Name
                    ResourceGroupName    = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                    ApplicationGroupName = $applicationGroupName
                    DesktopName          = Get-SmartAvdResourceNameFromId -ResourceId $desktop.id
                    FriendlyName         = Get-SmartAvdObjectPropertyValue -InputObject $desktopProperties -PropertyName @('friendlyName')
                    Description          = Get-SmartAvdObjectPropertyValue -InputObject $desktopProperties -PropertyName @('description')
                    ResourceId           = $desktop.id
                }
            }
        }

        foreach ($resource in $scalingPlans) {
            $details = Get-SmartAvdResourceById -ResourceId $resource.id -ApiVersion $AvdApiVersion -Operation 'Scaling plan details'
            if ($null -eq $details) { $details = $resource }
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            $hostPoolReferences = @(Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('hostPoolReferences'))
            $scalingPlanRows += [pscustomobject]@{
                RunId                 = $runId
                SubscriptionId        = $subscription.Id
                SubscriptionName      = $subscription.Name
                ResourceGroupName     = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                ScalingPlanName       = Get-SmartAvdResourceNameFromId -ResourceId $resource.id
                Location              = $details.location
                HostPoolType          = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('hostPoolType')
                TimeZone              = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('timeZone')
                ExclusionTag          = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('exclusionTag')
                FriendlyName          = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('friendlyName')
                Description           = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('description')
                HostPoolReferenceCount = $hostPoolReferences.Count
                HostPoolReferencesJson = ConvertTo-SmartAvdCompactJson -Value $hostPoolReferences
                TagsJson              = ConvertTo-SmartAvdCompactJson -Value $details.tags
                ResourceId            = $resource.id
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                 = $runId
            SubscriptionId        = $subscription.Id
            SubscriptionName      = $subscription.Name
            TenantId              = $subscription.TenantId
            HostPoolCount         = $hostPools.Count
            WorkspaceCount        = $workspaces.Count
            ApplicationGroupCount = $applicationGroups.Count
            ScalingPlanCount      = $scalingPlans.Count
            SessionHostCount      = @($sessionHostRows | Where-Object { $_.SubscriptionId -eq $subscription.Id }).Count
            ApplicationCount      = @($applicationRows | Where-Object { $_.SubscriptionId -eq $subscription.Id }).Count
            DesktopCount          = @($desktopRows | Where-Object { $_.SubscriptionId -eq $subscription.Id }).Count
        }
    }

    Export-SmartAvdCsv -Name 'AVD_HostPools' -Rows $hostPoolRows
    Export-SmartAvdCsv -Name 'AVD_Workspaces' -Rows $workspaceRows
    Export-SmartAvdCsv -Name 'AVD_ApplicationGroups' -Rows $applicationGroupRows
    Export-SmartAvdCsv -Name 'AVD_Applications' -Rows $applicationRows
    Export-SmartAvdCsv -Name 'AVD_Desktops' -Rows $desktopRows
    Export-SmartAvdCsv -Name 'AVD_SessionHosts' -Rows $sessionHostRows
    Export-SmartAvdCsv -Name 'AVD_ScalingPlans' -Rows $scalingPlanRows
    Export-SmartAvdCsv -Name 'AVD_PrivateEndpointConnections' -Rows $privateEndpointRows
    Export-SmartAvdCsv -Name 'AVD_EstateSummary' -Rows $summaryRows

    Write-SmartAvdLog -Level SUCCESS -Message ("Completed AVD estate inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    if (Get-Command -Name Write-SmartAvdLog -ErrorAction SilentlyContinue) {
        Write-SmartAvdLog -Level ERROR -Message $_.Exception.Message
    }
    throw
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDRBzVBZXShHWTc
# xYxa1Ukp6vMZbe05z56I5YQRGV2NxKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAPsJzYUEoxWOalHTV3rc46IVEkebchaR08576+49Z4YjANBgkqhkiG9w0B
# AQEFAASCAYBkOWVGymsdmq94laZYh4GaWdYCyWO0BHgeE0O5QY+Eesig7kQ8cmg7
# dHv8ft1N2XDXsAS/EZ2dDBUdqAVXZWNJYlpdUqeW1me8VV0ngqgLTuD9u2HlIJGl
# wSTHx60bTabNlqVKG347up7cmK6hmpNqEmZGqpRH6JIX/fncGw80V6M3wfJ/kEkS
# cGq/9wA6SOIvj66mUzTF3gcd8QKqktRv2UNNBuSlfvhhiIUJrZ+f6IluTO3JkuhS
# jueMUBTN3WtD9U602D16UoK0CdYdsnI7wPfMMnUM4nT3zCuCXNhL/D2r2FXiem4K
# y+5LbxC2V0ez/10IaL/l5aBWmot9aDQythesRbIUwWvwvGxpW0prIiL/5WehE2N8
# Vy/W2m+3AdLweiMXCINdysWaS0NZLwF6fNEBuY2sSKIAJiw9Orj3ocriZX/N9QXa
# wH021/8GR+guMmTFpKHqqrn+KTNUAUVrunC7HOxbMrnoM9h8+ROVqucAYWdYE68Q
# pbLKtXGNweI=
# SIG # End signature block
