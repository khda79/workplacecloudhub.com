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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDRBzVBZXShHWTc
# xYxa1Ukp6vMZbe05z56I5YQRGV2NxKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAPsJzYUEoxWOalHTV3
# rc46IVEkebchaR08576+49Z4YjANBgkqhkiG9w0BAQEFAASCAYAYcLgmEJzhaPbX
# dyBlRwsONoXeNsadRxVHujG5UvhuLvVW1zRcHXTaRN360gfX1hEj1gksdEVfc1MC
# NrSvBHFTQHnJCNoc4PHpzJMj/powTaQjizSz5YbB4lenTt4lVT/vjisJhkXJUWqf
# dHnzicsSK4da9gTXLgw+Thbzw2yQA0OUX/AHviPTEIQvbkBqwPh38bLBUNbxLh0z
# lFt0jb5qUO+EvZw/XeONVXqmEK4Kv8mXs+sanaLCRm5lYZNcy2RFlhvG3ZCGfF/e
# OVsaB4O4SU0KNsxtf1iI8c2rBKhdglDvCqK9NLzPATI6AtAXd7Ph6eTAb5XdU4Sm
# BPh1cp+DE1KK31zqrIId/ePCDdhYfCosicZlzEuj5b8y32Rxzs+QVUd1yEq+9m1Y
# 69wctA6ZFx+4pZjCsnTZQBSBsa2RQJfkvSdyNnMPz69+4IRXUNwa2Ct1TPYWwyL8
# RscAFJhFTGsVb6JxtKtQpSbSKK0SOhPHHUmyKP5FICSiHVbr4YI=
# SIG # End signature block
