<#
.SYNOPSIS
Exports Azure RBAC assignments and role definitions.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
role assignments, custom role definitions, privileged assignments, and a subscription-level
RBAC summary.

The script is read-only. It does not modify Azure role assignments.

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

.PARAMETER IncludeClassicAdministrators
Attempts to export classic subscription administrators when supported by the installed Az modules.

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
    [switch]$IncludeClassicAdministrators
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
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\RBAC'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-RBAC-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-ScopeType {
    [CmdletBinding()]
    param([string]$Scope)

    if ([string]::IsNullOrWhiteSpace($Scope)) { return 'Unknown' }
    if ($Scope -match '^/subscriptions/[^/]+$') { return 'Subscription' }
    if ($Scope -match '^/subscriptions/[^/]+/resourceGroups/[^/]+$') { return 'ResourceGroup' }
    if ($Scope -match '^/providers/Microsoft.Management/managementGroups/') { return 'ManagementGroup' }
    if ($Scope -match '^/subscriptions/[^/]+/resourceGroups/[^/]+/providers/') { return 'Resource' }
    return 'Other'
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure RBAC inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts
    Import-RequiredModule -Name Az.Resources

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

    $roleAssignmentRows = @()
    $customRoleRows = @()
    $privilegedRows = @()
    $classicAdminRows = @()
    $summaryRows = @()
    $privilegedRoles = @(
        'Owner',
        'Contributor',
        'User Access Administrator',
        'Role Based Access Control Administrator'
    )

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $assignments = @(Invoke-SafeInventoryBlock -Name "Role assignments for $($subscription.Name)" -ScriptBlock {
            Get-AzRoleAssignment -Scope ("/subscriptions/{0}" -f $subscription.Id)
        })

        foreach ($assignment in $assignments) {
            $scopeType = Get-ScopeType -Scope $assignment.Scope
            $isPrivileged = $assignment.RoleDefinitionName -in $privilegedRoles
            $row = [pscustomobject]@{
                RunId              = $runId
                SubscriptionId     = $subscription.Id
                SubscriptionName   = $subscription.Name
                TenantId           = $subscription.TenantId
                Scope              = $assignment.Scope
                ScopeType          = $scopeType
                RoleDefinitionName = $assignment.RoleDefinitionName
                RoleDefinitionId   = $assignment.RoleDefinitionId
                ObjectId           = $assignment.ObjectId
                ObjectType         = $assignment.ObjectType
                DisplayName        = $assignment.DisplayName
                SignInName         = $assignment.SignInName
                CanDelegate        = $assignment.CanDelegate
                Condition          = $assignment.Condition
                ConditionVersion   = $assignment.ConditionVersion
                Description        = $assignment.Description
                IsPrivilegedRole   = $isPrivileged
            }
            $roleAssignmentRows += $row
            if ($isPrivileged) { $privilegedRows += $row }
        }

        $customRoles = @(Invoke-SafeInventoryBlock -Name "Custom role definitions for $($subscription.Name)" -ScriptBlock {
            Get-AzRoleDefinition | Where-Object { $_.IsCustom -eq $true }
        })

        foreach ($role in $customRoles) {
            $customRoleRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                Name             = $role.Name
                Id               = $role.Id
                IsCustom         = $role.IsCustom
                Description      = $role.Description
                ActionsJson      = ConvertTo-CompactJson -Value $role.Actions
                NotActionsJson   = ConvertTo-CompactJson -Value $role.NotActions
                DataActionsJson  = ConvertTo-CompactJson -Value $role.DataActions
                NotDataActionsJson = ConvertTo-CompactJson -Value $role.NotDataActions
                AssignableScopesJson = ConvertTo-CompactJson -Value $role.AssignableScopes
            }
        }

        if ($IncludeClassicAdministrators) {
            $classicAdmins = @(Invoke-SafeInventoryBlock -Name "Classic administrators for $($subscription.Name)" -ScriptBlock {
                Get-AzRoleAssignment -IncludeClassicAdministrators
            })

            foreach ($admin in $classicAdmins) {
                $classicAdminRows += [pscustomobject]@{
                    RunId            = $runId
                    SubscriptionId   = $subscription.Id
                    SubscriptionName = $subscription.Name
                    Scope            = $admin.Scope
                    RoleDefinitionName = $admin.RoleDefinitionName
                    ObjectId         = $admin.ObjectId
                    ObjectType       = $admin.ObjectType
                    DisplayName      = $admin.DisplayName
                    SignInName       = $admin.SignInName
                }
            }
        }

        $ownerCount = @($assignments | Where-Object { $_.RoleDefinitionName -eq 'Owner' }).Count
        $privilegedCount = @($assignments | Where-Object { $_.RoleDefinitionName -in $privilegedRoles }).Count
        $groupAssignmentCount = @($assignments | Where-Object { $_.ObjectType -eq 'Group' }).Count
        $servicePrincipalAssignmentCount = @($assignments | Where-Object { $_.ObjectType -eq 'ServicePrincipal' }).Count

        $summaryRows += [pscustomobject]@{
            RunId                           = $runId
            SubscriptionId                  = $subscription.Id
            SubscriptionName                = $subscription.Name
            TenantId                        = $subscription.TenantId
            RoleAssignmentCount             = $assignments.Count
            PrivilegedAssignmentCount       = $privilegedCount
            OwnerAssignmentCount            = $ownerCount
            GroupAssignmentCount            = $groupAssignmentCount
            ServicePrincipalAssignmentCount = $servicePrincipalAssignmentCount
            CustomRoleCount                 = $customRoles.Count
        }
    }

    Export-SmartAzureCsv -Name 'Azure_RBAC_RoleAssignments' -Rows $roleAssignmentRows
    Export-SmartAzureCsv -Name 'Azure_RBAC_PrivilegedAssignments' -Rows $privilegedRows
    Export-SmartAzureCsv -Name 'Azure_RBAC_CustomRoles' -Rows $customRoleRows
    Export-SmartAzureCsv -Name 'Azure_RBAC_ClassicAdministrators' -Rows $classicAdminRows
    Export-SmartAzureCsv -Name 'Azure_RBAC_Summary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed RBAC inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}
