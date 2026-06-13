<#
.SYNOPSIS
Exports a Citrix Virtual Apps and Desktops on-premises delivery inventory.

.DESCRIPTION
Uses the local Citrix PowerShell SDK to export catalogs, delivery groups, published
resources, access policies, entitlement policies, assignment policies, tags, and reboot
schedules from an on-premises CVAD site.

The script is read-only and does not modify Citrix configuration.

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
Requires: Citrix PowerShell SDK
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$AdminAddress,
    [string]$OutputRoot,
    [string]$LatestOutputRoot,
    [int]$MaxRecordCount = 0
)

$ErrorActionPreference = 'Stop'
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path -Path $d -ChildPath 'Config\SmartCitrix-TenantContext.ps1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartCitrix-TenantContext.ps1 not found.'
}
. $tenantContextPath
$null = Initialize-SmartCitrixTenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ScriptLocalConfig = Get-SmartCitrixScriptLocalConfig -ScriptPath $PSCommandPath
if ([string]::IsNullOrWhiteSpace($AdminAddress)) { $AdminAddress = Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'AdminAddress' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { $LatestOutputRoot = Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue '' }
if ($MaxRecordCount -le 0) { $MaxRecordCount = [int](Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'MaxRecordCount' -DefaultValue 250000) }

$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartCitrixOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Citrix\OnPrem\CVAD\Delivery'
$runOutputRoot = Join-Path -Path $resolvedOutputRoots.OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartCitrix-OnPrem-CVADDelivery-Inventory_{0}.log" -f $runId)

Import-SmartCitrixCoreModule -StartPath $PSScriptRoot
Set-SmartCitrixCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $resolvedOutputRoots.LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $resolvedOutputRoots.LatestOutputRoot -ItemType Directory -Force | Out-Null
    Write-SmartCitrixLog -Message "Starting SmartCitrix on-prem CVAD delivery inventory. RunId=$runId"

    Import-SmartCitrixPowerShellComponent -Name @('Citrix.Broker.Admin.V2') -Required

    $exports = @(
        @{ Csv = 'Citrix_OnPrem_CVAD_Catalogs'; Source = 'BrokerCatalog'; Command = 'Get-BrokerCatalog' },
        @{ Csv = 'Citrix_OnPrem_CVAD_DeliveryGroups'; Source = 'BrokerDesktopGroup'; Command = 'Get-BrokerDesktopGroup' },
        @{ Csv = 'Citrix_OnPrem_CVAD_Applications'; Source = 'BrokerApplication'; Command = 'Get-BrokerApplication' },
        @{ Csv = 'Citrix_OnPrem_CVAD_ApplicationGroups'; Source = 'BrokerApplicationGroup'; Command = 'Get-BrokerApplicationGroup' },
        @{ Csv = 'Citrix_OnPrem_CVAD_Desktops'; Source = 'BrokerDesktop'; Command = 'Get-BrokerDesktop' },
        @{ Csv = 'Citrix_OnPrem_CVAD_AccessPolicyRules'; Source = 'BrokerAccessPolicyRule'; Command = 'Get-BrokerAccessPolicyRule' },
        @{ Csv = 'Citrix_OnPrem_CVAD_AppEntitlementPolicyRules'; Source = 'BrokerAppEntitlementPolicyRule'; Command = 'Get-BrokerAppEntitlementPolicyRule' },
        @{ Csv = 'Citrix_OnPrem_CVAD_EntitlementPolicyRules'; Source = 'BrokerEntitlementPolicyRule'; Command = 'Get-BrokerEntitlementPolicyRule' },
        @{ Csv = 'Citrix_OnPrem_CVAD_AssignmentPolicyRules'; Source = 'BrokerAssignmentPolicyRule'; Command = 'Get-BrokerAssignmentPolicyRule' },
        @{ Csv = 'Citrix_OnPrem_CVAD_Tags'; Source = 'BrokerTag'; Command = 'Get-BrokerTag' },
        @{ Csv = 'Citrix_OnPrem_CVAD_RebootSchedulesV2'; Source = 'BrokerRebootScheduleV2'; Command = 'Get-BrokerRebootScheduleV2' },
        @{ Csv = 'Citrix_OnPrem_CVAD_RebootSchedules'; Source = 'BrokerRebootSchedule'; Command = 'Get-BrokerRebootSchedule' }
    )

    $summaryRows = @()
    foreach ($export in $exports) {
        $rows = @(Invoke-SmartCitrixSafeInventoryBlock -Name $export.Source -ScriptBlock {
            Invoke-SmartCitrixSdkCommand -CommandName $export.Command -AdminAddress $AdminAddress -MaxRecordCount $MaxRecordCount
        })
        Export-SmartCitrixCsv -Name $export.Csv -Rows (ConvertTo-SmartCitrixFlatRows -Rows $rows -SourceObject $export.Source)
        $summaryRows += New-SmartCitrixSummaryRow -Name $export.Source -Count $rows.Count
    }

    Export-SmartCitrixCsv -Name 'Citrix_OnPrem_CVAD_DeliverySummary' -Rows $summaryRows
    Write-SmartCitrixLog -Level SUCCESS -Message 'SmartCitrix on-prem CVAD delivery inventory completed.'
}
catch {
    Write-SmartCitrixLog -Level ERROR -Message $_.Exception.Message
    throw
}
