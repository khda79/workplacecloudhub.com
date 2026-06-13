<#
.SYNOPSIS
Exports Citrix CVAD hosting and power management inventory.

.DESCRIPTION
Uses the local Citrix PowerShell SDK to export hypervisor connections, hosting status,
power time schemes, power actions, delayed actions, catalog reboot schedules, and reboot
cycles from an on-premises CVAD site.

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
$resolvedOutputRoots = Resolve-SmartCitrixOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Citrix\OnPrem\CVAD\HostingPower'
$runOutputRoot = Join-Path -Path $resolvedOutputRoots.OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartCitrix-OnPrem-CVADHostingPower-Inventory_{0}.log" -f $runId)

Import-SmartCitrixCoreModule -StartPath $PSScriptRoot
Set-SmartCitrixCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $resolvedOutputRoots.LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $resolvedOutputRoots.LatestOutputRoot -ItemType Directory -Force | Out-Null
    Write-SmartCitrixLog -Message "Starting SmartCitrix on-prem CVAD hosting/power inventory. RunId=$runId"

    Import-SmartCitrixPowerShellComponent -Name @('Citrix.Broker.Admin.V2') -Required
    Import-SmartCitrixPowerShellComponent -Name @('Citrix.Host.Admin.V2', 'Citrix.MachineCreation.Admin.V2')

    $exports = @(
        @{ Csv = 'Citrix_OnPrem_CVAD_HypervisorConnections'; Source = 'BrokerHypervisorConnection'; Command = 'Get-BrokerHypervisorConnection' },
        @{ Csv = 'Citrix_OnPrem_CVAD_HypervisorConnectionStatus'; Source = 'BrokerHypervisorConnectionStatus'; Command = 'Get-BrokerHypervisorConnectionStatus' },
        @{ Csv = 'Citrix_OnPrem_CVAD_HostConnections'; Source = 'HypHypervisorConnection'; Command = 'Get-HypHypervisorConnection' },
        @{ Csv = 'Citrix_OnPrem_CVAD_HostResources'; Source = 'HypHostingUnit'; Command = 'Get-HypHostingUnit' },
        @{ Csv = 'Citrix_OnPrem_CVAD_MachineCreationCatalogs'; Source = 'ProvScheme'; Command = 'Get-ProvScheme' },
        @{ Csv = 'Citrix_OnPrem_CVAD_PowerTimeSchemes'; Source = 'BrokerPowerTimeScheme'; Command = 'Get-BrokerPowerTimeScheme' },
        @{ Csv = 'Citrix_OnPrem_CVAD_HostingPowerActions'; Source = 'BrokerHostingPowerAction'; Command = 'Get-BrokerHostingPowerAction' },
        @{ Csv = 'Citrix_OnPrem_CVAD_DelayedHostingPowerActions'; Source = 'BrokerDelayedHostingPowerAction'; Command = 'Get-BrokerDelayedHostingPowerAction' },
        @{ Csv = 'Citrix_OnPrem_CVAD_CatalogRebootSchedules'; Source = 'BrokerCatalogRebootSchedule'; Command = 'Get-BrokerCatalogRebootSchedule' },
        @{ Csv = 'Citrix_OnPrem_CVAD_RebootCycles'; Source = 'BrokerRebootCycle'; Command = 'Get-BrokerRebootCycle' }
    )

    $summaryRows = @()
    foreach ($export in $exports) {
        $rows = @(Invoke-SmartCitrixSafeInventoryBlock -Name $export.Source -ScriptBlock {
            Invoke-SmartCitrixSdkCommand -CommandName $export.Command -AdminAddress $AdminAddress -MaxRecordCount $MaxRecordCount
        })
        Export-SmartCitrixCsv -Name $export.Csv -Rows (ConvertTo-SmartCitrixFlatRows -Rows $rows -SourceObject $export.Source)
        $summaryRows += New-SmartCitrixSummaryRow -Name $export.Source -Count $rows.Count
    }

    Export-SmartCitrixCsv -Name 'Citrix_OnPrem_CVAD_HostingPowerSummary' -Rows $summaryRows
    Write-SmartCitrixLog -Level SUCCESS -Message 'SmartCitrix on-prem CVAD hosting/power inventory completed.'
}
catch {
    Write-SmartCitrixLog -Level ERROR -Message $_.Exception.Message
    throw
}
