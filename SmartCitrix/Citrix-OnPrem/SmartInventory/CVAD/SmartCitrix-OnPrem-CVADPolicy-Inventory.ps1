<#
.SYNOPSIS
Exports Citrix CVAD policy inventory.

.DESCRIPTION
Uses available Citrix Broker policy cmdlets to export policy sets, policies, settings,
filters, and policy test status from an on-premises CVAD site.

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
$resolvedOutputRoots = Resolve-SmartCitrixOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Citrix\OnPrem\CVAD\Policy'
$runOutputRoot = Join-Path -Path $resolvedOutputRoots.OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartCitrix-OnPrem-CVADPolicy-Inventory_{0}.log" -f $runId)

Import-SmartCitrixCoreModule -StartPath $PSScriptRoot
Set-SmartCitrixCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $resolvedOutputRoots.LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $resolvedOutputRoots.LatestOutputRoot -ItemType Directory -Force | Out-Null
    Write-SmartCitrixLog -Message "Starting SmartCitrix on-prem CVAD policy inventory. RunId=$runId"

    Import-SmartCitrixPowerShellComponent -Name @('Citrix.Broker.Admin.V2') -Required

    $exports = @(
        @{ Csv = 'Citrix_OnPrem_CVAD_PolicySets'; Source = 'BrokerGpoPolicySet'; Command = 'Get-BrokerGpoPolicySet' },
        @{ Csv = 'Citrix_OnPrem_CVAD_Policies'; Source = 'BrokerGpoPolicy'; Command = 'Get-BrokerGpoPolicy' },
        @{ Csv = 'Citrix_OnPrem_CVAD_PolicySettings'; Source = 'BrokerGpoSetting'; Command = 'Get-BrokerGpoSetting' },
        @{ Csv = 'Citrix_OnPrem_CVAD_PolicySettingDefinitions'; Source = 'BrokerGpoSettingDefinition'; Command = 'Get-BrokerGpoSettingDefinition' },
        @{ Csv = 'Citrix_OnPrem_CVAD_PolicyFilters'; Source = 'BrokerGpoFilter'; Command = 'Get-BrokerGpoFilter' },
        @{ Csv = 'Citrix_OnPrem_CVAD_PolicyFilterDefinitions'; Source = 'BrokerGpoFilterDefinition'; Command = 'Get-BrokerGpoFilterDefinition' },
        @{ Csv = 'Citrix_OnPrem_CVAD_PolicyTestStatus'; Source = 'BrokerGpoTestPolicyStatus'; Command = 'Get-BrokerGpoTestPolicyStatus' },
        @{ Csv = 'Citrix_OnPrem_CVAD_SettingTestStatus'; Source = 'BrokerGpoTestSettingStatus'; Command = 'Get-BrokerGpoTestSettingStatus' },
        @{ Csv = 'Citrix_OnPrem_CVAD_FilterTestStatus'; Source = 'BrokerGpoTestFilterStatus'; Command = 'Get-BrokerGpoTestFilterStatus' }
    )

    $summaryRows = @()
    foreach ($export in $exports) {
        $rows = @(Invoke-SmartCitrixSafeInventoryBlock -Name $export.Source -ScriptBlock {
            Invoke-SmartCitrixSdkCommand -CommandName $export.Command -AdminAddress $AdminAddress -MaxRecordCount $MaxRecordCount
        })
        Export-SmartCitrixCsv -Name $export.Csv -Rows (ConvertTo-SmartCitrixFlatRows -Rows $rows -SourceObject $export.Source)
        $summaryRows += New-SmartCitrixSummaryRow -Name $export.Source -Count $rows.Count
    }

    Export-SmartCitrixCsv -Name 'Citrix_OnPrem_CVAD_PolicySummary' -Rows $summaryRows
    Write-SmartCitrixLog -Level SUCCESS -Message 'SmartCitrix on-prem CVAD policy inventory completed.'
}
catch {
    Write-SmartCitrixLog -Level ERROR -Message $_.Exception.Message
    throw
}
