<#
.SYNOPSIS
Exports Citrix StoreFront on-premises inventory.

.DESCRIPTION
Uses installed StoreFront PowerShell modules to export deployment, store, authentication,
receiver, farm, gateway, and beacon data where supported by the local StoreFront version.

The script is read-only and does not modify StoreFront configuration.

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
Requires: Citrix StoreFront PowerShell modules
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
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
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { $LatestOutputRoot = Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue '' }
if ($MaxRecordCount -le 0) { $MaxRecordCount = [int](Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'MaxRecordCount' -DefaultValue 250000) }

$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartCitrixOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Citrix\OnPrem\StoreFront'
$runOutputRoot = Join-Path -Path $resolvedOutputRoots.OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartCitrix-OnPrem-StoreFront-Inventory_{0}.log" -f $runId)

Import-SmartCitrixCoreModule -StartPath $PSScriptRoot
Set-SmartCitrixCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $resolvedOutputRoots.LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $resolvedOutputRoots.LatestOutputRoot -ItemType Directory -Force | Out-Null
    Write-SmartCitrixLog -Message "Starting SmartCitrix on-prem StoreFront inventory. RunId=$runId"

    Import-SmartCitrixPowerShellComponent -Name @('Citrix.StoreFront', 'Citrix.StoreFront.Stores', 'Citrix.StoreFront.Authentication', 'Citrix.StoreFront.WebReceiver', 'Citrix.StoreFront.Roaming') -Required

    $exports = @(
        @{ Csv = 'Citrix_OnPrem_StoreFront_Deployment'; Source = 'STFDeployment'; Command = 'Get-STFDeployment' },
        @{ Csv = 'Citrix_OnPrem_StoreFront_Stores'; Source = 'STFStoreService'; Command = 'Get-STFStoreService' },
        @{ Csv = 'Citrix_OnPrem_StoreFront_AuthenticationServices'; Source = 'STFAuthenticationService'; Command = 'Get-STFAuthenticationService' },
        @{ Csv = 'Citrix_OnPrem_StoreFront_WebReceiverServices'; Source = 'STFWebReceiverService'; Command = 'Get-STFWebReceiverService' },
        @{ Csv = 'Citrix_OnPrem_StoreFront_Farms'; Source = 'STFStoreFarm'; Command = 'Get-STFStoreFarm' },
        @{ Csv = 'Citrix_OnPrem_StoreFront_Gateways'; Source = 'STFRoamingGateway'; Command = 'Get-STFRoamingGateway' },
        @{ Csv = 'Citrix_OnPrem_StoreFront_Beacons'; Source = 'STFBeacon'; Command = 'Get-STFBeacon' }
    )

    $summaryRows = @()
    foreach ($export in $exports) {
        $rows = @(Invoke-SmartCitrixSafeInventoryBlock -Name $export.Source -ScriptBlock {
            Invoke-SmartCitrixSdkCommand -CommandName $export.Command -MaxRecordCount $MaxRecordCount
        })
        Export-SmartCitrixCsv -Name $export.Csv -Rows (ConvertTo-SmartCitrixFlatRows -Rows $rows -SourceObject $export.Source)
        $summaryRows += New-SmartCitrixSummaryRow -Name $export.Source -Count $rows.Count
    }

    Export-SmartCitrixCsv -Name 'Citrix_OnPrem_StoreFront_Summary' -Rows $summaryRows
    Write-SmartCitrixLog -Level SUCCESS -Message 'SmartCitrix on-prem StoreFront inventory completed.'
}
catch {
    Write-SmartCitrixLog -Level ERROR -Message $_.Exception.Message
    throw
}
