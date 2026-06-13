<#
.SYNOPSIS
Exports Citrix CVAD machine and session health inventory.

.DESCRIPTION
Uses the local Citrix PowerShell SDK to export machines, sessions, desktops, and summary
health signals from an on-premises CVAD site.

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
$resolvedOutputRoots = Resolve-SmartCitrixOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Citrix\OnPrem\CVAD\Health'
$runOutputRoot = Join-Path -Path $resolvedOutputRoots.OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartCitrix-OnPrem-CVADMachineSessionHealth-Inventory_{0}.log" -f $runId)

Import-SmartCitrixCoreModule -StartPath $PSScriptRoot
Set-SmartCitrixCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $resolvedOutputRoots.LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartCitrixScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $resolvedOutputRoots.LatestOutputRoot -ItemType Directory -Force | Out-Null
    Write-SmartCitrixLog -Message "Starting SmartCitrix on-prem CVAD machine/session health inventory. RunId=$runId"

    Import-SmartCitrixPowerShellComponent -Name @('Citrix.Broker.Admin.V2') -Required

    $machines = @(Invoke-SmartCitrixSafeInventoryBlock -Name 'BrokerMachine' -ScriptBlock {
        Invoke-SmartCitrixSdkCommand -CommandName 'Get-BrokerMachine' -AdminAddress $AdminAddress -MaxRecordCount $MaxRecordCount
    })
    $sessions = @(Invoke-SmartCitrixSafeInventoryBlock -Name 'BrokerSession' -ScriptBlock {
        Invoke-SmartCitrixSdkCommand -CommandName 'Get-BrokerSession' -AdminAddress $AdminAddress -MaxRecordCount $MaxRecordCount
    })
    $desktops = @(Invoke-SmartCitrixSafeInventoryBlock -Name 'BrokerDesktop' -ScriptBlock {
        Invoke-SmartCitrixSdkCommand -CommandName 'Get-BrokerDesktop' -AdminAddress $AdminAddress -MaxRecordCount $MaxRecordCount
    })

    Export-SmartCitrixCsv -Name 'Citrix_OnPrem_CVAD_Machines' -Rows (ConvertTo-SmartCitrixFlatRows -Rows $machines -SourceObject 'BrokerMachine')
    Export-SmartCitrixCsv -Name 'Citrix_OnPrem_CVAD_Sessions' -Rows (ConvertTo-SmartCitrixFlatRows -Rows $sessions -SourceObject 'BrokerSession')
    Export-SmartCitrixCsv -Name 'Citrix_OnPrem_CVAD_Desktops' -Rows (ConvertTo-SmartCitrixFlatRows -Rows $desktops -SourceObject 'BrokerDesktop')

    $machineSummaryRows = @()
    $machineSummaryRows += New-SmartCitrixSummaryRow -Name 'MachinesTotal' -Count $machines.Count
    $machineSummaryRows += New-SmartCitrixSummaryRow -Name 'MachinesInMaintenanceMode' -Count @($machines | Where-Object { (Get-SmartCitrixObjectPropertyValue -InputObject $_ -PropertyName @('InMaintenanceMode')) -eq $true }).Count
    $machineSummaryRows += New-SmartCitrixSummaryRow -Name 'MachinesUnregistered' -Count @($machines | Where-Object { [string](Get-SmartCitrixObjectPropertyValue -InputObject $_ -PropertyName @('RegistrationState')) -ne 'Registered' }).Count
    $machineSummaryRows += New-SmartCitrixSummaryRow -Name 'MachinesPowerOff' -Count @($machines | Where-Object { [string](Get-SmartCitrixObjectPropertyValue -InputObject $_ -PropertyName @('PowerState')) -in @('Off', 'PoweredOff') }).Count
    $machineSummaryRows += New-SmartCitrixSummaryRow -Name 'SessionsTotal' -Count $sessions.Count
    $machineSummaryRows += New-SmartCitrixSummaryRow -Name 'SessionsDisconnected' -Count @($sessions | Where-Object { [string](Get-SmartCitrixObjectPropertyValue -InputObject $_ -PropertyName @('SessionState')) -eq 'Disconnected' }).Count
    $machineSummaryRows += New-SmartCitrixSummaryRow -Name 'SessionsActive' -Count @($sessions | Where-Object { [string](Get-SmartCitrixObjectPropertyValue -InputObject $_ -PropertyName @('SessionState')) -eq 'Active' }).Count
    Export-SmartCitrixCsv -Name 'Citrix_OnPrem_CVAD_HealthSummary' -Rows $machineSummaryRows

    $registrationRows = @($machines | Group-Object -Property RegistrationState | ForEach-Object {
        [pscustomobject]@{ RunId = $runId; Dimension = 'RegistrationState'; Name = $_.Name; Count = $_.Count }
    })
    $powerRows = @($machines | Group-Object -Property PowerState | ForEach-Object {
        [pscustomobject]@{ RunId = $runId; Dimension = 'PowerState'; Name = $_.Name; Count = $_.Count }
    })
    $sessionStateRows = @($sessions | Group-Object -Property SessionState | ForEach-Object {
        [pscustomobject]@{ RunId = $runId; Dimension = 'SessionState'; Name = $_.Name; Count = $_.Count }
    })
    Export-SmartCitrixCsv -Name 'Citrix_OnPrem_CVAD_HealthSummary_ByState' -Rows @($registrationRows + $powerRows + $sessionStateRows)

    Write-SmartCitrixLog -Level SUCCESS -Message 'SmartCitrix on-prem CVAD machine/session health inventory completed.'
}
catch {
    Write-SmartCitrixLog -Level ERROR -Message $_.Exception.Message
    throw
}
