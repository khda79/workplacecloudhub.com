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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBQS35jGRLTsbsf
# dogkjnnNz9HMDJGqQvYe3Z7CKVB5qKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBWDjrTncUkAUtIoMZFnKk1WoqW2eWejL5a4jP1xJjoXzANBgkqhkiG9w0B
# AQEFAASCAYCHVYF9GDEoQJSTE8UGLHZNjGC5EujOBL41uR3TpNicl0wLQc0RBrf6
# wA7iLcBQKl9mfIM/Wo4AR82FrUhcNcNPq7nsaSQ+WTcRCwcsBL7jSJxC5iudduM2
# 4bpoOvMI6g4n/GSV7YQe6O5KbuMfgOhOgrRAComuiBvnZtOnCMPg6fb85DKK2dYb
# Iw+zSHX0W2k9pG1DYFMyMBA33vH1LY8WrVWL83jJzGjG47rmr0x3sDk/ZnlSMxPC
# MTd54rfUd5IIJY1wvPxz8325rSd5g208kcYMMFfKXcLZ+8qaugIR57UTIHh8FQBc
# zvY7iKdErl4AG7HwCjZE5S5zwA12ywleujAUTFqwRnpUwTznSpC/kbOWiPNdi4uh
# 2stOoJPvGWvuzXVDuOiuUwllso8pSoCC5sk4Nb3GfnqNIb6FZuF2hs18J3noH+52
# FzN52bdoaZhbuKKuFJcT3JR2CzphIhzy4w24HHrrazCzYpp/MctOkWO0LWaj3tI5
# TjEMX6OlpWI=
# SIG # End signature block
