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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDNBDUZpyY1WIGx
# Vo+1SbXhmRMR5auzA31q2fJ7xNRzxqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBfRMxNmVyXTX9Hbd0eQZ8edHbT5DWB4+uGuAxYGNk0FzANBgkqhkiG9w0B
# AQEFAASCAYB8nor7FHSiK4sc1YukmE3raUSmFDlSY/k3R6BrmjDRnhsLw1NbcenB
# vDHbrVlZD6mQfd/aiO7bUCebZqF02bGjmzCTXL1Bh+2SFpH1VYraG02jI9mf1slJ
# +R8jFB2tMtvQ1WcT3AdCQmylQgNuRgdUQ1vfGdA4tORE6jh1u0NGNoLAYD05anHS
# a6IS+nDPNgmwQVfG9/3uXHqbBiGLL5U+6/gGioI7vLS+gL5j9h/hmMxga2Quf/hQ
# 0EZ8aKjvFSG6NEIFny8WA03Hs+jTQOOv5ncU7Mv2YoUfreA8VYsE9aK583wP9qir
# YL9ohFoGgJHAe6Yd6ZEhpxiY2TQ5MJudtiBDjqoaFsYgidG6hbzeULsItBPAZ4Y7
# y39cK+TEMTchtAD8+3g2At467NQEYVKZr9dTuXFYhGLalcyEr+E5+cHvjt0k9uf2
# 1Je8e7Fi28YVjpr8cZuhwlW9pyIxe29kGZMOcIed4ixEEktYxRSftup7McsKKmhp
# ClNKrkn+1JA=
# SIG # End signature block
