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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDCXNrclXGMfS0m
# GvMGvYHYBg3XLDqgrddMKPi5sHSae6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCC4q6MyzujEf+L1fjtn6nRPfFCOSW/oyHGeziFwzyFuezANBgkqhkiG9w0B
# AQEFAASCAYAutaJJTpdMV5QhxG/grOMo/KSIHpfGM+1ttjUNGxo2RbIo/48cMKhY
# RRKeg+3HrWmevL+M1c3GNDGzfhexN8S5EAXtd1Z11sfzGSAKIbqbbfOXePoX1AyM
# oT2LqeA+aQH4vuRivJrV01Mt6d2szeNctm6bTuyeolzNZFIAXBFbhWepXN/vGF0/
# 1LT9AGzz+kSdUs79gn0TvbnHrOj/MfRrEQmgvr/dvqVvhU5MoWBm9NgxIO0YK88X
# VMSeZVXPXeGCLdUCBoPduukHbptbXagyxQIUkmqTdySvozNym+NVFWji0PUUEXFt
# HcZhFN8yVxKULxY3oCVPNtFMxqpho1+doT4KFlGSI9ET/uowOrZXEN+nqd82Ie57
# rFvlbMR9Gyk4j8GpoZfoOK/v0dxmhyC7Is+VgXmIQYHMLMJ9+1UYBlhyxJRa6er/
# bbxOOp0ImZdi5+ldv8WuWrNq5XvBrZFv5s708uZzdBJTNwZUUKbJFcl3aT+w7r1Y
# JPGOT66FmbA=
# SIG # End signature block
