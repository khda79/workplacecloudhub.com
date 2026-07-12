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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBxGsMTwRFDSRZC
# zxX9gFTGsC3ZkdQ1rUGWzls+ByabLKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBfV0+gR4HHErIRxURmIo+qPum1KRcj5xIbplPT+qTbsTANBgkqhkiG9w0B
# AQEFAASCAYCoWC0T2UeQgMdXTdm9+v07lSBSWXcZTkiAplxczlEVMjjexZzxkX17
# wkGZSRPE++OGRXpkxyIsdApw/oJydrG59SnhO/ABvX8Nh0owWZcdZKHPuisw3K56
# zKk7j63lD+uz4/ADlz+PONHzxBORNprOljuOumC/1HahBAG1xQFiVWGLk14nnheL
# 5DiYrBYNJFG2rBJUc/fzHTN7poJuGO8pCLnXd/mqZ2WY9h1Z6lT6xY7fb8e0YgQ4
# nYta8kVqvtqtkLbDRE6j0rPPiJ2mv9x+1KkSFDg3fcktoCubno/EbHx9H3dIBC5T
# jckaJEQ90NHO4FKcmaYKgQOwa1FOMR7ibn74/8zUrrDtgNWPsXDwIDvWkW63zTjH
# 9afZmgYOHCqvoX7pA+H3kds7qtbO+97Cg8WxhJ5zhcbSrn9UvpZgCPW0A8UHH32V
# PkwVAtlkUL/ZJKuX2iq/6Pc+m2jT9Jn5nVQ9MaSkqw6tbmV56E+7BgTYjN53gjgU
# pxDOhJrX+dE=
# SIG # End signature block
