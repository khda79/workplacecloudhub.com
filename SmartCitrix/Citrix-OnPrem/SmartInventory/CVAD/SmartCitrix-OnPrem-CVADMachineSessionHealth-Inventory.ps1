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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBxGsMTwRFDSRZC
# zxX9gFTGsC3ZkdQ1rUGWzls+ByabLKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBfV0+gR4HHErIRxURm
# Io+qPum1KRcj5xIbplPT+qTbsTANBgkqhkiG9w0BAQEFAASCAYCqCwGuhLserVXD
# JQKrDEpmTnFVoQ8PFqwGDtmd0ER8XerB9yQb+dCWdSfZDvG73RkeL5+t+4XGfdsZ
# jF24/bzHd6VKeIcaG8mAhKYuBXBz5ImRrSgIHIiVmVCbqP9x5n/3tzsTbiESeqfL
# OU3iQ/4/JNozksnkjQ+1t00ZeGce689X4+SjyDYL8baom7NVLhtzQri7yaLxGU5y
# 1hC5iOvyphtsad2Xx/ndUZ3MIVbBJvetjjXLHT98kmWJuaMgFCgYJPVwcm1FqiUe
# bC4wn0zLf1UmMJPdPU2ceYOTun2J+Z8ny9Vz5h+18rFFgvzTS675XU19O7TMPRRO
# cu7U4MIUHNV0H3IJ/Zw8WdDs3dNDuLqn9j7x6a9/Z9UqyucbWT3FtpGpq3U7KQ/2
# aLFWrc9Dn/E3hLhYlUK/BbF7EsZHcaP7ptPyHG7BV6dKnWuQGce8z7TvXCaVJ1r7
# /3hEY74bypuB1/3aPzqRUbDXorC4/0SUAXusL+ArTs7EWk8NG48=
# SIG # End signature block
