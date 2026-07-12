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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDNBDUZpyY1WIGx
# Vo+1SbXhmRMR5auzA31q2fJ7xNRzxqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBfRMxNmVyXTX9Hbd0e
# QZ8edHbT5DWB4+uGuAxYGNk0FzANBgkqhkiG9w0BAQEFAASCAYCjKiu2FkIFQTkm
# Tm8cSKwnwwpqEm7EOz4Bo4AndiBzEK6aYKsYo5ZeOVP01BZSFDcn6gg0um5wXM0R
# D1MWeoB0wLJc75BJtmxEwPLVRWX+7o+3MSr9LLSBVdkSilCa+7txMCObsegp9pfw
# czSF5qanZ9fDXIOGzYSNRos2RD3D9gZN2uUh3fJXV6tih6YjJ8rq7NQX4HnQsBDv
# HsY9CGa4LMrZ+NYi6DpGosM/i90Rr3lgF71bMTcO1OED/MxtvhBdMJY65c52E2O/
# tKQtogQERcaI0mYe5S0zu9DvqjEj0dknDpy/wzVyIJcNwjA6qUo+C9OVz1PTH4zw
# 5sPUROEkZOAZB428KJrHu2e2dKmN0itXoha3CaHKmixGfeWtbnWDBeaZaw+XQqq4
# J2IA/g2e1z8AapKIhB6VI9HOOgvijHa7AqaJmkY1qoLBT8L8LY9Ia39ixHIwanoe
# Q4iQxZq/rGDI3V3P+Bq9pgL+8S89xFevdSlaWgkfYzHk+TICKdg=
# SIG # End signature block
