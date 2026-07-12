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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDCXNrclXGMfS0m
# GvMGvYHYBg3XLDqgrddMKPi5sHSae6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC4q6MyzujEf+L1fjtn
# 6nRPfFCOSW/oyHGeziFwzyFuezANBgkqhkiG9w0BAQEFAASCAYCR5BT9/sr7wVO7
# A0lC46hbgZLZkDTig1G+WqgmWhBr0WpClmNX6xBVSTYv9cBMfTpxN8mokozlxzKN
# XL2Nc/kg58168XSo8XMN7p+hvHUC1ACY+XEB0n0w5K57WGuCjKO/ox+lNqrZcuYt
# zxrLzyWdontMCtSnqzHKEOQgyaZhjCb6NEhfv+VlrQ2MXTbgW4LhnzaQh/AAFC5w
# GGO6VG1XqBu99jq1DiNSK6J0wZy+bjdWMx6btUoPiHVMdtt9DuPu1n0VQ/2UyD3j
# KFIr6Land9vsZQBeXWZSQHacP9A/f8VDC+sO9Lo681zjXdpR7ku3ABOwTLQ/zVlu
# rBzp0YkgFP+fFOd3Kq1lcqFD80Rz1Aikqtn3QbhPVLpgQSYjvR5+XskwBNjyqH8i
# 1GHDuA5pro8aw6YBqFC9ev1Cp1/DUTH1U6wbHOPzozkiT2pHgCd/kCL75vpT+dRS
# 4LBVxnBtK0c4zS7TCZ6et2x/NAXWekHJgVXLNCU7sK8zrqk09xw=
# SIG # End signature block
