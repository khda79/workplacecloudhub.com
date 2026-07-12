<#
.SYNOPSIS
Exports common Azure cost optimization signals.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
unattached managed disks, old snapshots, unused public IP addresses, stopped or deallocated
virtual machines, and a subscription-level summary.

The script is read-only. It does not delete, stop, resize, or modify Azure resources.

.PARAMETER TenantId
Optional tenant ID or domain used by Connect-AzAccount.

.PARAMETER Tenant
Local tenant profile key used to isolate SmartAzure output folders. Defaults to test.

.PARAMETER SubscriptionId
Optional list of subscription IDs to include. When omitted, all visible subscriptions are included.

.PARAMETER SnapshotOlderThanDays
Snapshots older than this age are exported as review candidates. Default is 30 days.

.PARAMETER OutputRoot
Historical output root. A run-specific folder is created below this path.

.PARAMETER LatestOutputRoot
Latest CSV output root. Stable non-timestamped copies are written here.

.PARAMETER Connect
Forces an Azure sign-in before inventory.

.PARAMETER UseDeviceCode
Uses device code authentication for Connect-AzAccount.

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
Requires: Az.Accounts, Az.Compute, Az.Network
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$TenantId,
    [string[]]$SubscriptionId,
    [int]$SnapshotOlderThanDays = 30,
    [string]$OutputRoot,
    [string]$LatestOutputRoot,
    [switch]$Connect,
    [switch]$UseDeviceCode
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path -Path $d -ChildPath 'Config\SmartAzure-TenantContext.ps1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartAzure-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartAzureEffectiveConfig = Initialize-SmartAzureTenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ScriptLocalConfig = Get-SmartAzureScriptLocalConfig -ScriptPath $PSCommandPath
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    $TenantId = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'AzureTenantId' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') {
        $TenantId = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue ''
    }
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') {
        $TenantId = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue ''
    }
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue ''
}
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) {
    $LatestOutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue ''
}

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\Cost'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-CostOptimization-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-ResourceGroupNameFromId {
    [CmdletBinding()]
    param([string]$ResourceId)

    if ($ResourceId -match '/resourceGroups/([^/]+)') { return $matches[1] }
    return ''
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure cost optimization inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts
    Import-RequiredModule -Name Az.Compute
    Import-RequiredModule -Name Az.Network

    $context = Connect-SmartAzureCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }

    if ($subscriptions.Count -eq 0) {
        throw "No Azure subscriptions were found for the current identity and filters."
    }

    $snapshotCutoff = (Get-Date).AddDays(-1 * [Math]::Abs($SnapshotOlderThanDays))
    $unattachedDiskRows = @()
    $oldSnapshotRows = @()
    $unusedPublicIpRows = @()
    $stoppedVmRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $disks = @(Invoke-SafeInventoryBlock -Name "Disk inventory for $($subscription.Name)" -ScriptBlock { Get-AzDisk })
        $unattachedDisks = @($disks | Where-Object { [string]::IsNullOrWhiteSpace($_.ManagedBy) })
        foreach ($disk in $unattachedDisks) {
            $unattachedDiskRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                ResourceGroupName = $disk.ResourceGroupName
                Name             = $disk.Name
                Location         = $disk.Location
                DiskSizeGB       = $disk.DiskSizeGB
                SkuName          = if ($disk.Sku) { $disk.Sku.Name } else { '' }
                OsType           = $disk.OsType
                DiskState        = Get-ObjectPropertyValue -InputObject $disk -PropertyName @('DiskState')
                TimeCreated      = $disk.TimeCreated
                ManagedBy        = $disk.ManagedBy
                TagsJson         = ConvertTo-CompactJson -Value $disk.Tags
                ResourceId       = $disk.Id
            }
        }

        $snapshots = @(Invoke-SafeInventoryBlock -Name "Snapshot inventory for $($subscription.Name)" -ScriptBlock { Get-AzSnapshot })
        $oldSnapshots = @($snapshots | Where-Object { $_.TimeCreated -and $_.TimeCreated -lt $snapshotCutoff })
        foreach ($snapshot in $oldSnapshots) {
            $oldSnapshotRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                ResourceGroupName = $snapshot.ResourceGroupName
                Name             = $snapshot.Name
                Location         = $snapshot.Location
                DiskSizeGB       = $snapshot.DiskSizeGB
                SkuName          = if ($snapshot.Sku) { $snapshot.Sku.Name } else { '' }
                TimeCreated      = $snapshot.TimeCreated
                AgeDays          = if ($snapshot.TimeCreated) { [int]((Get-Date) - $snapshot.TimeCreated).TotalDays } else { $null }
                OsType           = $snapshot.OsType
                TagsJson         = ConvertTo-CompactJson -Value $snapshot.Tags
                ResourceId       = $snapshot.Id
            }
        }

        $publicIps = @(Invoke-SafeInventoryBlock -Name "Public IP inventory for $($subscription.Name)" -ScriptBlock { Get-AzPublicIpAddress })
        $unusedPublicIps = @($publicIps | Where-Object {
            $ipConfig = Get-ObjectPropertyValue -InputObject $_ -PropertyName @('IpConfiguration')
            $natGateway = Get-ObjectPropertyValue -InputObject $_ -PropertyName @('NatGateway')
            $null -eq $ipConfig -and $null -eq $natGateway
        })
        foreach ($publicIp in $unusedPublicIps) {
            $unusedPublicIpRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                ResourceGroupName = $publicIp.ResourceGroupName
                Name             = $publicIp.Name
                Location         = $publicIp.Location
                IpAddress        = $publicIp.IpAddress
                PublicIpAllocationMethod = $publicIp.PublicIpAllocationMethod
                PublicIpAddressVersion = $publicIp.PublicIpAddressVersion
                SkuName          = if ($publicIp.Sku) { $publicIp.Sku.Name } else { '' }
                IdleTimeoutInMinutes = $publicIp.IdleTimeoutInMinutes
                DnsName          = if ($publicIp.DnsSettings) { $publicIp.DnsSettings.DomainNameLabel } else { '' }
                Fqdn             = if ($publicIp.DnsSettings) { $publicIp.DnsSettings.Fqdn } else { '' }
                TagsJson         = ConvertTo-CompactJson -Value $publicIp.Tag
                ResourceId       = $publicIp.Id
            }
        }

        $vms = @(Invoke-SafeInventoryBlock -Name "VM power state inventory for $($subscription.Name)" -ScriptBlock { Get-AzVM -Status })
        $stoppedVms = @($vms | Where-Object {
            $statusCodes = @($_.Statuses | ForEach-Object { $_.Code })
            $statusCodes -contains 'PowerState/deallocated' -or $statusCodes -contains 'PowerState/stopped'
        })
        foreach ($vm in $stoppedVms) {
            $powerState = @($vm.Statuses | Where-Object { $_.Code -like 'PowerState/*' } | Select-Object -First 1)
            $stoppedVmRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                ResourceGroupName = $vm.ResourceGroupName
                Name             = $vm.Name
                Location         = $vm.Location
                VmSize           = $vm.HardwareProfile.VmSize
                PowerState       = if ($powerState) { $powerState.DisplayStatus } else { '' }
                OsType           = if ($vm.StorageProfile.OsDisk) { $vm.StorageProfile.OsDisk.OsType } else { '' }
                OsDiskName       = if ($vm.StorageProfile.OsDisk) { $vm.StorageProfile.OsDisk.Name } else { '' }
                LicenseType      = $vm.LicenseType
                TagsJson         = ConvertTo-CompactJson -Value $vm.Tags
                ResourceId       = $vm.Id
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                 = $runId
            SubscriptionId        = $subscription.Id
            SubscriptionName      = $subscription.Name
            TenantId              = $subscription.TenantId
            UnattachedDiskCount   = $unattachedDisks.Count
            OldSnapshotCount      = $oldSnapshots.Count
            UnusedPublicIpCount   = $unusedPublicIps.Count
            StoppedVmCount        = $stoppedVms.Count
        }
    }

    Export-SmartAzureCsv -Name 'Azure_Cost_UnattachedDisks' -Rows $unattachedDiskRows
    Export-SmartAzureCsv -Name 'Azure_Cost_OldSnapshots' -Rows $oldSnapshotRows
    Export-SmartAzureCsv -Name 'Azure_Cost_UnusedPublicIPs' -Rows $unusedPublicIpRows
    Export-SmartAzureCsv -Name 'Azure_Cost_StoppedVMs' -Rows $stoppedVmRows
    Export-SmartAzureCsv -Name 'Azure_Cost_Summary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed cost optimization inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDKOXHBJbV6153e
# ACO0VQKlNJcl0uw9iSeBg/k20ZTCoqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBlsaup8nzV3djIcxZr
# swA4b6H+oebWluiywpqoFB+adTANBgkqhkiG9w0BAQEFAASCAYA58I9KwEZ+SFDB
# /jMXIYH4d27NDhhlVAADpqUh7RiO+cjEQBDCXAzM6NnNxT4O/8671wFJ7721hWqv
# RoJvR4xRgtLoSsh7smdbx5U8WqN8UTv/XZWLC345JYiuMABFb3aaYCm5SuN+Nz/c
# eNUEOMNkOqRcQeCr5Y5JRZMrh/n3AkM3FeWbg3On8CKUbXeMST8s/5ScLfUFC6eu
# CU49v7kav3jCIR+WkgoChiDjNZifbvwJGqxhpL+evmaOwHgRPQTWetMYtE4Gc85O
# cu+JPWluxi3lBdX5HV7G8Ky74G9RvpGWD085ODIgjt2wUKVXuPR+KkpL9vslQkGg
# j4jXIxJEkhKbmTwMdTHvgvo5tPsxoaRFs0IAPOwtsBFJZDpIWr1r/5vMeLsnCxB9
# 6Bbehy9Wbw1I7wyB2PwuY9zLnfZHuwHRvH4dXEGh2xq8/+p3RXdje2c/xKRqx1jY
# Jr6F6QoTqNnMncuJUbTbjvvdBpPY+edDU8yCLOFtkVXgy8mpOtE=
# SIG # End signature block
