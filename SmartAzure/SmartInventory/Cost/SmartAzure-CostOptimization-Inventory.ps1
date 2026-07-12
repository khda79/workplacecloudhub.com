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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDKOXHBJbV6153e
# ACO0VQKlNJcl0uw9iSeBg/k20ZTCoqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBlsaup8nzV3djIcxZrswA4b6H+oebWluiywpqoFB+adTANBgkqhkiG9w0B
# AQEFAASCAYAsS7s2Xz7/8Rql065Q+M35gvAwfd/XSu1dt2tRV9hoCn0+2bAzcB5+
# /1/KueRaYwTuYTKPUPm/vkGeMGGfxtXJ3At33jtcr9EFnEwTY/FhYxWOY8unnU4a
# WxFyAOzdIp8mtfHR2eKk8+KZnqvySYqYQHfE9XLCr1d+kGrznJNcZQ12gmLkj8Dm
# 36v9B44CIbJ3ZNWC/FCArf7Q0+WH0HDJvoscJQD5ohvEyafayoB5kEdVs4ltBw6/
# TZwVvHdCnAzYMSAEv89GAHRMenaURDmPs15Hc5lbau/Ea6Yz9zctAARV9d4Alp8l
# SbVWsu1C9pQFEdGO0rxMFZlVHUxIfQXq6ELpoTBBM4qIzOYrykpWrfhLEP4lAVZL
# XL8kITlVP85vuU4AWhxiy3tnGgWjF9MIqzj0a57+m/Au6UbTQB0XlhSx47QTy/Gs
# R4G9+EBp1IJoWWaE/4p9c0iD22os7qfEOtHwBOY76TUzJFqZomSMjsaJI4JdDHA8
# CmiT7x3jBkI=
# SIG # End signature block
