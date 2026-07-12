<#
.SYNOPSIS
Exports candidate FSLogix profile storage used by Azure Virtual Desktop environments.

.DESCRIPTION
Connects with Az PowerShell, identifies resource groups containing Azure Virtual Desktop
resources, then exports storage accounts and Azure Files shares in those resource groups.
Shares and accounts are marked as FSLogix candidates when their names or tags contain
common profile container signals.

The script is read-only. It does not modify storage accounts, shares, or networking.

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
Requires: Az.Accounts
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$TenantId,
    [string[]]$SubscriptionId,
    [string]$OutputRoot,
    [string]$LatestOutputRoot,
    [switch]$Connect,
    [switch]$UseDeviceCode,
    [string]$AvdResourceApiVersion = '2021-04-01',
    [string]$StorageApiVersion = '2023-01-01'
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    exit 1
}

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path -Path $d -ChildPath 'Config\SmartAzureVirtualDesktop-TenantContext.ps1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartAzureVirtualDesktop-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartAvdEffectiveConfig = Initialize-SmartAvdTenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ScriptLocalConfig = Get-SmartAvdScriptLocalConfig -ScriptPath $PSCommandPath
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    $TenantId = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'AzureTenantId' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') { $TenantId = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '' }
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') { $TenantId = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue '' }
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { $LatestOutputRoot = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue '' }

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartAvdOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'AzureVirtualDesktop\FSLogix'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzureVirtualDesktop-FSLogixStorage-Inventory_{0}.log" -f $runId)

Import-SmartAvdCoreModule -StartPath $PSScriptRoot
Set-SmartAvdCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-FSLogixCandidateReason {
    [CmdletBinding()]
    param(
        [string]$Name,
        [AllowNull()]$Tags
    )

    $signals = @()
    if ($Name -match '(?i)(fslogix|profile|profiles|container|containers|avd|wvd)') { $signals += 'NameSignal' }
    if ($null -ne $Tags) {
        $tagJson = ConvertTo-SmartAvdCompactJson -Value $Tags
        if ($tagJson -match '(?i)(fslogix|profile|profiles|container|containers|avd|wvd)') { $signals += 'TagSignal' }
    }
    return ($signals | Select-Object -Unique) -join ';'
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAvdLog -Message "Starting SmartAzureVirtualDesktop FSLogix storage inventory. RunId=$runId"
    Import-SmartAvdRequiredModule -Name Az.Accounts
    $context = Connect-SmartAvdCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw 'No Azure subscriptions were found for the current identity and filters.' }

    $storageAccountRows = @()
    $fileShareRows = @()
    $summaryRows = @()
    $avdResourceTypes = @(
        'Microsoft.DesktopVirtualization/hostPools',
        'Microsoft.DesktopVirtualization/workspaces',
        'Microsoft.DesktopVirtualization/applicationGroups'
    )

    foreach ($subscription in $subscriptions) {
        Write-SmartAvdLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $avdResourceGroups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($resourceType in $avdResourceTypes) {
            $resources = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType $resourceType -ApiVersion $AvdResourceApiVersion)
            foreach ($resource in $resources) {
                $resourceGroupName = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                if (-not [string]::IsNullOrWhiteSpace($resourceGroupName)) { [void]$avdResourceGroups.Add($resourceGroupName) }
            }
        }

        if ($avdResourceGroups.Count -eq 0) {
            $summaryRows += [pscustomobject]@{
                RunId                = $runId
                SubscriptionId       = $subscription.Id
                SubscriptionName     = $subscription.Name
                TenantId             = $subscription.TenantId
                AvdResourceGroupCount = 0
                StorageAccountCount  = 0
                FileShareCount       = 0
                CandidateShareCount  = 0
            }
            continue
        }

        $storageAccounts = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.Storage/storageAccounts' -ApiVersion '2021-04-01')
        $storageAccounts = @($storageAccounts | Where-Object { $avdResourceGroups.Contains((Get-SmartAvdResourceGroupNameFromId -ResourceId $_.id)) })
        $candidateShareCount = 0
        $fileShareCount = 0

        foreach ($account in $storageAccounts) {
            $details = Get-SmartAvdResourceById -ResourceId $account.id -ApiVersion $StorageApiVersion -Operation 'Storage account details'
            if ($null -eq $details) { $details = $account }
            $properties = Get-SmartAvdObjectPropertyValue -InputObject $details -PropertyName @('properties')
            $networkAcls = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('networkAcls')
            $accountCandidateReason = Get-FSLogixCandidateReason -Name (Get-SmartAvdResourceNameFromId -ResourceId $account.id) -Tags $details.tags

            $storageAccountRows += [pscustomobject]@{
                RunId                    = $runId
                SubscriptionId           = $subscription.Id
                SubscriptionName         = $subscription.Name
                ResourceGroupName        = Get-SmartAvdResourceGroupNameFromId -ResourceId $account.id
                StorageAccountName       = Get-SmartAvdResourceNameFromId -ResourceId $account.id
                Location                 = $details.location
                Kind                     = $details.kind
                SkuName                  = Get-SmartAvdNestedPropertyValue -InputObject $details -Path @('sku', 'name')
                SkuTier                  = Get-SmartAvdNestedPropertyValue -InputObject $details -Path @('sku', 'tier')
                PublicNetworkAccess      = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('publicNetworkAccess')
                MinimumTlsVersion        = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('minimumTlsVersion')
                SupportsHttpsTrafficOnly = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('supportsHttpsTrafficOnly'))
                AllowSharedKeyAccess     = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('allowSharedKeyAccess'))
                AllowBlobPublicAccess    = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('allowBlobPublicAccess'))
                DefaultNetworkAction     = Get-SmartAvdObjectPropertyValue -InputObject $networkAcls -PropertyName @('defaultAction')
                Bypass                   = Get-SmartAvdObjectPropertyValue -InputObject $networkAcls -PropertyName @('bypass')
                IsFslogixCandidate       = -not [string]::IsNullOrWhiteSpace($accountCandidateReason)
                CandidateReason          = $accountCandidateReason
                TagsJson                 = ConvertTo-SmartAvdCompactJson -Value $details.tags
                ResourceId               = $account.id
            }

            $shares = @(Get-SmartAvdChildResources -ParentResourceId $account.id -ChildPath 'fileServices/default/shares' -ApiVersion $StorageApiVersion -Operation "Azure Files shares for $($account.name)")
            foreach ($share in $shares) {
                $fileShareCount++
                $shareProperties = Get-SmartAvdObjectPropertyValue -InputObject $share -PropertyName @('properties')
                $shareName = Get-SmartAvdResourceNameFromId -ResourceId $share.id
                $shareCandidateReason = Get-FSLogixCandidateReason -Name $shareName -Tags $details.tags
                if (-not [string]::IsNullOrWhiteSpace($shareCandidateReason)) { $candidateShareCount++ }

                $fileShareRows += [pscustomobject]@{
                    RunId                 = $runId
                    SubscriptionId        = $subscription.Id
                    SubscriptionName      = $subscription.Name
                    ResourceGroupName     = Get-SmartAvdResourceGroupNameFromId -ResourceId $account.id
                    StorageAccountName    = Get-SmartAvdResourceNameFromId -ResourceId $account.id
                    ShareName             = $shareName
                    IsFslogixCandidate    = -not [string]::IsNullOrWhiteSpace($shareCandidateReason)
                    CandidateReason       = $shareCandidateReason
                    EnabledProtocol       = Get-SmartAvdObjectPropertyValue -InputObject $shareProperties -PropertyName @('enabledProtocols')
                    AccessTier            = Get-SmartAvdObjectPropertyValue -InputObject $shareProperties -PropertyName @('accessTier')
                    ShareQuotaGb          = Get-SmartAvdObjectPropertyValue -InputObject $shareProperties -PropertyName @('shareQuota')
                    ProvisionedIops       = Get-SmartAvdObjectPropertyValue -InputObject $shareProperties -PropertyName @('provisionedIops')
                    ProvisionedBandwidthMibps = Get-SmartAvdObjectPropertyValue -InputObject $shareProperties -PropertyName @('provisionedBandwidthMibps')
                    LastModifiedTime      = Get-SmartAvdObjectPropertyValue -InputObject $shareProperties -PropertyName @('lastModifiedTime')
                    Deleted               = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $shareProperties -PropertyName @('deleted'))
                    Version               = Get-SmartAvdObjectPropertyValue -InputObject $shareProperties -PropertyName @('version')
                    ResourceId            = $share.id
                }
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                 = $runId
            SubscriptionId        = $subscription.Id
            SubscriptionName      = $subscription.Name
            TenantId              = $subscription.TenantId
            AvdResourceGroupCount = $avdResourceGroups.Count
            StorageAccountCount   = $storageAccounts.Count
            FileShareCount        = $fileShareCount
            CandidateShareCount   = $candidateShareCount
        }
    }

    Export-SmartAvdCsv -Name 'AVD_FSLogix_StorageAccounts' -Rows $storageAccountRows
    Export-SmartAvdCsv -Name 'AVD_FSLogix_FileShares' -Rows $fileShareRows
    Export-SmartAvdCsv -Name 'AVD_FSLogix_Summary' -Rows $summaryRows

    Write-SmartAvdLog -Level SUCCESS -Message ("Completed FSLogix storage inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    if (Get-Command -Name Write-SmartAvdLog -ErrorAction SilentlyContinue) { Write-SmartAvdLog -Level ERROR -Message $_.Exception.Message }
    throw
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC0Bis++Wm3Vmi8
# Ms+neTQjEaQasNpXvccZd/KJUqQwiaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDC5HrsaCLJ1z3zlUze
# HT+I+tKcRHdWTZT91ohCk4Sa3jANBgkqhkiG9w0BAQEFAASCAYBBmKuOVw/JJp87
# 9ydm8ahxL4KKCQJTeR2cR0pAouTLDrci+4NuUBXa2NOridZrTBstbZkkE8sAbjyZ
# bAVJsKG8CYKfr/UF5GWvKTLi0HIIlg3tCIKSTxY+O+qa1zWXDp+xfiKBgav+uHW0
# d6Lee+6/rJjQjKJ+qnLpu9BaUZwv/5mJ38jwePrko5jQSxko31aO3n0WJRALiicA
# e0IvslyuhEpp6hxvIgN5iKx9pJSNvYzEUv2UL1dF4QG2AOiXGHrTd2ata6tVTwoM
# UvLyqCzqFIVaIAUth/n74/HmDhUi2E8iBTIVCXorronNw47Uc+7w2sZMd6Kgld09
# KYzMlJ0i2c1qTGLMObo7C4GKQTFGOzKCU10IargRTwbOd9/zPmhN3puDHJ55o3Ms
# E7+qaBGjdWNnOPiaBO5JRLNqOOC6AjKJbNf2BaRElOS1F6Uj8LzIhAYc3U+M2BlG
# 7hWEJPRWHdsVfkq14kionLj2Usx/PbxxcXuB0r8w+LyDtth5m5w=
# SIG # End signature block
