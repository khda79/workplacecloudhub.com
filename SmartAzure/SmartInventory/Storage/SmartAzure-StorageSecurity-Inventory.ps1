<#
.SYNOPSIS
Exports Azure Storage account security posture inventory.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
storage account public network access, anonymous blob access, shared key access, HTTPS only,
minimum TLS version, firewall settings, and blob service properties.

The script is read-only and uses Azure Resource Manager REST calls through Invoke-AzRestMethod.

.PARAMETER Tenant
Local tenant profile key used to isolate SmartAzure output folders. Defaults to test.

.PARAMETER TenantId
Optional tenant ID or domain used by Connect-AzAccount. If omitted, the SmartAzure tenant
profile AzureTenantId, TenantId, or OrgDomain is used when available.

.PARAMETER SubscriptionId
Optional list of subscription IDs to include. When omitted, all visible subscriptions are included.

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
    [string]$StorageApiVersion = '2023-01-01'
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
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { $LatestOutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue '' }

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\Storage'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-StorageSecurity-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-ResourceGroupFromId {
    [CmdletBinding()]
    param([string]$ResourceId)
    if ($ResourceId -match '/resourceGroups/([^/]+)') { return $matches[1] }
    return ''
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure Storage security inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts

    $context = Connect-SmartAzureCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw "No Azure subscriptions were found for the current identity and filters." }

    $accountRows = @()
    $blobServiceRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null
        $encodedSubscriptionId = [uri]::EscapeDataString($subscription.Id)
        $accounts = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Storage/storageAccounts?api-version={1}" -f $encodedSubscriptionId, $StorageApiVersion) -Operation "Storage accounts for $($subscription.Name)"

        foreach ($account in $accounts) {
            $properties = Get-ObjectPropertyValue -InputObject $account -PropertyName @('properties')
            $network = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('networkAcls')
            $accountRows += [pscustomobject]@{
                RunId                   = $runId
                SubscriptionId          = $subscription.Id
                SubscriptionName        = $subscription.Name
                ResourceGroup           = Get-ResourceGroupFromId -ResourceId $account.id
                StorageAccountName      = $account.name
                Location                = $account.location
                Kind                    = $account.kind
                SkuName                 = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $account -PropertyName @('sku')) -PropertyName @('name')
                PublicNetworkAccess     = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('publicNetworkAccess')
                AllowBlobPublicAccess   = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('allowBlobPublicAccess')
                AllowSharedKeyAccess    = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('allowSharedKeyAccess')
                SupportsHttpsTrafficOnly = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('supportsHttpsTrafficOnly')
                MinimumTlsVersion       = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('minimumTlsVersion')
                DefaultAction           = Get-ObjectPropertyValue -InputObject $network -PropertyName @('defaultAction')
                Bypass                  = Get-ObjectPropertyValue -InputObject $network -PropertyName @('bypass')
                IpRulesCount            = @((Get-ObjectPropertyValue -InputObject $network -PropertyName @('ipRules'))).Count
                VirtualNetworkRulesCount = @((Get-ObjectPropertyValue -InputObject $network -PropertyName @('virtualNetworkRules'))).Count
                PrivateEndpointConnectionCount = @((Get-ObjectPropertyValue -InputObject $properties -PropertyName @('privateEndpointConnections'))).Count
                AccessTier              = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('accessTier')
                CreationTime            = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('creationTime')
                TagsJson                = ConvertTo-CompactJson -Value $account.tags
                ResourceId              = $account.id
            }

            $blobUri = "https://management.azure.com{0}/blobServices/default?api-version={1}" -f $account.id, $StorageApiVersion
            $blobServices = Invoke-SmartAzureArmGetPaged -Uri $blobUri -Operation "Blob service properties for $($account.name)"
            foreach ($blobService in $blobServices) {
                $blobProperties = Get-ObjectPropertyValue -InputObject $blobService -PropertyName @('properties')
                $deleteRetention = Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('deleteRetentionPolicy')
                $containerDeleteRetention = Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('containerDeleteRetentionPolicy')
                $blobServiceRows += [pscustomobject]@{
                    RunId                         = $runId
                    SubscriptionId                = $subscription.Id
                    SubscriptionName              = $subscription.Name
                    ResourceGroup                 = Get-ResourceGroupFromId -ResourceId $account.id
                    StorageAccountName            = $account.name
                    IsVersioningEnabled           = Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('isVersioningEnabled')
                    ChangeFeedEnabled             = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('changeFeed')) -PropertyName @('enabled')
                    BlobDeleteRetentionEnabled    = Get-ObjectPropertyValue -InputObject $deleteRetention -PropertyName @('enabled')
                    BlobDeleteRetentionDays       = Get-ObjectPropertyValue -InputObject $deleteRetention -PropertyName @('days')
                    ContainerDeleteRetentionEnabled = Get-ObjectPropertyValue -InputObject $containerDeleteRetention -PropertyName @('enabled')
                    ContainerDeleteRetentionDays  = Get-ObjectPropertyValue -InputObject $containerDeleteRetention -PropertyName @('days')
                    CorsRulesJson                 = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $blobProperties -PropertyName @('cors')) -PropertyName @('corsRules'))
                    PropertiesJson                = ConvertTo-CompactJson -Value $blobProperties
                    ResourceId                    = $blobService.id
                }
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                         = $runId
            SubscriptionId                = $subscription.Id
            SubscriptionName              = $subscription.Name
            TenantId                      = $subscription.TenantId
            StorageAccountCount           = $accounts.Count
            PublicNetworkEnabledCount     = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.PublicNetworkAccess -eq 'Enabled' }).Count
            BlobPublicAccessAllowedCount  = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.AllowBlobPublicAccess -eq $true }).Count
            SharedKeyAllowedCount         = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.AllowSharedKeyAccess -ne $false }).Count
            NonHttpsOnlyCount             = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.SupportsHttpsTrafficOnly -ne $true }).Count
            FirewallAllowByDefaultCount   = @($accountRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.DefaultAction -eq 'Allow' }).Count
        }
    }

    Export-SmartAzureCsv -Name 'Azure_Storage_AccountsSecurity' -Rows $accountRows
    Export-SmartAzureCsv -Name 'Azure_Storage_BlobServiceSecurity' -Rows $blobServiceRows
    Export-SmartAzureCsv -Name 'Azure_Storage_Summary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed Storage security inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxH6nvkWVB+W20
# CmLUDmXsKkrp/SrqW5S4hRRhgtcrHaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAMDbPqDsILmiwax/hLcpZvsBL2khmwwbgKsyLEJKY17TANBgkqhkiG9w0B
# AQEFAASCAYC47gX5nQKM1fqrme6KPeOHfTXdnWq7iSTAGY9LdiNmYzzaCCNzIfcs
# G7dPkmxwiAtTzF4t0gyrTTOeNbr3pWMfuzGGRXCRPNwtokJi1kqkmWTKqdfn10zX
# IYIv2HMYAVww4Ux9W3NzSotFC3ERoGP4IcPvamICqdJZwPbCU0t4PkZdQi4Q0dHo
# QuLBu1foivnhj/w4K+TcLFm3VuXTRgYvPh5G2P3B7cKoSqizXfI280ttJVk2chyo
# eK9qFNtE4xti0sSHmAiISretkAjqZKqNP6u+6cFvkbMkqX9cUsa/anA4mwDSm+ds
# Jh++QqDb4bxrapniMkgRKYtjNqlZrTXR2NUWHhraNyOE7nQjvjSOSTJJl3u1E4gp
# SWuqKFYp6Ltn3W4YbQHOpd82ywEbpnDMcLjAE1mRF0NNdB2ZPQYfRuZDO44Rubk8
# L3K6TaFAgWWsf2a7Q3L1gLWjqcjbr/ZGNj9Bd+DbJ//44UjwY/0w2BJZKkq7MH2N
# b5NaCb2caJ8=
# SIG # End signature block
