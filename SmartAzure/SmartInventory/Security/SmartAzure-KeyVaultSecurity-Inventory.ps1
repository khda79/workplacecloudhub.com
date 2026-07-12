<#
.SYNOPSIS
Exports Azure Key Vault security posture inventory.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
Key Vault public access, purge protection, soft delete, RBAC/access policy model, secrets, and
certificates nearing expiration.

The script is read-only and uses Azure Resource Manager REST calls through Invoke-AzRestMethod.

.PARAMETER Tenant
Local tenant profile key used to isolate SmartAzure output folders. Defaults to test.

.PARAMETER TenantId
Optional tenant ID or domain used by Connect-AzAccount. If omitted, the SmartAzure tenant
profile AzureTenantId, TenantId, or OrgDomain is used when available.

.PARAMETER SubscriptionId
Optional list of subscription IDs to include. When omitted, all visible subscriptions are included.

.PARAMETER ExpiringWithinDays
Secrets and certificates expiring within this number of days are highlighted. Default is 90.

.PARAMETER SkipSecretAndCertificateDetails
Skips management-plane secret and certificate enumeration.

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
    [int]$ExpiringWithinDays = 90,
    [switch]$SkipSecretAndCertificateDetails,
    [string]$OutputRoot,
    [string]$LatestOutputRoot,
    [switch]$Connect,
    [switch]$UseDeviceCode,
    [string]$KeyVaultApiVersion = '2023-07-01'
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
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\KeyVault'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-KeyVaultSecurity-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function ConvertFrom-UnixTime {
    [CmdletBinding()]
    param([AllowNull()]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    try { return [DateTimeOffset]::FromUnixTimeSeconds([int64]$Value).UtcDateTime } catch { return $null }
}

function Get-ResourceGroupFromId {
    [CmdletBinding()]
    param([string]$ResourceId)
    if ($ResourceId -match '/resourceGroups/([^/]+)') { return $matches[1] }
    return ''
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure Key Vault security inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts

    $context = Connect-SmartAzureCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw "No Azure subscriptions were found for the current identity and filters." }

    $vaultRows = @()
    $accessPolicyRows = @()
    $secretRows = @()
    $certificateRows = @()
    $summaryRows = @()
    $expirationLimit = (Get-Date).ToUniversalTime().AddDays($ExpiringWithinDays)

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null
        $encodedSubscriptionId = [uri]::EscapeDataString($subscription.Id)
        $vaults = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.KeyVault/vaults?api-version={1}" -f $encodedSubscriptionId, $KeyVaultApiVersion) -Operation "Key Vaults for $($subscription.Name)"

        foreach ($vault in $vaults) {
            $properties = Get-ObjectPropertyValue -InputObject $vault -PropertyName @('properties')
            $network = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('networkAcls')
            $accessPolicies = @((Get-ObjectPropertyValue -InputObject $properties -PropertyName @('accessPolicies')))
            $vaultRg = Get-ResourceGroupFromId -ResourceId $vault.id
            $enableRbac = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('enableRbacAuthorization')
            $vaultRows += [pscustomobject]@{
                RunId                 = $runId
                SubscriptionId        = $subscription.Id
                SubscriptionName      = $subscription.Name
                ResourceGroup         = $vaultRg
                VaultName             = $vault.name
                Location              = $vault.location
                TenantId              = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('tenantId')
                PublicNetworkAccess   = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('publicNetworkAccess')
                NetworkDefaultAction  = Get-ObjectPropertyValue -InputObject $network -PropertyName @('defaultAction')
                NetworkBypass         = Get-ObjectPropertyValue -InputObject $network -PropertyName @('bypass')
                IpRulesCount          = @((Get-ObjectPropertyValue -InputObject $network -PropertyName @('ipRules'))).Count
                VirtualNetworkRulesCount = @((Get-ObjectPropertyValue -InputObject $network -PropertyName @('virtualNetworkRules'))).Count
                SoftDeleteEnabled     = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('enableSoftDelete')
                SoftDeleteRetentionDays = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('softDeleteRetentionInDays')
                PurgeProtectionEnabled = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('enablePurgeProtection')
                RbacAuthorizationEnabled = $enableRbac
                AccessPolicyCount     = $accessPolicies.Count
                SkuName               = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('sku')) -PropertyName @('name')
                TagsJson              = ConvertTo-CompactJson -Value $vault.tags
                ResourceId            = $vault.id
            }

            foreach ($policy in $accessPolicies) {
                $permissions = Get-ObjectPropertyValue -InputObject $policy -PropertyName @('permissions')
                $accessPolicyRows += [pscustomobject]@{
                    RunId            = $runId
                    SubscriptionId   = $subscription.Id
                    SubscriptionName = $subscription.Name
                    ResourceGroup    = $vaultRg
                    VaultName        = $vault.name
                    TenantId         = Get-ObjectPropertyValue -InputObject $policy -PropertyName @('tenantId')
                    ObjectId         = Get-ObjectPropertyValue -InputObject $policy -PropertyName @('objectId')
                    ApplicationId    = Get-ObjectPropertyValue -InputObject $policy -PropertyName @('applicationId')
                    KeyPermissionsJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $permissions -PropertyName @('keys'))
                    SecretPermissionsJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $permissions -PropertyName @('secrets'))
                    CertificatePermissionsJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $permissions -PropertyName @('certificates'))
                    StoragePermissionsJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $permissions -PropertyName @('storage'))
                }
            }

            if (-not $SkipSecretAndCertificateDetails) {
                $secretUri = "https://management.azure.com{0}/secrets?api-version={1}" -f $vault.id, $KeyVaultApiVersion
                $secrets = Invoke-SmartAzureArmGetPaged -Uri $secretUri -Operation "Key Vault secrets for $($vault.name)"
                foreach ($secret in $secrets) {
                    $secretProperties = Get-ObjectPropertyValue -InputObject $secret -PropertyName @('properties')
                    $attributes = Get-ObjectPropertyValue -InputObject $secretProperties -PropertyName @('attributes')
                    $expiresOn = ConvertFrom-UnixTime -Value (Get-ObjectPropertyValue -InputObject $attributes -PropertyName @('exp'))
                    $secretRows += [pscustomobject]@{
                        RunId            = $runId
                        SubscriptionId   = $subscription.Id
                        SubscriptionName = $subscription.Name
                        ResourceGroup    = $vaultRg
                        VaultName        = $vault.name
                        SecretName       = $secret.name
                        Enabled          = Get-ObjectPropertyValue -InputObject $attributes -PropertyName @('enabled')
                        CreatedOn        = ConvertFrom-UnixTime -Value (Get-ObjectPropertyValue -InputObject $attributes -PropertyName @('created'))
                        UpdatedOn        = ConvertFrom-UnixTime -Value (Get-ObjectPropertyValue -InputObject $attributes -PropertyName @('updated'))
                        ExpiresOn        = $expiresOn
                        IsExpired        = if ($expiresOn) { $expiresOn -lt (Get-Date).ToUniversalTime() } else { $false }
                        ExpiresWithinDays = if ($expiresOn) { $expiresOn -le $expirationLimit } else { $false }
                        ContentType      = Get-ObjectPropertyValue -InputObject $secretProperties -PropertyName @('contentType')
                        ResourceId       = $secret.id
                    }
                }

                $certificateUri = "https://management.azure.com{0}/certificates?api-version={1}" -f $vault.id, $KeyVaultApiVersion
                $certificates = Invoke-SmartAzureArmGetPaged -Uri $certificateUri -Operation "Key Vault certificates for $($vault.name)"
                foreach ($certificate in $certificates) {
                    $certificateProperties = Get-ObjectPropertyValue -InputObject $certificate -PropertyName @('properties')
                    $attributes = Get-ObjectPropertyValue -InputObject $certificateProperties -PropertyName @('attributes')
                    $policy = Get-ObjectPropertyValue -InputObject $certificateProperties -PropertyName @('policy')
                    $expiresOn = ConvertFrom-UnixTime -Value (Get-ObjectPropertyValue -InputObject $attributes -PropertyName @('exp'))
                    $certificateRows += [pscustomobject]@{
                        RunId            = $runId
                        SubscriptionId   = $subscription.Id
                        SubscriptionName = $subscription.Name
                        ResourceGroup    = $vaultRg
                        VaultName        = $vault.name
                        CertificateName  = $certificate.name
                        Enabled          = Get-ObjectPropertyValue -InputObject $attributes -PropertyName @('enabled')
                        CreatedOn        = ConvertFrom-UnixTime -Value (Get-ObjectPropertyValue -InputObject $attributes -PropertyName @('created'))
                        UpdatedOn        = ConvertFrom-UnixTime -Value (Get-ObjectPropertyValue -InputObject $attributes -PropertyName @('updated'))
                        ExpiresOn        = $expiresOn
                        IsExpired        = if ($expiresOn) { $expiresOn -lt (Get-Date).ToUniversalTime() } else { $false }
                        ExpiresWithinDays = if ($expiresOn) { $expiresOn -le $expirationLimit } else { $false }
                        IssuerName       = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $policy -PropertyName @('issuerParameters')) -PropertyName @('name')
                        Subject          = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $policy -PropertyName @('x509CertificateProperties')) -PropertyName @('subject')
                        ResourceId       = $certificate.id
                    }
                }
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                     = $runId
            SubscriptionId            = $subscription.Id
            SubscriptionName          = $subscription.Name
            TenantId                  = $subscription.TenantId
            KeyVaultCount             = $vaults.Count
            PublicNetworkEnabledCount = @($vaultRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.PublicNetworkAccess -eq 'Enabled' }).Count
            PurgeProtectionDisabledCount = @($vaultRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.PurgeProtectionEnabled -ne $true }).Count
            SoftDeleteDisabledCount   = @($vaultRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.SoftDeleteEnabled -eq $false }).Count
            AccessPolicyModelCount    = @($vaultRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.RbacAuthorizationEnabled -ne $true }).Count
            ExpiringSecretCount       = @($secretRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.ExpiresWithinDays -eq $true }).Count
            ExpiringCertificateCount  = @($certificateRows | Where-Object { $_.SubscriptionId -eq $subscription.Id -and $_.ExpiresWithinDays -eq $true }).Count
        }
    }

    Export-SmartAzureCsv -Name 'Azure_KeyVault_VaultSecurity' -Rows $vaultRows
    Export-SmartAzureCsv -Name 'Azure_KeyVault_AccessPolicies' -Rows $accessPolicyRows
    Export-SmartAzureCsv -Name 'Azure_KeyVault_Secrets' -Rows $secretRows
    Export-SmartAzureCsv -Name 'Azure_KeyVault_Certificates' -Rows $certificateRows
    Export-SmartAzureCsv -Name 'Azure_KeyVault_Summary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed Key Vault security inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDei8tFsQQlYjDV
# dkLgR4d1gWXrg+9F/ZZKHi9RqV6lGaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCTMiw3x6pUx8THs0TqMG7XgLjDWWfYbXv4FPw0qev2RjANBgkqhkiG9w0B
# AQEFAASCAYAD/3892kEo+KEglNjRuByC7pMWTcKlDN4zF4qrcbasyQuz++o3m2dc
# 4H0jQnK1bLLmVBB7/tg8x9oXq/VfH2AMvAROx5ukqJ87aRO9bjaxToOdUupSNjwr
# E/zNsezM9CxUfoUxDSqj8doyng3jz/qGhwdB62KE5PTuqZqv+M+NNENPnM+2BcSB
# xU0Gd5+GDQ6D6K6yhhcjU8oFUn4XeyTYK96a3rdyoWEpUfwuRjoRBVe86Y6det3p
# jJwyuEaznJeTZu9moctHuwI/I5NRzwJyH5APsnpKV1rMHfVeLKtlexUUPYDAIlV1
# E8z8k0tk/sCIRoG1eOaRMo4AFJQ1ziZjcev78+DhRPEn3M71KInAgqCQ29iur9Gi
# BtlXM+Ih1Sqv423mVwguVAjOKz9U0DnJWb/A5e1R4OyaPz4whtS11xw0YW2ZayTD
# i5zNY335QznNMmRPGhFFgxiUBTPEZT/lDqXJw8kOqleBCgSvzw2RGoq+cHcj6tBF
# DGOyjmSkZu4=
# SIG # End signature block
