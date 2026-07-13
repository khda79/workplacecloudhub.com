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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDei8tFsQQlYjDV
# dkLgR4d1gWXrg+9F/ZZKHi9RqV6lGaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIJMyLDfHqlTHxMezROowbteAuMNZZ9hte/gU/DSp6/ZGMA0GCSqG
# SIb3DQEBAQUABIIBgIZRbZ+tOcLJKt3IST+34FZL9kzqYEvejEK6Abr8ogjsBHeL
# qYr+tnqob5iSgzS0i8Kh3LETtTbnYMQrijwqjGFVYcGSwGoBpOBv0lIau8weuwP/
# h8PrIRdR7lLU9GclcmIe0evGuZX4CWq5AqYfbT34W7y1PZahMogKX9sGdIlH89aX
# BPWgd4rCPJSenvqBbaDr7q3YAMUO3vfKrjvgoxjInMnAuoI+Ue4PL9natbsIcc6y
# 9TDL3Bk/D2UOIGUvC5COiQyzoiiuuJBqhpsTG2Lf0Hl9mPyY9a6QN0iranrU+PGh
# b6AGWmnpmqJo2I6NCRxGwK3tplCslgE1SvWLmZPmoMWoUoy/8k1B7Fe7OrgiJqBC
# 9iEnY/OMkCYGnPik1b/V7PwhsSi154kAciQh4lChjvRD93c1JOdlskhy+BVjRNJj
# r+5JWD0nr5gPqQysVHKmlUCZNeVlqLEpzFIcW32cLqLYbVNw4XSLoKMaZDEZB2kK
# CM8VVE2UaJfrU/yZm6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MTNaMC8GCSqGSIb3DQEJBDEiBCCt5IF6hpiVdvS7FO/6oDIX8MgrMskkHBx4W7VO
# j4YM0zANBgkqhkiG9w0BAQEFAASCAgCswCvpEkjniJ0+cq2eN0DzTGJeSo3Urr5+
# H0fAVgc5R3ags9/lRtSgVdcuLLhrhncE+YLPGYLfIK1BUjSiILRL5h8jJA5i3UpK
# TMTypvwV/WE8gbhUGLYWka9lB/hu+6FV1MJ18wxgmuRUBJ4eppPgXypipKMVEqfA
# iBLqvgU6FRzx8fW7Ws8eXy4b7ZabV2pA/f6u/qU81yJhGuv9NFXEBXsNi80lw/h5
# RgKrIrrHG4cFKBj6NrLj4gjSUmIa2L1P6M9gXG1LfU5s2mgs8hvIJMUeteybFYs0
# FtTPJPPX7HXaT6u6pdcvb6Pp5a2TxPC/9Hu3pjaNQwE3tSl26z9YGa17TumMIo1A
# 4F459vMv/7MBwHQEBIz0wLI9bhP4VC/09YO8fqlY4Td7kz0APfrHGPkYsti2ZAis
# X18/1kqdlzPbFnqsIr7Xq4Kp/M99vHOCDjFuOdr0ttQa+473cDlLUvUIutZieH4D
# jo5nLWYnhc94E2KYHXWyjk9p8g0DVQ4kdYXRYvVIJdKWFRCxhjAs0NESkzppLmTQ
# /S44smSyThmHjbJ7t5BEPTFiRDz5Z5W0cMEC0aeGxiTaHJRA16u+vC3usCDzA3E4
# 0LhAbVUU4K7wNcSpgRevfOLONPsqozevOPwiaTa99Fhg7duyZN+CqS/iX1YPHalh
# qzU94PH3mw==
# SIG # End signature block
