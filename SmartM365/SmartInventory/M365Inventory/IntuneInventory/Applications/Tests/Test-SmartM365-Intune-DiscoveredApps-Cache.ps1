<#
.SYNOPSIS
    Regression test for the Discovered Apps cache validation path.
.DESCRIPTION
    Loads only the cache-related functions from the inventory script, creates an
    incomplete manifest plus a duplicate app/device relation, and verifies that
    the contaminated app is rejected while a valid app is reused.
    No Microsoft Graph connection or production CSV is used.
.VERSION
1.0
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$inventoryScriptPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Intune-DiscoveredApps-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) {
    throw ($parseErrors | Out-String)
}

$requiredFunctions = @(
    'Write-DiscoveredAppsCsvRows',
    'Get-DiscoveredAppsAppDeviceCount',
    'Get-DiscoveredAppsDeviceDetailCacheManifestPath',
    'ConvertTo-DiscoveredAppsCacheStatsMap',
    'Read-DiscoveredAppsDeviceDetailCacheManifest',
    'Use-DiscoveredAppsDeviceDetailCache'
)
$definitions = @($ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $requiredFunctions -contains $node.Name
}, $true) | Sort-Object { $_.Extent.StartOffset })
if ($definitions.Count -ne $requiredFunctions.Count) {
    throw "Expected $($requiredFunctions.Count) cache function definitions, found $($definitions.Count)."
}
foreach ($definition in $definitions) {
    Invoke-Expression $definition.Extent.Text
}

$script:TestLogs = [System.Collections.Generic.List[string]]::new()
function WriteLog {
    param([string]$Message, [string]$Level)
    $script:TestLogs.Add("$Level|$Message") | Out-Null
}
function Add-SmartM365TenantKey {
    process {
        [pscustomobject]@{
            TenantKey = 'tenant-test'
            AppId     = [string]$_.AppId
            DeviceId  = [string]$_.DeviceId
        }
    }
}
function Assert-Equal {
    param($Actual, $Expected, [string]$Label)
    if ($Actual -ne $Expected) {
        throw "$Label expected '$Expected', got '$Actual'."
    }
}

$testRoot = Join-Path $env:TEMP ("SmartM365-DiscoveredApps-CacheTest-" + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
    $cachePath = Join-Path $testRoot 'Intune_DiscoveredApps_AppDeviceRelations.csv'
    $partialPath = Join-Path $testRoot 'partial.csv'
    @(
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-contaminated'; DeviceId = 'device-1' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-contaminated'; DeviceId = 'device-1' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-contaminated'; DeviceId = 'device-2' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-valid'; DeviceId = 'device-3' }
    ) | Export-Csv -LiteralPath $cachePath -NoTypeInformation -Encoding UTF8

    $cacheItem = Get-Item -LiteralPath $cachePath
    $manifestPath = Get-DiscoveredAppsDeviceDetailCacheManifestPath -CsvPath $cachePath
    [ordered]@{
        CacheManifestVersion = 2
        SourceCsvLength      = [int64]$cacheItem.Length
        AppCount             = 1
        TotalRows            = 2
        TotalDeviceRows      = 2
        Stats                = @(
            [pscustomobject]@{
                AppId = 'app-contaminated'; AppName = 'Contaminated'; AppVersion = '1.0'
                Publisher = 'Publisher'; Platform = 'windows'; Rows = 2; DeviceRows = 2
                MetadataOk = $true; EnrichmentOk = $true
            }
        )
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

    $targetApps = @(
        [pscustomobject]@{
            id = 'app-contaminated'; displayName = 'Contaminated'; version = '1.0'
            publisher = 'Publisher'; platform = 'windows'; deviceCount = 2
        }
        [pscustomobject]@{
            id = 'app-valid'; displayName = 'Valid'; version = '1.0'
            publisher = 'Publisher'; platform = 'windows'; deviceCount = 1
        }
    )
    $processed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $cached = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $actualCounts = @{}

    $result = Use-DiscoveredAppsDeviceDetailCache `
        -CachePath $cachePath `
        -TargetApps $targetApps `
        -PartialPath $partialPath `
        -ProcessedAppIds $processed `
        -CachedAppIds $cached `
        -ActualDeviceCountsByAppId $actualCounts `
        -MaxAgeDays 7

    Assert-Equal $result.Used $true 'Cache Used'
    Assert-Equal $result.ManifestUsed $false 'ManifestUsed'
    Assert-Equal $result.Apps 1 'Reusable apps'
    Assert-Equal $result.Rows 1 'Reused rows'
    Assert-Equal $result.RejectedApps 1 'Rejected apps'
    Assert-Equal $result.RejectedRows 3 'Rejected rows'
    Assert-Equal $result.ExcessRows 1 'Excess rows'
    Assert-Equal $processed.Contains('app-valid') $true 'Valid app processed'
    Assert-Equal $processed.Contains('app-contaminated') $false 'Contaminated app excluded'
    Assert-Equal $actualCounts.ContainsKey('app-valid') $true 'Valid actual count present'
    Assert-Equal $actualCounts.ContainsKey('app-contaminated') $false 'Contaminated actual count absent'

    $partialRows = @(Import-Csv -LiteralPath $partialPath)
    Assert-Equal $partialRows.Count 1 'Partial row count'
    Assert-Equal $partialRows[0].AppId 'app-valid' 'Partial AppId'
    Assert-Equal $partialRows[0].DeviceId 'device-3' 'Partial DeviceId'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*incomplete and will not be trusted*' }).Count 1 'Incomplete manifest warning'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*excess rows detected: 1*' }).Count 1 'Excess row warning'

    $script:TestLogs.Clear()
    $trustedCachePath = Join-Path $testRoot 'trusted-cache.csv'
    $trustedPartialPath = Join-Path $testRoot 'trusted-partial.csv'
    @(
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-stale'; DeviceId = 'device-1' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-stale'; DeviceId = 'device-2' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-current'; DeviceId = 'device-3' }
    ) | Export-Csv -LiteralPath $trustedCachePath -NoTypeInformation -Encoding UTF8

    $trustedCacheItem = Get-Item -LiteralPath $trustedCachePath
    $trustedManifestPath = Get-DiscoveredAppsDeviceDetailCacheManifestPath -CsvPath $trustedCachePath
    [ordered]@{
        CacheManifestVersion = 3
        SourceCsvLength      = [int64]$trustedCacheItem.Length
        AppCount             = 2
        TotalRows            = 3
        TotalDeviceRows      = 3
        Stats                = @(
            [pscustomobject]@{
                AppId = 'app-stale'; AppName = 'Stale'; AppVersion = '1.0'
                Publisher = 'Publisher'; Platform = 'windows'; Rows = 2; DeviceRows = 2
                MetadataOk = $true; EnrichmentOk = $true
            }
            [pscustomobject]@{
                AppId = 'app-current'; AppName = 'Current'; AppVersion = '1.0'
                Publisher = 'Publisher'; Platform = 'windows'; Rows = 1; DeviceRows = 1
                MetadataOk = $true; EnrichmentOk = $true
            }
        )
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $trustedManifestPath -Encoding UTF8

    $trustedTargetApps = @(
        [pscustomobject]@{
            id = 'app-stale'; displayName = 'Stale'; version = '1.0'
            publisher = 'Publisher'; platform = 'windows'; deviceCount = 1
        }
        [pscustomobject]@{
            id = 'app-current'; displayName = 'Current'; version = '1.0'
            publisher = 'Publisher'; platform = 'windows'; deviceCount = 1
        }
    )
    $trustedProcessed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $trustedCached = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $trustedActualCounts = @{}

    $trustedResult = Use-DiscoveredAppsDeviceDetailCache `
        -CachePath $trustedCachePath `
        -TargetApps $trustedTargetApps `
        -PartialPath $trustedPartialPath `
        -ProcessedAppIds $trustedProcessed `
        -CachedAppIds $trustedCached `
        -ActualDeviceCountsByAppId $trustedActualCounts `
        -MaxAgeDays 7

    Assert-Equal $trustedResult.Used $true 'Trusted cache Used'
    Assert-Equal $trustedResult.ManifestUsed $true 'Trusted ManifestUsed'
    Assert-Equal $trustedResult.Apps 1 'Trusted reusable apps'
    Assert-Equal $trustedResult.Rows 1 'Trusted reused rows'
    Assert-Equal $trustedResult.RejectedApps 1 'Trusted refresh apps'
    Assert-Equal $trustedResult.ExcessRows 1 'Trusted excess rows'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like 'INFO|App-device relation cache refresh required for 1 app(s)*' }).Count 1 'Trusted refresh info'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like 'WARNING|*cache validation rejected*' }).Count 0 'Trusted refresh warnings'
    Write-Host 'PASS: untrusted cache defects remain warnings; trusted manifest drift is logged as information.' -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBWcVKZcw9yT51R
# 6F3gOPS9fQdA+raIuuakJwkZwTe4XqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIP5sHUP0OGCPDhYRTKfTfM0PrdGjoWxOMjDDwJO50rbhMA0GCSqG
# SIb3DQEBAQUABIIBgFygRE0qsudt3u3C2sLGF/O5exwVpCT6ZCFgkvaDwaa3vCdy
# zYWrmIO6Oyi+YYdhdlwhFpDDAtv4LiyQ6Cu5ZryyLjc7fkKdJgVqv9V+jPzXe4rk
# sFUkhoT4v4BBY3FELNVBTn6u6pOniXJQ2ZG+u/XSdlZm2Nci+qpU3WyCbnkhxkG2
# RyMA3RHc81m3+cWK06i2QLgD9ogVLTgtnEXxxjae6njk3YWYJmjntSCwfkw20ULa
# gelBLKsCQQELRXIjcy+ifYaakJwNcGUg7/yhyJXpmrrRPSiVrcRZFHTkjKLw44Mf
# 6Xz40X1ENDyN7qhK+WKQ312h/myugFZJt9XCu1LQWb9ITW+xOz1nvZjOxRkiS0mL
# Hya0o4FFacJ6neJi+90oZSvwSKinRWGo+NnVl7KvN0GYbtGkQsTVKPS2QlSs5W2W
# 8YCOoJ5ZrzA32MqTd0rjUiJKqL8cVml5ML58wDc/CA+1LfxheQEPCJU49dQfUYCC
# Ma17JN0cPZx8IMdCfaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjAxNzI2
# NTZaMC8GCSqGSIb3DQEJBDEiBCCpqXZ51zyqB1rdLRXcdmcM2ooU3FgJ0IP+vwRL
# q3WHoDANBgkqhkiG9w0BAQEFAASCAgDMW98BuJB1JHqdkPoCOAaWQ3sHQTwOg9GP
# L9lz1PIMaewyULHJSm+cNyB3weYItb8pFo4bEObSjEbC0Zn3loyrqekBrQzc66Nx
# cjoGpL58bZyfCY5afIppQAmxmp6fEJf9Y8mcYVUa30lwc8R8DiaeIgXcIM4IOnUV
# bruPiRmCsB0p4kpvUTHeDVYVp9gUlKiStyo1ah7Vb7Kvqof4p7Lsk3dTI29mw8Ru
# gSJ5Cy20gWV/fxuSyXscvupgixXwrnxDfXEE0IRrBLRSJT6lzd7EqHdrIIvRoQVZ
# PRiA+Qh1cQQokfruD4KG25jmj/vh3AD7zCh/iQf95g7L2cMC0PV4q0SzclIHoHnt
# VI1LZGPNNHDVk9dAivGSQunTAyQeTF5mr8xuNFHAYxMrvNmolm2QlSGivoEmcn5q
# FHHYJwuxFAg2wlVOZh3kCgkSs1fWScozZkJN8Z6+GHPA8kyO9EJLQvLskRoHioME
# tIEy+EvRBJrC6XZcU1j8POBlR3HMZ8nqAiom7lFNaOkkS2OTQ/W7eulgJIayzV+j
# fydf6V26v3bBNks3mCZD0bbdcGCVAA7YQhHYG9G9hii2mfYPwI15uUJnUBRlW8+D
# etAfVUpyr64t6vxokFYtfArRWK8hLpd7MLryWxc80qTUidFOSK3pm+O81NBnIcCg
# wb0tD2dXzg==
# SIG # End signature block
