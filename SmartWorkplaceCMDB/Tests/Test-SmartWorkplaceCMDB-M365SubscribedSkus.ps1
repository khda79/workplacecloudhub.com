<#
.SYNOPSIS
Runs offline Microsoft 365 subscribed SKU collector and normalizer tests.

.VERSION
0.1.0
#>
[CmdletBinding()]
param()

$ScriptVersion = '0.1.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:Passed = 0
$script:Failed = 0

function Invoke-SmartWorkplaceCMDBM365SkuTest {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][scriptblock]$Test)
    try {
        & $Test
        $script:Passed++
        Write-Information "PASS $Name" -InformationAction Continue
    }
    catch {
        $script:Failed++
        Write-Information "FAIL $Name - $($_.Exception.Message)" -InformationAction Continue
    }
}

function Assert-SmartWorkplaceCMDBM365SkuTrue {
    [CmdletBinding()]
    param([Parameter(Mandatory)][bool]$Condition, [Parameter(Mandatory)][string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBM365SkuThrow {
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Action, [Parameter(Mandatory)][string]$MessagePattern)
    $caught = $null
    try { & $Action } catch { $caught = $_ }
    if ($null -eq $caught) {
        throw "Expected an exception matching '$MessagePattern'."
    }
    if ($caught.Exception.Message -notmatch $MessagePattern) {
        throw "Unexpected exception: $($caught.Exception.Message)"
    }
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$collector = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Collect.ps1'
$normalizer = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Normalize.ps1'
$coreModule = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContract = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContract = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$fixture = Join-Path $PSScriptRoot 'Fixtures\M365SubscribedSkus.sample.json'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$tempRoot = Join-Path $tempBase ('SmartWorkplaceCMDB-M365Sku-Tests-{0}' -f [guid]::NewGuid().ToString('N'))
Import-Module $coreModule -Force

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $identity = @{
        Tenant = 'audit'; OrganizationKey = 'contoso'; EnvironmentKey = 'prod'
        TenantKey = 'contoso-prod'; TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        NoConfigWrite = $true
    }

    Invoke-SmartWorkplaceCMDBM365SkuTest 'Collector ValidateOnly stays read-only' {
        $root = Join-Path $tempRoot 'Validate'
        & $collector @identity -DataRootPath $root -InputJsonPath $fixture -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBM365SkuTrue (-not (Test-Path $root)) 'ValidateOnly created output.'
    }

    $runtime = Join-Path $tempRoot 'Runtime'
    Invoke-SmartWorkplaceCMDBM365SkuTest 'Collect fixture subscribed SKUs' {
        $script:Collection = & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture
        $rows = @(Import-Csv $script:Collection.RawLatestOutputPath)
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            $script:Collection.SkuCount -eq 2 -and $rows.Count -eq 2
        ) 'Expected two subscribed SKUs.'
    }

    Invoke-SmartWorkplaceCMDBM365SkuTest 'Validate raw contract and MaxItems' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtime 'DATA-LAST') `
            -ContractPath $rawContract)
        $sku = @($results | Where-Object Name -eq 'M365_SubscribedSkus.csv')
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            $sku.Count -eq 1 -and $sku[0].Status -eq 'Valid'
        ) 'Raw subscribed SKU contract is invalid.'
        $bounded = & $collector @identity `
            -DataRootPath (Join-Path $tempRoot 'Bounded') `
            -InputJsonPath $fixture `
            -MaxItems 1
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            @(Import-Csv $bounded.RawLatestOutputPath).Count -eq 1
        ) 'MaxItems was not honored.'
    }

    Invoke-SmartWorkplaceCMDBM365SkuTest 'Normalize licenses and Power BI dimension' {
        $script:Normalization = & $normalizer @identity -DataRootPath $runtime
        $cmdb = @(Import-Csv $script:Normalization.CmdbLicenseOutputPath)
        $dim = @(Import-Csv $script:Normalization.DimLicenseSkuOutputPath)
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            $cmdb.Count -eq 2 -and $dim.Count -eq 2
        ) 'Curated license counts are invalid.'
        $enterprise = @($cmdb | Where-Object SkuPartNumber -eq 'EXAMPLE_ENTERPRISE')
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            $enterprise.Count -eq 1 -and
            $enterprise[0].ConsumedUnits -eq '80' -and
            $enterprise[0].EnabledUnits -eq '100'
        ) 'License unit mapping is invalid.'
    }

    Invoke-SmartWorkplaceCMDBM365SkuTest 'Validate curated contracts and stable keys' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtime 'DATA-LAST') `
            -ContractPath $curatedContract)
        $licenses = @($results | Where-Object Name -in @('CMDB_Licenses.csv', 'DimLicenseSku.csv'))
        $cmdb = @(Import-Csv $script:Normalization.CmdbLicenseOutputPath)
        $dim = @(Import-Csv $script:Normalization.DimLicenseSkuOutputPath)
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            $licenses.Count -eq 2 -and
            @($licenses | Where-Object Status -ne 'Valid').Count -eq 0
        ) 'Curated license contracts are invalid.'
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            $cmdb[0].CmdbLicenseId -eq $dim[0].TenantSkuKey
        ) 'License keys are inconsistent.'
    }

    Invoke-SmartWorkplaceCMDBM365SkuTest 'Aggregate duplicate skuId subscriptions' {
        $rows = @(Import-Csv $script:Collection.RawLatestOutputPath)
        $additional = $rows[0].PSObject.Copy()
        $additional.SubscribedSkuId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb_11111111-aaaa-bbbb-cccc-111111111111'
        $additional.ConsumedUnits = '20'
        $additional.EnabledUnits = '25'
        $aggregatePath = Join-Path $tempRoot 'Aggregate.csv'
        @($rows + $additional) | Export-Csv $aggregatePath -NoTypeInformation -Encoding UTF8
        $result = & $normalizer @identity `
            -DataRootPath (Join-Path $tempRoot 'Aggregate') `
            -RawInputPath $aggregatePath
        $cmdb = @(Import-Csv $result.CmdbLicenseOutputPath)
        $enterprise = @($cmdb | Where-Object SkuPartNumber -eq 'EXAMPLE_ENTERPRISE')
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            $result.RawSkuCount -eq 3 -and $result.LicenseCount -eq 2 -and
            $enterprise[0].ConsumedUnits -eq '100' -and
            $enterprise[0].EnabledUnits -eq '125'
        ) 'Duplicate SKU aggregation is invalid.'
    }

    Invoke-SmartWorkplaceCMDBM365SkuTest 'Reject identity mismatch and duplicate subscription IDs' {
        $rows = @(Import-Csv $script:Collection.RawLatestOutputPath)
        $mismatch = Join-Path $tempRoot 'Mismatch.csv'
        $rows[0].TenantKey = 'other-prod'
        $rows | Export-Csv $mismatch -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBM365SkuThrow {
            & $normalizer @identity `
                -DataRootPath (Join-Path $tempRoot 'Mismatch') `
                -RawInputPath $mismatch | Out-Null
        } 'identity mismatch'
        $rows = @(Import-Csv $script:Collection.RawLatestOutputPath)
        $duplicate = Join-Path $tempRoot 'Duplicate.csv'
        $rows[1].SubscribedSkuId = $rows[0].SubscribedSkuId
        $rows | Export-Csv $duplicate -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBM365SkuThrow {
            & $normalizer @identity `
                -DataRootPath (Join-Path $tempRoot 'Duplicate') `
                -RawInputPath $duplicate | Out-Null
        } 'duplicate SubscribedSkuId'
    }

    Invoke-SmartWorkplaceCMDBM365SkuTest 'Publish valid schema-only outputs for an empty collection' {
        $emptyFixture = Join-Path $tempRoot 'Empty.json'
        '{"value":[]}' | Set-Content -LiteralPath $emptyFixture -Encoding UTF8
        $emptyRoot = Join-Path $tempRoot 'Empty'
        $emptyCollection = & $collector @identity -DataRootPath $emptyRoot -InputJsonPath $emptyFixture
        $emptyNormalization = & $normalizer @identity -DataRootPath $emptyRoot
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $emptyRoot 'DATA-LAST') `
            -ContractPath $curatedContract)
        $licenses = @($results | Where-Object Name -in @('CMDB_Licenses.csv', 'DimLicenseSku.csv'))
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            $emptyCollection.SkuCount -eq 0 -and
            $emptyNormalization.LicenseCount -eq 0 -and
            @($licenses | Where-Object Status -ne 'Valid').Count -eq 0
        ) 'Empty collection outputs are not valid schema-only CSVs.'
    }

    Invoke-SmartWorkplaceCMDBM365SkuTest 'Normalizer ValidateOnly preserves outputs' {
        $beforeCmdb = (Get-FileHash $script:Normalization.CmdbLicenseOutputPath -Algorithm SHA256).Hash
        $beforeDim = (Get-FileHash $script:Normalization.DimLicenseSkuOutputPath -Algorithm SHA256).Hash
        & $normalizer @identity -DataRootPath $runtime -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBM365SkuTrue (
            $beforeCmdb -eq (Get-FileHash $script:Normalization.CmdbLicenseOutputPath -Algorithm SHA256).Hash -and
            $beforeDim -eq (Get-FileHash $script:Normalization.DimLicenseSkuOutputPath -Algorithm SHA256).Hash
        ) 'ValidateOnly modified curated license outputs.'
    }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolved) -like 'SmartWorkplaceCMDB-M365Sku-Tests-*' -and
        (Test-Path $resolved)) {
        Remove-Item $resolved -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB Microsoft 365 subscribed SKUs tests completed. Version={0}; Passed={1}; Failed={2}" -f
    $ScriptVersion, $script:Passed, $script:Failed
) -InformationAction Continue
if ($script:Failed -gt 0) {
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCILhWZQxFVOPt7
# Mg/AHUqHde7jYpNBqKqDYAa3xG72aKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIJdUGGoFIkGa2FIq9Byv6udgvVjIzd/VyTcCV+aNCt9xMA0GCSqG
# SIb3DQEBAQUABIIBgBv/boSF2Bd11azgWllYIrtzkbWf2Wkf3O3h9nQL9cQ6TEnG
# WGLQOWeM4YuP22heRJSseAOg1IEKP8V/c2M6McvJYaxmlYyQlrOIOSRmd73ZnCwP
# +wXqOAjlBSGM2RHlildBCeyOVxp6/0N2UcAhlMNGBQXr8DoJT2E9HrScH5adRIWo
# zEzbIVaLUXl/NtfJfRANxrj36JblmAaeI/Qr8Ase+lyBm39i2108ZkyiOPtYODNY
# r3jLDa5pmjy1XCYyHFaGiQkFMcYA2LPXPGvmoZD+jfIgKt+VEcdJeCeYDj7egaUf
# 75YHrBx4PFUvSQGRiEYZwrpzKlaA9jocI3PONQLeBCocSLkpuq8UBybDrwGUyw7x
# RLAKktWX5Mzg3g+vEEI574FcBGmZ/QqVlW0+G7kJSUtgMh98OMUq0WyB9jq1d9jb
# HmA9UPjW41GASw/EGjGqBEm69PFffzYlgfuha0oCc6kClCQLri3lqYFK40mHMGhA
# dsMdLEwFU2BIeV7IzKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTExNTE1
# MjBaMC8GCSqGSIb3DQEJBDEiBCBKb5q+xSnEM2rbzjz3ECuCyj01EsXHmv/+LX7z
# mKt+fzANBgkqhkiG9w0BAQEFAASCAgCJPAcDLckA/Hk3C0UmS6kKGc7Cq4NDdOaE
# eHxDpNIBt+TZBaat014tDwyr9rzZqjnteLT054ymZjvjb4MwXwm5hWriEIPIdE+D
# tCB3cDa1RETgc/+x5hf+JceTQkkgOTB96LTPwdIHQ1bgB8O0wgjFdh1tRJ7JaTG1
# sNhjNHSVvmMJ3EUZx2RapHzAiFUAAHhD3wTt8ZH60Kcysr+8H8P0WXSAc1kwCGkJ
# OuWjO0i+OBjhs9wgBJI7FA9PEJCVVUa8on7mJTMEeOYHRFSum2+CQawxOpFSvOa2
# Uj9W3LAaKG7RjUkzg4g+S0qgbkeCnnu5EFXDaGmVjF5kUotrprbq06bmXWqBPh+r
# O9doaHuO+dzu3daTfU9dnw5SguXFiCa6C0RglHuMcYUEn40J+cpQEfKMXcBSkedP
# UAHSNhkK8bPCWrYAfXPgQJldGQ7tTJ4epCUUZ+uQBuE/lmR68K8w1rbdbXATeO2e
# TmwwbMeTbBch+xWSmiDAQbVMOIG69LgpvBueMZMpEJJsDFZlTd39NIR0I1wKCbz0
# JBBl/PY5graspVXos/PZsI0HIzU0wm8E6/L7FfMn1TwD9dw93mE85mNrBRVuACcG
# PJVwXIPNgTM9R/OVDgZORV2d3D+GAbMo5o+bkKxIw3Ormaty9uqIboJpnHXFGGhX
# ZsvYBVhgaQ==
# SIG # End signature block
