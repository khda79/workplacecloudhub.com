<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB Entra groups collector and normalizer tests.

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

function Invoke-SmartWorkplaceCMDBEntraGroupTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [scriptblock]$Test
    )

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

function Assert-SmartWorkplaceCMDBEntraGroupTrue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [bool]$Condition,

        [Parameter(Mandatory)]
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBEntraGroupThrow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Action,

        [Parameter(Mandatory)]
        [string]$MessagePattern
    )

    $caught = $null
    try {
        & $Action
    }
    catch {
        $caught = $_
    }
    if ($null -eq $caught) {
        throw "Expected an exception matching '$MessagePattern'."
    }
    if ($caught.Exception.Message -notmatch $MessagePattern) {
        throw "Unexpected exception: $($caught.Exception.Message)"
    }
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$collectorPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraGroups-Collect.ps1'
$normalizerPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraGroups-Normalize.ps1'
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$fixturePath = Join-Path $PSScriptRoot 'Fixtures\EntraGroups.sample.json'
$tempBase = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$tempRoot = Join-Path $tempBase ('SmartWorkplaceCMDB-EntraGroups-Tests-{0}' -f [guid]::NewGuid().ToString('N'))

Import-Module $coreModulePath -Force

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $identityParameters = @{
        Tenant          = 'audit'
        OrganizationKey = 'contoso'
        EnvironmentKey  = 'prod'
        TenantKey       = 'contoso-prod'
        TenantId        = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        NoConfigWrite   = $true
    }

    Invoke-SmartWorkplaceCMDBEntraGroupTest -Name 'Validate offline group collection without output writes' -Test {
        $validateRoot = Join-Path $tempRoot 'ValidateOnly'
        & $collectorPath @identityParameters `
            -DataRootPath $validateRoot `
            -InputJsonPath $fixturePath `
            -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition (-not (Test-Path -LiteralPath $validateRoot)) `
            -Message 'Group collector ValidateOnly created an output folder.'
    }

    $runtimeRoot = Join-Path $tempRoot 'Runtime'
    Invoke-SmartWorkplaceCMDBEntraGroupTest -Name 'Collect two offline Entra groups' -Test {
        $script:CollectionResult = & $collectorPath @identityParameters `
            -DataRootPath $runtimeRoot `
            -InputJsonPath $fixturePath
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition ($script:CollectionResult.GroupCount -eq 2) `
            -Message 'Expected two collected groups.'
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition (Test-Path -LiteralPath $script:CollectionResult.RawLatestOutputPath -PathType Leaf) `
            -Message 'Latest raw Entra groups CSV was not created.'
    }

    Invoke-SmartWorkplaceCMDBEntraGroupTest -Name 'Validate raw group contract and MaxItems' -Test {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtimeRoot 'DATA-LAST') `
            -ContractPath $rawContractPath)
        $groupResult = @($results | Where-Object Name -eq 'Entra_Groups.csv')
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition ($groupResult.Count -eq 1 -and $groupResult[0].Status -eq 'Valid') `
            -Message 'Latest raw Entra groups CSV is not contract-valid.'

        $boundedRoot = Join-Path $tempRoot 'Bounded'
        $bounded = & $collectorPath @identityParameters `
            -DataRootPath $boundedRoot `
            -InputJsonPath $fixturePath `
            -MaxItems 1
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition (@(Import-Csv -LiteralPath $bounded.RawLatestOutputPath).Count -eq 1) `
            -Message 'Group collector did not honor MaxItems 1.'
    }

    Invoke-SmartWorkplaceCMDBEntraGroupTest -Name 'Normalize groups into CMDB_Groups and DimGroup' -Test {
        $script:NormalizationResult = & $normalizerPath @identityParameters -DataRootPath $runtimeRoot
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition ($script:NormalizationResult.GroupCount -eq 2) `
            -Message 'Expected two normalized groups.'
        $cmdbRows = @(Import-Csv -LiteralPath $script:NormalizationResult.CmdbGroupOutputPath)
        $dimRows = @(Import-Csv -LiteralPath $script:NormalizationResult.DimGroupOutputPath)
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition ($cmdbRows.Count -eq 2 -and $dimRows.Count -eq 2) `
            -Message 'Curated group row counts are invalid.'
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition ($dimRows[0].TenantGroupKey -eq $cmdbRows[0].CmdbGroupId) `
            -Message 'DimGroup does not preserve the CMDB group key.'
    }

    Invoke-SmartWorkplaceCMDBEntraGroupTest -Name 'Validate curated group contracts' -Test {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtimeRoot 'DATA-LAST') `
            -ContractPath $curatedContractPath)
        $groupResults = @($results | Where-Object Name -in @('CMDB_Groups.csv', 'DimGroup.csv'))
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition ($groupResults.Count -eq 2 -and @($groupResults | Where-Object Status -ne 'Valid').Count -eq 0) `
            -Message 'Curated group CSV headers are not contract-valid.'
    }

    Invoke-SmartWorkplaceCMDBEntraGroupTest -Name 'Reject group identity mismatch' -Test {
        $mismatchPath = Join-Path $tempRoot 'Entra_Groups_Mismatch.csv'
        $rows = @(Import-Csv -LiteralPath $script:CollectionResult.RawLatestOutputPath)
        $rows[0].TenantKey = 'other-prod'
        $rows | Export-Csv -LiteralPath $mismatchPath -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBEntraGroupThrow -Action {
            & $normalizerPath @identityParameters `
                -DataRootPath (Join-Path $tempRoot 'Mismatch') `
                -RawInputPath $mismatchPath | Out-Null
        } -MessagePattern 'identity mismatch'
    }

    Invoke-SmartWorkplaceCMDBEntraGroupTest -Name 'Reject duplicate group source identifiers' -Test {
        $duplicatePath = Join-Path $tempRoot 'Entra_Groups_Duplicate.csv'
        $rows = @(Import-Csv -LiteralPath $script:CollectionResult.RawLatestOutputPath)
        $rows[1].SourceGroupId = $rows[0].SourceGroupId
        $rows | Export-Csv -LiteralPath $duplicatePath -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBEntraGroupThrow -Action {
            & $normalizerPath @identityParameters `
                -DataRootPath (Join-Path $tempRoot 'Duplicate') `
                -RawInputPath $duplicatePath | Out-Null
        } -MessagePattern 'duplicate SourceGroupId'
    }

    Invoke-SmartWorkplaceCMDBEntraGroupTest -Name 'Keep curated groups unchanged in ValidateOnly' -Test {
        $beforeCmdb = (Get-FileHash $script:NormalizationResult.CmdbGroupOutputPath -Algorithm SHA256).Hash
        $beforeDim = (Get-FileHash $script:NormalizationResult.DimGroupOutputPath -Algorithm SHA256).Hash
        & $normalizerPath @identityParameters -DataRootPath $runtimeRoot -ValidateOnly | Out-Null
        $afterCmdb = (Get-FileHash $script:NormalizationResult.CmdbGroupOutputPath -Algorithm SHA256).Hash
        $afterDim = (Get-FileHash $script:NormalizationResult.DimGroupOutputPath -Algorithm SHA256).Hash
        Assert-SmartWorkplaceCMDBEntraGroupTrue `
            -Condition ($beforeCmdb -eq $afterCmdb -and $beforeDim -eq $afterDim) `
            -Message 'Group normalizer ValidateOnly modified curated output.'
    }
}
finally {
    $resolved = [System.IO.Path]::GetFullPath($tempRoot)
    if ($resolved.StartsWith($tempBase, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolved) -like 'SmartWorkplaceCMDB-EntraGroups-Tests-*' -and
        (Test-Path -LiteralPath $resolved)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB Entra groups tests completed. Version={0}; Passed={1}; Failed={2}" -f
    $ScriptVersion,
    $script:Passed,
    $script:Failed
) -InformationAction Continue

if ($script:Failed -gt 0) {
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBxO/f8r0GiVsDm
# GaT69P5CPC/HI4XjD37ckknMV1sblKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDXO7PnQXFI3UYbXCOCibMo/iAPmrneFuYLEyJ0inG4QMA0GCSqG
# SIb3DQEBAQUABIIBgD74dz44wOuSmg3Ga2F9nX7ipW46RWi2MGa9PsO+XLJibUbQ
# pobtaYNB2QylVot/Nn4pFSRlAsweufMEw8yndP/r5vfxHK3voVlv1mV6qKmIGK21
# RPQjTOye11jB6vJPzpGVNB7edWOWU3OzVgP5hdgwZhvJ6Z4nTCumL4pDEjT5G9Hf
# klMSvVDYmqvwY6h6fYQ9Zks6W+mGhq7GnOdEeWoIZcdzwuzS2pBye+fCzyp8qg+S
# KT8hSOmog36BP5vMoyPTeL79PX1D5nAwr0MKuff1zTZtPvUQKEGr0cYUpPikLXfL
# U1bOiXcqD89LNU/wvvkldwZoWy90ujrOWiXR3bDjfxYAIfATnxyx8Jnuk0M1e0NJ
# dMxgygXDikZJo5UhfI0/r8WFYdaEPpO/qIZHssqixOBic9tyBeA7oKbLB+ads4zm
# /k6PtLU0elfI4rPfM8121LiS5YVqOuIKKRKWlkr09A+J1spbhNzxgIxh2ysnZpX+
# Qkfzp9F+nn4SdQFfm6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxNDA1
# MDVaMC8GCSqGSIb3DQEJBDEiBCCXQwmNxT/d82ukzlstV3FNMfZjSrWY4caUF7kL
# +kTJgzANBgkqhkiG9w0BAQEFAASCAgBEkLCi7jlU0FIeoIgh3C8W27mIt+5oWNzg
# qdeVHMhZtkLY27Y/dLQ/MMLnoqU0Ed4Ge+Kw9r/Im+ff/CXr2cq9FfLpw6CqqZ45
# 6SJ0QcpBQXdbIr1kEIGJEAENSRDzFuCpH199zqQyfaV18QA9b/uLHrgo5s3ec96M
# kZVdA9iB5Q8GbljXN1Y6SB49d6e3hfzDT0zC0JZPwXC2Jg9zFTdM0qvhjuPT0v9a
# 9b1GmpZEMVpppPZJ9o47z7pl2Gn+S4KV+ph0L8e4wFoQAOk3mYDkt3CoQzPP9iq9
# iMwLRf3uzsclwzZONeaMCRkm5RvRQz+Ae77bUlAzc89CxW55g71iRZIN/3vphMGn
# iiMEket2gFlE+wpdcyyaMMsaYA7ACl4nqb5r6J8yD7f6a7/oaeG5VDU2bEorAsU2
# kDGLJtdgDvNpwGAntqBAQoww6RuLNoh+j4mu7Zs+YFGbl1lFBEzOYuumG7Chl57P
# bHA9725BW1kI6kkB6+0fEUL/fLFuYhnNeGUJ2FyI9v+VFcX608RRFKex1eBSySGc
# ZoHNQ2qb52+gbuJZLCd63JcOOfpbCRgmI3wA7REs6daQJV7eOTbVxv5A+ykrIXvw
# ahl/eaOhTWKnEhnqMSSsJW9RdCdn7LyRJjnxtlmMjfM7ITEr2xR4aC3/aZ30iYTi
# 4lxLtWpsLw==
# SIG # End signature block
