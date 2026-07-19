<#
.SYNOPSIS
Runs offline Microsoft 365 user license assignment pipeline tests.

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

function Invoke-SmartWorkplaceCMDBUserLicenseTest {
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

function Assert-SmartWorkplaceCMDBUserLicenseTrue {
    [CmdletBinding()]
    param([Parameter(Mandatory)][bool]$Condition, [Parameter(Mandatory)][string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-SmartWorkplaceCMDBUserLicenseThrow {
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Action, [Parameter(Mandatory)][string]$MessagePattern)
    $caught = $null
    try { & $Action } catch { $caught = $_ }
    if ($null -eq $caught) { throw "Expected an exception matching '$MessagePattern'." }
    if ($caught.Exception.Message -notmatch $MessagePattern) {
        throw "Unexpected exception: $($caught.Exception.Message)"
    }
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$collector = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Collect.ps1'
$normalizer = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Normalize.ps1'
$userCollector = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Collect.ps1'
$userNormalizer = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Normalize.ps1'
$skuCollector = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Collect.ps1'
$skuNormalizer = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Normalize.ps1'
$coreModule = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContract = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContract = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$fixture = Join-Path $PSScriptRoot 'Fixtures\M365UserLicenseAssignments.sample.json'
$userFixture = Join-Path $PSScriptRoot 'Fixtures\EntraUsers.sample.json'
$skuFixture = Join-Path $PSScriptRoot 'Fixtures\M365SubscribedSkus.sample.json'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$tempRoot = Join-Path $tempBase ('SmartWorkplaceCMDB-UserLicense-Tests-{0}' -f [guid]::NewGuid().ToString('N'))
Import-Module $coreModule -Force

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $identity = @{
        Tenant = 'audit'; OrganizationKey = 'contoso'; EnvironmentKey = 'prod'
        TenantKey = 'contoso-prod'; TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        NoConfigWrite = $true
    }

    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Collector ValidateOnly stays read-only' {
        $root = Join-Path $tempRoot 'Validate'
        & $collector @identity -DataRootPath $root -InputJsonPath $fixture -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBUserLicenseTrue (-not (Test-Path $root)) 'ValidateOnly created output.'
    }

    $runtime = Join-Path $tempRoot 'Runtime'
    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Prepare references and collect raw assignments' {
        & $userCollector @identity -DataRootPath $runtime -InputJsonPath $userFixture | Out-Null
        & $userNormalizer @identity -DataRootPath $runtime | Out-Null
        & $skuCollector @identity -DataRootPath $runtime -InputJsonPath $skuFixture | Out-Null
        & $skuNormalizer @identity -DataRootPath $runtime | Out-Null
        $script:Collection = & $collector @identity -DataRootPath $runtime -InputJsonPath $fixture
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            $script:Collection.UserCount -eq 3 -and
            $script:Collection.AssignmentCount -eq 3
        ) 'Expected three users and three raw assignment states.'
    }

    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Validate raw contract and user-bounded MaxItems' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtime 'DATA-LAST') `
            -ContractPath $rawContract)
        $assignment = @($results | Where-Object Name -eq 'M365_UserLicenseAssignments.csv')
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            $assignment.Count -eq 1 -and $assignment[0].Status -eq 'Valid'
        ) 'Raw assignment contract is invalid.'
        $bounded = & $collector @identity `
            -DataRootPath (Join-Path $tempRoot 'Bounded') `
            -InputJsonPath $fixture `
            -MaxItems 1
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            $bounded.UserCount -eq 1 -and $bounded.AssignmentCount -eq 2
        ) 'MaxItems did not limit users as documented.'
    }

    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Normalize one fact row per user and SKU' {
        $script:Normalization = & $normalizer @identity -DataRootPath $runtime
        $facts = @(Import-Csv $script:Normalization.FactUserLicenseOutputPath)
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            $script:Normalization.RawAssignmentCount -eq 3 -and
            $script:Normalization.FactAssignmentCount -eq 2 -and
            $script:Normalization.CollapsedAssignmentCount -eq 1
        ) 'Fact assignment counters are invalid.'
        Assert-SmartWorkplaceCMDBUserLicenseTrue ($facts.Count -eq 2) 'Expected two fact rows.'
    }

    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Select the most critical assignment state' {
        $facts = @(Import-Csv $script:Normalization.FactUserLicenseOutputPath)
        $first = @($facts | Where-Object CmdbUserId -eq 'contoso-prod|entra-user|11111111-1111-1111-1111-111111111111')
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            $first.Count -eq 1 -and
            $first[0].AssignmentState -eq 'ActiveWithError' -and
            $first[0].AssignedDateTime -like '2026-07-02T08:00:00*'
        ) 'Critical assignment selection is invalid.'
    }

    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Validate fact contract, keys, and references' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtime 'DATA-LAST') `
            -ContractPath $curatedContract)
        $fact = @($results | Where-Object Name -eq 'FactUserLicense.csv')
        $rows = @(Import-Csv $script:Normalization.FactUserLicenseOutputPath)
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            $fact.Count -eq 1 -and $fact[0].Status -eq 'Valid'
        ) 'FactUserLicense contract is invalid.'
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            @($rows | Group-Object TenantUserKey,TenantSkuKey | Where-Object Count -gt 1).Count -eq 0 -and
            $script:Normalization.OrphanUserCount -eq 0 -and
            $script:Normalization.OrphanSkuCount -eq 0
        ) 'Fact keys or references are invalid.'
    }

    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Report direct and group assignment counts' {
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            $script:Normalization.DirectAssignmentCount -eq 1 -and
            $script:Normalization.GroupAssignmentCount -eq 2
        ) 'Direct and group assignment counts are invalid.'
    }

    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Reject identity mismatch and duplicate raw keys' {
        $rows = @(Import-Csv $script:Collection.RawLatestOutputPath)
        $mismatch = Join-Path $tempRoot 'Mismatch.csv'
        $rows[0].TenantKey = 'other-prod'
        $rows | Export-Csv $mismatch -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBUserLicenseThrow {
            & $normalizer @identity `
                -DataRootPath (Join-Path $tempRoot 'Mismatch') `
                -RawInputPath $mismatch | Out-Null
        } 'identity mismatch'
        $rows = @(Import-Csv $script:Collection.RawLatestOutputPath)
        $duplicate = Join-Path $tempRoot 'Duplicate.csv'
        $rows[1].RawAssignmentKey = $rows[0].RawAssignmentKey
        $rows | Export-Csv $duplicate -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBUserLicenseThrow {
            & $normalizer @identity `
                -DataRootPath (Join-Path $tempRoot 'Duplicate') `
                -RawInputPath $duplicate | Out-Null
        } 'duplicate RawAssignmentKey'
    }

    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Publish valid schema-only fact for no assignments' {
        $emptyFixture = Join-Path $tempRoot 'Empty.json'
        '{"value":[]}' | Set-Content -LiteralPath $emptyFixture -Encoding UTF8
        $emptyRoot = Join-Path $tempRoot 'Empty'
        $emptyCollection = & $collector @identity -DataRootPath $emptyRoot -InputJsonPath $emptyFixture
        $emptyNormalization = & $normalizer @identity -DataRootPath $emptyRoot
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $emptyRoot 'DATA-LAST') `
            -ContractPath $curatedContract)
        $fact = @($results | Where-Object Name -eq 'FactUserLicense.csv')
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            $emptyCollection.AssignmentCount -eq 0 -and
            $emptyNormalization.FactAssignmentCount -eq 0 -and
            $fact[0].Status -eq 'Valid'
        ) 'Empty fact output is invalid.'
    }

    Invoke-SmartWorkplaceCMDBUserLicenseTest 'Normalizer ValidateOnly preserves fact output' {
        $before = (Get-FileHash $script:Normalization.FactUserLicenseOutputPath -Algorithm SHA256).Hash
        & $normalizer @identity -DataRootPath $runtime -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBUserLicenseTrue (
            $before -eq (Get-FileHash $script:Normalization.FactUserLicenseOutputPath -Algorithm SHA256).Hash
        ) 'ValidateOnly modified FactUserLicense.'
    }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolved) -like 'SmartWorkplaceCMDB-UserLicense-Tests-*' -and
        (Test-Path $resolved)) {
        Remove-Item $resolved -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB user license assignments tests completed. Version={0}; Passed={1}; Failed={2}" -f
    $ScriptVersion, $script:Passed, $script:Failed
) -InformationAction Continue
if ($script:Failed -gt 0) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAcwkXvXalOt9AQ
# nnBvR1IhA1xb+LfctyUPRlTQ1V7aWKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMFVNhaQmqKcTf8FD+YPFlTyC9i5zLBfwCsZv5rx4xEuMA0GCSqG
# SIb3DQEBAQUABIIBgDcLp5efytGQ4sRwQ1b50QAp1RJvKqnV8HmymaZ5vozwRKt+
# HwmfBZAkLd/NNBdVS/mLFmnM+aEZCjzV/x7TRoXuG6rzrrOQHgsx+rTsKMMSpM0n
# y+bwjYV+Rgk7iYcZ27WrhL3x4KOtCGSyDD8fj+ta81gr9f8zRbuaMOP73dpVAZ5j
# GodDeg+9MJ+/rrvXhFhs73QlmyH8Y05ZuDXUCEIy1h0MEXxcHN3qQl8+kBlF4NX1
# C6xPB3RpLbK6S52YB5NYBkVNcmjQ/SZQrMLuA8JUXYkkNrF/ZyuTAD3RFWn8ALfv
# tKyLtypr5DhSyUSXcJsQwGZEKekz26QhZl5nbm5uHncVCTNhg+ifKfi25/YgDo+t
# 7lvmrC2GXsizC/pO1atrme0T1L0yiuvIAfS7gcdUZAnDbLODzoM2ERYJ9o3OPlqP
# Y47PljDed6UOvO5tBxYug/HBwpBBuCCkRuHj+pyGPRoKUcN5tYXtkBTfTk1NBJYP
# 9AZ2AtFRzrdb4I/0Q6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxOTA0
# MzFaMC8GCSqGSIb3DQEJBDEiBCCztiVs32poHbSWounlggMvqIrplgkYllrvlIkh
# t5QrdDANBgkqhkiG9w0BAQEFAASCAgCsDYzTi0RId1lcQ05u8aoVoYNCrj9FJuMv
# Hvuldm6I0Hn3z+xv7SNA4ja99PaRIZ4wtDHZuI9c88NprCYLC4+/ekcptCUkjoft
# N3QuJWnf5Xz4Rv0cC0HqDXl67UwoJKb/g7gf+4cybtChcNMosP+VCfR/3msrhnd5
# 6sG7RISg/LxTSszF6NduXQoUcOGbFyNH9xsESCrDMBSTir4E8qFuw64LeYDuWTS5
# w2RZw9dWBnvVe9yQu+MmzEfm4RJuGNcMLmhf2JKccT1klZMKPq8p31mp47u9qbO3
# Nke7koD2d3FkxyqX9r/8bCo2Eexr1oOBf8M5E+1gkLEB40svW8maP45b9wnO2TFk
# +fh7PeJaFhixiqOmHlDq40orKIWOw1SlnMNLKDSSLSfkjJ/nDsSb8C79ZVAAE+EZ
# 9nhf8Ce8U5IxSY+DbfwIgvvMDd2hYpFxCF9GQIbS62tY00x6fdMEkEb7rJYZa2Tz
# vyjLcVV7jkrEsTdy5Be+t5vWLWmuG8idOCyCzGp0Y5YV/dnCi2YUMEEOGYmlEhaH
# GXITJMudlJmJqU/OPxEgDMbxiUhlh8IxjtJANtZ0vYt7l0EK2EIraOeS7xwEnHYF
# dg710qy6b2xnRfMhJbm/2eulkKHUv01NFUyLaPr5h+pDDeN2vqoFu4h7iF+fDSnJ
# bt49lRAbyg==
# SIG # End signature block
