<#
.SYNOPSIS
Runs autonomous SmartWorkplaceCMDB safety and CSV contract tests.

.VERSION
0.2.0
#>
[CmdletBinding()]
param()

$ScriptVersion = '0.2.0'
$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$modulePath = Join-Path -Path $projectRoot -ChildPath 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$buildPath = Join-Path -Path $projectRoot -ChildPath 'Build\SmartWorkplaceCMDB-Build.ps1'
$reportPath = Join-Path -Path $projectRoot -ChildPath 'Reports\SmartWorkplaceCMDB-Report.ps1'

$script:Passed = 0
$script:Failed = 0
$script:Failures = New-Object System.Collections.Generic.List[string]

function Invoke-SmartWorkplaceCMDBTest {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [scriptblock]$Test
    )

    try {
        & $Test
        $script:Passed++
        Write-Information ("PASS {0}" -f $Name) -InformationAction Continue
    }
    catch {
        $script:Failed++
        $message = '{0}: {1}' -f $Name, $_.Exception.Message
        $script:Failures.Add($message) | Out-Null
        Write-Information ("FAIL {0}" -f $message) -InformationAction Continue
    }
}

function Assert-SmartWorkplaceCMDBTrue {
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

function Assert-SmartWorkplaceCMDBThrow {
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
    if ([string]$caught.Exception.Message -notmatch $MessagePattern) {
        throw "Exception did not match '$MessagePattern': $($caught.Exception.Message)"
    }
}

Import-Module $modulePath -Force

$tempRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ('SmartWorkplaceCMDB-Tests-{0}' -f [guid]::NewGuid().ToString('N'))
$resolvedTempRoot = [System.IO.Path]::GetFullPath($tempRoot)
$expectedPrefix = [System.IO.Path]::GetFullPath((Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'SmartWorkplaceCMDB-Tests-'))
if (-not $resolvedTempRoot.StartsWith($expectedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Unsafe test root: $resolvedTempRoot"
}

New-Item -ItemType Directory -Path $resolvedTempRoot -Force | Out-Null

try {
    Invoke-SmartWorkplaceCMDBTest -Name 'Valid identity and normalized paths' -Test {
        $paths = Resolve-SmartWorkplaceCMDBTenantPath -Tenant 'Audit-01' -OrganizationKey 'Contoso' -EnvironmentKey 'Prod' -TenantKey 'contoso-prod'
        Assert-SmartWorkplaceCMDBTrue -Condition ($paths.ProfileKey -ceq 'audit-01') -Message 'ProfileKey was not normalized.'
        Assert-SmartWorkplaceCMDBTrue -Condition ($paths.TenantKey -ceq 'contoso-prod') -Message 'TenantKey was not normalized.'
        Assert-SmartWorkplaceCMDBTrue -Condition ([System.IO.Path]::IsPathRooted($paths.DataRootPath)) -Message 'DataRootPath is not absolute.'
    }

    foreach ($invalidKey in @('..\outside', '../outside', 'acme division', 'prod/eu', '-leading', 'trailing-')) {
        $keyForTest = $invalidKey
        Invoke-SmartWorkplaceCMDBTest -Name ("Reject invalid ProfileKey '{0}'" -f $keyForTest) -Test {
            Assert-SmartWorkplaceCMDBThrow -Action {
                Resolve-SmartWorkplaceCMDBTenantPath -Tenant $keyForTest -OrganizationKey 'contoso' -EnvironmentKey 'prod' | Out-Null
            } -MessagePattern 'ProfileKey.*invalid'
        }
    }

    Invoke-SmartWorkplaceCMDBTest -Name 'Reject mismatched TenantKey' -Test {
        Assert-SmartWorkplaceCMDBThrow -Action {
            Resolve-SmartWorkplaceCMDBTenantPath -Tenant 'audit' -OrganizationKey 'contoso' -EnvironmentKey 'prod' -TenantKey 'wrong-key' | Out-Null
        } -MessagePattern 'must equal OrganizationKey-EnvironmentKey'
    }

    Invoke-SmartWorkplaceCMDBTest -Name 'Reject invalid TenantId' -Test {
        Assert-SmartWorkplaceCMDBThrow -Action {
            Resolve-SmartWorkplaceCMDBTenantPath -Tenant 'audit' -OrganizationKey 'contoso' -EnvironmentKey 'prod' -TenantId 'not-a-guid' | Out-Null
        } -MessagePattern 'TenantId.*invalid'
    }

    $configRoot = Join-Path -Path $resolvedTempRoot -ChildPath 'Config'
    $globalConfigPath = Join-Path -Path $configRoot -ChildPath 'SmartWorkplaceCMDB.global.local.json'
    $tenantConfigPath = Join-Path -Path $configRoot -ChildPath 'audit.local.json'
    New-Item -ItemType Directory -Path $configRoot -Force | Out-Null

    [ordered]@{
        ConfigVersion  = '0.1.0'
        ProfileKey     = 'audit'
        OrganizationKey = 'contoso'
        Output         = [ordered]@{
            DataRootPath = (Join-Path -Path $resolvedTempRoot -ChildPath 'ConfiguredData')
        }
    } | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $globalConfigPath -Encoding UTF8

    [ordered]@{
        ConfigVersion = '0.1.0'
        ProfileKey = 'audit'
        OrganizationKey = 'contoso'
        EnvironmentKey = 'prod'
        TenantKey = 'contoso-prod'
        MicrosoftGraph = [ordered]@{
            TenantId = '11111111-1111-1111-1111-111111111111'
        }
    } | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $tenantConfigPath -Encoding UTF8

    Invoke-SmartWorkplaceCMDBTest -Name 'Synchronize and merge runtime configuration' -Test {
        $context = Resolve-SmartWorkplaceCMDBContext -GlobalConfigPath $globalConfigPath -TenantConfigPath $tenantConfigPath
        Assert-SmartWorkplaceCMDBTrue -Condition ($context.Paths.ProfileKey -ceq 'audit') -Message 'Profile configuration was not applied.'
        Assert-SmartWorkplaceCMDBTrue -Condition ($context.Paths.TenantKey -ceq 'contoso-prod') -Message 'Tenant identity configuration was not merged.'
        Assert-SmartWorkplaceCMDBTrue -Condition ($context.Paths.TenantId -ceq '11111111-1111-1111-1111-111111111111') -Message 'TenantId configuration was not applied.'

        $globalRuntime = Get-Content -LiteralPath $globalConfigPath -Raw | ConvertFrom-Json
        $tenantRuntime = Get-Content -LiteralPath $tenantConfigPath -Raw | ConvertFrom-Json
        Assert-SmartWorkplaceCMDBTrue -Condition ($null -ne $globalRuntime.Collection) -Message 'Global runtime configuration was not completed from the template.'
        Assert-SmartWorkplaceCMDBTrue -Condition ($null -ne $tenantRuntime.ExchangeOnline) -Message 'Tenant runtime configuration was not completed from the template.'
    }

    $dataRoot = Join-Path -Path $resolvedTempRoot -ChildPath 'BuildData'
    Invoke-SmartWorkplaceCMDBTest -Name 'Create and validate all 22 CSV contracts' -Test {
        & $buildPath -GlobalConfigPath $globalConfigPath -TenantConfigPath $tenantConfigPath -DataRootPath $dataRoot -NoConfigWrite
        $results = @(Test-SmartWorkplaceCMDBCsvContract -LatestOutputRootPath (Join-Path -Path $dataRoot -ChildPath 'DATA-LAST'))
        Assert-SmartWorkplaceCMDBTrue -Condition ($results.Count -eq 22) -Message "Expected 22 contract results, got $($results.Count)."
        Assert-SmartWorkplaceCMDBTrue -Condition (@($results | Where-Object Status -ne 'Valid').Count -eq 0) -Message 'One or more generated CSV contracts are invalid.'
    }

    Invoke-SmartWorkplaceCMDBTest -Name 'Preserve compatible existing output' -Test {
        $dimDatePath = Join-Path -Path $dataRoot -ChildPath 'DATA-LAST\PowerBI\DimDate.csv'
        Add-Content -LiteralPath $dimDatePath -Value '2026-01-01,2026,Q1,1,January,1' -Encoding UTF8
        & $buildPath -GlobalConfigPath $globalConfigPath -TenantConfigPath $tenantConfigPath -DataRootPath $dataRoot -NoConfigWrite
        $lineCount = @(Get-Content -LiteralPath $dimDatePath).Count
        Assert-SmartWorkplaceCMDBTrue -Condition ($lineCount -eq 2) -Message 'A compatible existing CSV was overwritten.'
    }

    Invoke-SmartWorkplaceCMDBTest -Name 'Reject incompatible existing output' -Test {
        $dimUserPath = Join-Path -Path $dataRoot -ChildPath 'DATA-LAST\PowerBI\DimUser.csv'
        Set-Content -LiteralPath $dimUserPath -Value 'BrokenColumn' -Encoding UTF8
        Assert-SmartWorkplaceCMDBThrow -Action {
            & $buildPath -GlobalConfigPath $globalConfigPath -TenantConfigPath $tenantConfigPath -DataRootPath $dataRoot -NoConfigWrite
        } -MessagePattern 'incompatible CSV contract'
    }

    Invoke-SmartWorkplaceCMDBTest -Name 'Explicit ForceInitialize repairs schema-only output' -Test {
        & $buildPath -GlobalConfigPath $globalConfigPath -TenantConfigPath $tenantConfigPath -DataRootPath $dataRoot -NoConfigWrite -ForceInitialize
        $results = @(Test-SmartWorkplaceCMDBCsvContract -LatestOutputRootPath (Join-Path -Path $dataRoot -ChildPath 'DATA-LAST'))
        Assert-SmartWorkplaceCMDBTrue -Condition (@($results | Where-Object Status -ne 'Valid').Count -eq 0) -Message 'ForceInitialize did not restore the CSV contract.'
    }

    Invoke-SmartWorkplaceCMDBTest -Name 'Generate report from valid autonomous output' -Test {
        $reportOutput = Join-Path -Path $resolvedTempRoot -ChildPath 'SmartWorkplaceCMDB-Overview.html'
        & $reportPath -GlobalConfigPath $globalConfigPath -TenantConfigPath $tenantConfigPath -DataRootPath $dataRoot -OutputPath $reportOutput -NoConfigWrite
        Assert-SmartWorkplaceCMDBTrue -Condition (Test-Path -LiteralPath $reportOutput -PathType Leaf) -Message 'HTML report was not created.'
        $html = Get-Content -LiteralPath $reportOutput -Raw
        Assert-SmartWorkplaceCMDBTrue -Condition ($html -match 'Contract Readiness') -Message 'HTML report does not contain contract readiness.'
    }
}
finally {
    if (Test-Path -LiteralPath $resolvedTempRoot) {
        Remove-Item -LiteralPath $resolvedTempRoot -Recurse -Force
    }
}

Write-Information ("SmartWorkplaceCMDB tests completed. Version={0}; Passed={1}; Failed={2}" -f $ScriptVersion, $script:Passed, $script:Failed) -InformationAction Continue
if ($script:Failed -gt 0) {
    throw ("SmartWorkplaceCMDB tests failed: {0}" -f ($script:Failures -join ' | '))
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDw8i+z3Kai1Le4
# YNP175SFp9KFGPYrDR4scVqc0ewDIqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJovvoPqcjhcQQm19KtK2EBxcYsKs1RitNyJ9kcPb9lLMA0GCSqG
# SIb3DQEBAQUABIIBgKoPW9VJLfiQKz3JGJ3R+Oc+JnVH/nCotJv1LdAvsDTYM9Bs
# zDYZ6SNYc/6+Mp64U9NjOxt7vNQG0Oc4Hk3nVoJZ5G5SS2gMRL8Ihi9i7G8ehgM7
# VEZFfwltbGQv5vc2RPiE3681VLEbgIzz9d0sRjZLxFzqcv6MwEdEwFZLtVFA0i3T
# mWlyHQXz7ORePJkyG10/4E4xjja/HNJZQGr0QHO8lMgGv2bnbeXeyqH4kkoynQtc
# ga8YQxqT4esAHF8moNAk+SoZ0Rj5KLZIU1eXNytf+1snRLdNr9jT6OksEaU4FA7K
# 8KjEfDZYH20IXSjCHtihQjcHfrWhd2hMWPPtXaNJmAD6LGwjgoTY3L+1k+MnayLN
# XmEYOOux6JZpXY581onEGzrpNZlgCAO1wjRmoIm0+IPnrPdykmHcbQ0ACz5EsCIG
# xQR6sYW8eUZzdqB4Tv6L9+kmDMfyuIy080e0cmv+VYzTrxlSCuj5bpDDdq8h3Ey8
# EPastAfC3Ime65iv8qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxNzQ0
# MzBaMC8GCSqGSIb3DQEJBDEiBCDJ8n5uPuEkny79Z7cYLimV6MO6MFtdjlP59ZuC
# KeRwczANBgkqhkiG9w0BAQEFAASCAgBSHvlXc1pJvVk2oWPkW3G/K1mIPeZjZPzg
# 6XlwP3zs3a22rlPKCZePIEOSI7Hk26+mkynyExwvC+vUk7b2OViWESDwFPQXfCqX
# UrRtoOvlZBh/riyP77ros89ieKI9QgM6pwwtHNxc+KztvlSZwF5v9WpG4Gp9HzC0
# LxbWVKaOiSkWYNXVUgtFKqkkoIgljTP9+Lh/5f5U2yC6yHFWgjRujgqz5Nhsqgo2
# HelCnzH1lNeLPH6ehhNRbUcmnWaCV9GLyoKHrhEsLWVZq0Z/l64O0uj8H62cUf5r
# tUxWKHBEIiTLY+/orcJTLMWhINfpxIa/cSTSZh2159VtX9jHTOM/3XUgdqSyUJXW
# 5JidRrtDcnu/DQn4Bk/+6MRTXR3o8WNMWHQM3xNtQnFffIPo6Lq7vz88y6wHK+O/
# yWKIjXtcg0eC7T+vb/JJEdbtj5pBfMas5qzouhohcwqNlkrd7MJcHf374HnTIVpg
# GP8TJJwNLyzi4ypLEJIPkhflNnBqmfa42u7sbdTsPdFNu0zhIYI5i9/j7NkMbZSI
# aS1ekc27jHOeu9CnnNj+8tqx5EU3ZvQiRu7W5WjXOShjUkWMQdzG1bf6OAygS7ym
# 1ede+8fvza/pzNgYLef6AHGfKnELghXmJKR5upiOU77ngTs+AM4Tt3/ODRC9ncSj
# jklfc9gNBA==
# SIG # End signature block
