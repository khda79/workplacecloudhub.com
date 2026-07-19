<#
.SYNOPSIS
Runs offline tests for the SmartWorkplaceCMDB Active Directory collector.

.VERSION
0.2.0
#>
[CmdletBinding()]
param()

$ScriptVersion = '0.2.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Assert-SmartWorkplaceCMDBAdTrue {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Invoke-SmartWorkplaceCMDBAdTest {
    param([string]$Name, [scriptblock]$Test)
    try {
        & $Test
        $script:Passed++
        Write-Information "[PASS] $Name" -InformationAction Continue
    }
    catch {
        $script:Failed++
        Write-Error "[FAIL] $Name - $($_.Exception.Message)" -ErrorAction Continue
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$collector = Join-Path $projectRoot 'Collectors\ActiveDirectory\SmartWorkplaceCMDB-ActiveDirectory-Collect.ps1'
$normalizer = Join-Path $projectRoot 'Collectors\ActiveDirectory\SmartWorkplaceCMDB-ActiveDirectory-Normalize.ps1'
$orchestrator = Join-Path $projectRoot 'Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1'
$fixture = Join-Path $scriptRoot 'Fixtures\ActiveDirectory.sample.json'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.activedirectory.tables.json'
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
Import-Module $coreModulePath -Force

$tempBase = [IO.Path]::GetTempPath()
$tempRoot = Join-Path $tempBase (
    'SmartWorkplaceCMDB-ActiveDirectory-Tests-{0}' -f [guid]::NewGuid().ToString('N')
)
$runtimeRoot = Join-Path $tempRoot 'Runtime'
$identity = @{
    Tenant = 'test'
    OrganizationKey = 'contoso'
    EnvironmentKey = 'prod'
    TenantKey = 'contoso-prod'
    TenantId = '00000000-0000-0000-0000-000000000001'
    NoConfigWrite = $true
}
$script:Passed = 0
$script:Failed = 0

try {
    Invoke-SmartWorkplaceCMDBAdTest 'Validate fixture mode without AD connectivity' {
        $result = & $collector @identity `
            -DataRootPath $runtimeRoot `
            -InputJsonPath $fixture `
            -ValidateOnly
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition ((($result | Out-String) -match 'OfflineJson') -and
                (($result | Out-String) -match 'DomainCount\s*:\s*2')) `
            -Message 'Fixture validation did not report OfflineJson mode.'
    }

    $script:Collection = $null
    Invoke-SmartWorkplaceCMDBAdTest 'Collect five Active Directory raw tables offline' {
        $script:Collection = & $collector @identity `
            -DataRootPath $runtimeRoot `
            -InputJsonPath $fixture `
            -IncludeGroupMemberships
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $script:Collection.DomainCount -eq 2 -and
                $script:Collection.UserCount -eq 3 -and
                $script:Collection.GroupCount -eq 2 -and
                $script:Collection.ComputerCount -eq 3 -and
                $script:Collection.GroupMembershipCount -eq 3
            ) `
            -Message 'Unexpected Active Directory fixture row counts.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Validate all Active Directory raw contracts' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
                -LatestOutputRootPath (Join-Path $runtimeRoot 'DATA-LAST') `
                -ContractPath $rawContractPath |
                Where-Object Name -like 'ActiveDirectory_*.csv')
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $results.Count -eq 5 -and
                @($results | Where-Object Status -ne 'Valid').Count -eq 0
            ) `
            -Message 'One or more Active Directory raw contracts are invalid.'
    }

    $script:Normalization = $null
    Invoke-SmartWorkplaceCMDBAdTest 'Normalize Active Directory source tables without AD module' {
        $script:Normalization = & $normalizer @identity -DataRootPath $runtimeRoot
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $script:Normalization.DomainCount -eq 2 -and
                $script:Normalization.UserCount -eq 3 -and
                $script:Normalization.GroupCount -eq 2 -and
                $script:Normalization.ComputerCount -eq 3 -and
                $script:Normalization.GroupMembershipCount -eq 3
            ) `
            -Message 'Unexpected normalized Active Directory row counts.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Validate all normalized Active Directory contracts' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
                -LatestOutputRootPath (Join-Path $runtimeRoot 'DATA-LAST') `
                -ContractPath $curatedContractPath)
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $results.Count -eq 5 -and
                @($results | Where-Object Status -ne 'Valid').Count -eq 0
            ) `
            -Message 'One or more normalized Active Directory contracts are invalid.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Create stable tenant-scoped Active Directory keys' {
        $userPath = Join-Path $script:Normalization.CuratedOutputRootPath 'CMDB_ActiveDirectoryUsers.csv'
        $membershipPath = Join-Path $script:Normalization.CuratedOutputRootPath 'CMDB_ActiveDirectoryGroupMemberships.csv'
        $users = @(Import-Csv -LiteralPath $userPath)
        $memberships = @(Import-Csv -LiteralPath $membershipPath)
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $users[0].CmdbAdUserId -match '^contoso-prod\\|ad-user\\|' -and
                $memberships[0].CmdbAdGroupId -match '^contoso-prod\\|ad-group\\|' -and
                $memberships[0].CmdbAdMemberId -match '^contoso-prod\\|ad-(user|computer)\\|' -and
                @($users | Where-Object DomainDnsRoot -eq 'child.example.invalid').Count -eq 1 -and
                @($memberships | Where-Object DomainDnsRoot -eq 'child.example.invalid').Count -eq 1
            ) `
            -Message 'Active Directory normalized keys are not stable and tenant-scoped.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Honor MaxItems in an isolated output root' {
        $boundedRoot = Join-Path $tempRoot 'Bounded'
        $result = & $collector @identity `
            -DataRootPath $boundedRoot `
            -InputJsonPath $fixture `
            -IncludeGroupMemberships `
            -MaxItems 1
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $result.UserCount -eq 1 -and
                $result.ComputerCount -eq 1 -and
                $result.GroupMembershipCount -eq 1
            ) `
            -Message 'MaxItems did not bound the Active Directory fixture output.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Run the Active Directory orchestrator pipeline offline' {
        $orchestratedRoot = Join-Path $tempRoot 'Orchestrated'
        $result = & $orchestrator @identity `
            -DataRootPath $orchestratedRoot `
            -Pipeline ActiveDirectory `
            -FixtureRootPath (Split-Path -Parent $fixture)
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition ($result.StepCount -eq 2 -and $result.FailedStepCount -eq 0) `
            -Message 'The offline Active Directory orchestrator pipeline did not complete both steps.'
    }

    Invoke-SmartWorkplaceCMDBAdTest 'Validate centralized Active Directory launchers' {
        $launcherRoot = Join-Path $projectRoot 'Launchers\ActiveDirectory'
        $collectLauncher = Join-Path $launcherRoot 'Start-SmartWorkplaceCMDB-ActiveDirectory-Collect.cmd'
        $validateLauncher = Join-Path $launcherRoot 'Start-SmartWorkplaceCMDB-ActiveDirectory-Validate.cmd'
        $collectText = Get-Content -LiteralPath $collectLauncher -Raw
        $validateText = Get-Content -LiteralPath $validateLauncher -Raw
        Assert-SmartWorkplaceCMDBAdTrue `
            -Condition (
                $collectText -match '-Pipeline ActiveDirectory' -and
                $collectText -match '-Collect' -and
                $validateText -match '-Pipeline ActiveDirectory' -and
                $validateText -match '-ValidateOnly' -and
                $collectText -match '%\*' -and
                $validateText -match '%\*'
            ) `
            -Message 'The centralized Active Directory launchers are incomplete.'
    }
}
finally {
    $resolvedTempRoot = [IO.Path]::GetFullPath($tempRoot)
    if ($resolvedTempRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolvedTempRoot) -like 'SmartWorkplaceCMDB-ActiveDirectory-Tests-*' -and
        (Test-Path -LiteralPath $resolvedTempRoot)) {
        Remove-Item -LiteralPath $resolvedTempRoot -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB Active Directory tests completed. Version={0}; Passed={1}; Failed={2}" -f
    $ScriptVersion, $script:Passed, $script:Failed
) -InformationAction Continue
if ($script:Failed -gt 0) {
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBuKDuuoyzYWfMi
# cgkDFFhE8KwNp9m0pUlZMve1lqaiNaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBOVNFQRXspxkoNqpxnuLY9PF98MqhHBeIlTnfjC6KHZMA0GCSqG
# SIb3DQEBAQUABIIBgCDdEYNjnkQdF/vkQ8ppYSi3WrpHzmc080FZicq0BF8sx1Cb
# hJOhwVsil1IvY1SarAoCFuiBOLthWwm4Ck1w4vXTzR6h6UAokem8bMo4KkAUL07h
# xcpQVxPxsCxxEo4Vlsf5YIHKJzE+RXPOfYPPUfzHeAXYV1yg2yQChCwYKiYWX/kB
# mEuntPGQBKzDVJ8HsCW3zeA1CnmFf3GMEdCsZ+VVcP6AE9k6jgI0xWFUp2S0IhIK
# yZwj2mSN5djlC1Og7byFAT3NhXBsdwYqHQylr0mHXIdE4D6+xHuSHPk0Atm5HLxx
# K1StIXhV7rlxNmVMJN+YzshEgsT/cThpUwIkhbO/PoBzUz8C1RJFTDeXssZBydu5
# tcrb5iEodFOKhEBCa+oIuC7uOfN1xfENm1jc9n8CXvBWnBqNhiF5wbW9S0dS8mKg
# YaDzap0oBzy4qjlMLO9ffE+c7PLvHZYpIRWK1MWQ1k9+eLVJtdMC4aOFUKi/NMaW
# gvlrLg2lBIrNlc+v0KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMjM2
# MTdaMC8GCSqGSIb3DQEJBDEiBCA+miAGhrKsAA9PiaPjKCnBId42AFubcmSBQ6vl
# EUybXzANBgkqhkiG9w0BAQEFAASCAgCWobmscKgz0vCFr0DElVTmxUxgRgaj8b3Z
# n7VlDycPlhqVcZPvoJKsOYIfsN8XAofkgjRsR8JqoJfV+PW75E9OlbKmh0/UmOo3
# Jzru36PdsPNVtTga1d8XgDziHNaeFfIJVmVYYxTTJTwmrrjmsnlzH8ihAI+syC6r
# Wbw+E025pGXq1NNATzLzmeyvACnyVehoiJio+2+KUK9n8/COD4lKJgQk+FtiEdQm
# rzIkrOXtQxJBU/R6AwGCHXgW/8gVdb44z4rKc9GDUZluNIlyiQaFl7WxNzjrUrGN
# XUoadXD0SjTZISrsdnGupnoy/k0WlAjwsAmHnYWlRL7ZIy2uIS3C4C6oM0yDZcqw
# pnUBosI5xehTK6MqfR3Es/mejwRu4BWQ+OItJplaNlvtDcU+gI659X82mx9O+dS0
# NGpk1oUiVsKL9X/IXewsl0/ONg78JaBJyA5z9JYmwUW2D0BCX5Rta094i4IxhyMX
# fiho15tCo1a/QhYTHTxdqZaWqkRgosWUynykHqDuPN31fqC5XQeOoj+/13doc7GF
# EAsFhaP7+vw1+q4X2GEYp7uQ75YcV3IRVoJRkAUGZ36lXGz/TYCJxpehS9NjaPgL
# 3/oTTE+PFazZiTvdgyAvAWJOdxrJrZ/3J8QC5FpZwdM6ai7ZhtHt+60iIjfAc13M
# 0SER/Cedjw==
# SIG # End signature block
