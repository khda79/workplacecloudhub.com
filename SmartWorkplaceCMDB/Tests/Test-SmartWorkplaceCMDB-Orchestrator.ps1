<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB orchestrator and launcher tests.

.VERSION
0.1.2-beta.1
#>
[CmdletBinding()]
param()

$ScriptVersion = '0.1.2-beta.1'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:Passed = 0
$script:Failed = 0

function Invoke-SmartWorkplaceCMDBOrchestratorTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Test
    )
    try {
        & $Test
        $script:Passed++
        Write-Information "PASS $Name" -InformationAction Continue
    }
    catch {
        $script:Failed++
        Write-Information "FAIL $Name - $($_.Exception.Message)" `
            -InformationAction Continue
    }
}

function Assert-SmartWorkplaceCMDBOrchestratorTrue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBOrchestratorThrow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Body,
        [Parameter(Mandatory)][string]$ExpectedText
    )
    try {
        & $Body
    }
    catch {
        if ($_.Exception.Message -notlike "*$ExpectedText*") {
            throw "Unexpected exception: $($_.Exception.Message)"
        }
        return
    }
    throw "Expected exception containing '$ExpectedText'."
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$orchestrator = Join-Path $projectRoot 'Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1'
$fixtureRoot = Join-Path $PSScriptRoot 'Fixtures'
$modulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$contractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$adContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.activedirectory.tables.json'
$launcherRoot = Join-Path $projectRoot 'Launchers\Cloud'
Import-Module $modulePath -Force

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) (
    'SmartWorkplaceCMDB-Orchestrator-Tests-' +
    [guid]::NewGuid().ToString('N')
)
$identity = @{
    Tenant = 'test'
    OrganizationKey = 'contoso'
    EnvironmentKey = 'prod'
    TenantKey = 'contoso-prod'
    TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
    NoConfigWrite = $true
}

try {
    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Default validation stays read-only' {
        $validateRoot = Join-Path $tempRoot 'Validate'
        $result = & $orchestrator @identity `
            -DataRootPath $validateRoot `
            -FixtureRootPath $fixtureRoot `
            -ValidateOnly
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($result.Status -eq 'Validated' -and
                $result.Mode -eq 'Validate' -and
                $result.StepCount -eq 10 -and
                -not (Test-Path $validateRoot)) `
            'Default validation mode wrote output or returned invalid status.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Run complete offline fixture pipeline' {
        $runtimeRoot = Join-Path $tempRoot 'Runtime'
        $script:FullResult = & $orchestrator @identity `
            -DataRootPath $runtimeRoot `
            -FixtureRootPath $fixtureRoot
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($script:FullResult.Status -eq 'Completed' -and
                $script:FullResult.Mode -eq 'Fixture' -and
                $script:FullResult.StepCount -eq 22 -and
                $script:FullResult.FailedStepCount -eq 0) `
            'Full fixture orchestration status is invalid.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Publish all contracts and report' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $script:FullResult.LatestOutputRootPath `
            -ContractPath $contractPath)
        $rawResults = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $script:FullResult.LatestOutputRootPath `
            -ContractPath $rawContractPath)
        $adResults = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath $script:FullResult.LatestOutputRootPath `
            -ContractPath $adContractPath)
        $reportPath = Join-Path `
            $script:FullResult.LatestOutputRootPath `
            'SmartWorkplaceCMDB-Overview.html'
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($results.Count -eq 20 -and
                @($results | Where-Object Status -ne 'Valid').Count -eq 0 -and
                $rawResults.Count -eq 12 -and
                @($rawResults | Where-Object Status -ne 'Valid').Count -eq 0 -and
                $adResults.Count -eq 5 -and
                @($adResults | Where-Object Status -ne 'Valid').Count -eq 0 -and
                (Test-Path $reportPath -PathType Leaf)) `
            'Full fixture orchestration did not publish the complete model.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Write auditable step log' {
        $rows = @(Import-Csv $script:FullResult.LogPath)
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($rows.Count -eq 22 -and
                @($rows | Where-Object Status -ne 'Completed').Count -eq 0 -and
                @($rows | Group-Object Sequence |
                    Where-Object Count -gt 1).Count -eq 0) `
            'Orchestrator step log is incomplete or inconsistent.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Run bounded individual pipeline' {
        $boundedRoot = Join-Path $tempRoot 'Bounded'
        $result = & $orchestrator @identity `
            -DataRootPath $boundedRoot `
            -FixtureRootPath $fixtureRoot `
            -Pipeline EntraUsers `
            -MaxItems 1
        $rows = @(Import-Csv (
                Join-Path `
                    $result.LatestOutputRootPath `
                    'CMDB\CMDB_Users.csv'
            ))
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($result.StepCount -eq 2 -and
                $rows.Count -eq 1 -and
                $result.DataRootPath.StartsWith(([IO.Path]::GetFullPath($boundedRoot) + '\TestRuns\'), [StringComparison]::OrdinalIgnoreCase)) `
            'Bounded individual pipeline is invalid.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Reject unsafe mode combinations' {
        Assert-SmartWorkplaceCMDBOrchestratorThrow {
            & $orchestrator @identity `
                -Collect `
                -ValidateOnly | Out-Null
        } 'cannot be used together'
        Assert-SmartWorkplaceCMDBOrchestratorThrow {
            & $orchestrator @identity `
                -FixtureRootPath $fixtureRoot `
                -Pipeline Full `
                -MaxItems 1 | Out-Null
        } 'requires an individual source pipeline'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Reject missing fixtures before output' {
        $missingFixtureRoot = Join-Path $tempRoot 'MissingFixtures'
        New-Item -ItemType Directory -Path $missingFixtureRoot -Force | Out-Null
        $missingOutputRoot = Join-Path $tempRoot 'MissingOutput'
        Assert-SmartWorkplaceCMDBOrchestratorThrow {
            & $orchestrator @identity `
                -DataRootPath $missingOutputRoot `
                -FixtureRootPath $missingFixtureRoot `
                -Pipeline EntraUsers `
                -ValidateOnly | Out-Null
        } 'Required fixture is missing'
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            (-not (Test-Path $missingOutputRoot)) `
            'Missing fixture validation created output.'
    }

    Invoke-SmartWorkplaceCMDBOrchestratorTest 'Validate centralized Cloud launchers' {
        $expected = [ordered]@{
            'Start-SmartWorkplaceCMDB-Full-Validate.cmd' = '-ValidateOnly'
            'Start-SmartWorkplaceCMDB-Full-Collect.cmd' = '-Collect'
            'Start-SmartWorkplaceCMDB-EntraUsers.cmd' = '-Pipeline EntraUsers'
            'Start-SmartWorkplaceCMDB-EntraGroups.cmd' = '-Pipeline EntraGroups'
            'Start-SmartWorkplaceCMDB-EntraDevices.cmd' = '-Pipeline EntraDevices'
            'Start-SmartWorkplaceCMDB-IntuneDevices.cmd' = '-Pipeline IntuneDevices'
            'Start-SmartWorkplaceCMDB-M365SubscribedSkus.cmd' = '-Pipeline M365SubscribedSkus'
            'Start-SmartWorkplaceCMDB-M365UserLicenses.cmd' = '-Pipeline M365UserLicenses'
            'Start-SmartWorkplaceCMDB-ExchangeOnlineMailboxes.cmd' = '-Pipeline ExchangeOnlineMailboxes'
            'Start-SmartWorkplaceCMDB-CuratedOnly.cmd' = '-Pipeline CuratedOnly'
        }
        $files = @(Get-ChildItem $launcherRoot -Filter '*.cmd' -File)
        $forbiddenPattern = @(('Smart' + 'M365'),
            ('SmartWorkplace' + 'Dashboard')) -join '|'
        $invalid = 0
        foreach ($entry in $expected.GetEnumerator()) {
            $path = Join-Path $launcherRoot $entry.Key
            if (-not (Test-Path $path -PathType Leaf)) {
                $invalid++
                continue
            }
            $content = Get-Content $path -Raw
            if ($content -notlike '*PowerShell\7\pwsh.exe*' -or
                $content -notlike '*SmartWorkplaceCMDB-Orchestrator.ps1*' -or
                $content -notlike "*$($entry.Value)*" -or
                $content -notlike '*%**' -or
                $content -match $forbiddenPattern) {
                $invalid++
            }
        }
        Assert-SmartWorkplaceCMDBOrchestratorTrue `
            ($files.Count -eq $expected.Count -and $invalid -eq 0) `
            'Centralized Cloud launcher contract is invalid.'
    }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path $resolved)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB orchestrator tests completed. Version={0}; Passed={1}; Failed={2}" -f
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA6joOZWsosJqXD
# jQDaYr4T9xS9FTgxoYYNJn7K8H+MPKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMYGvawp/nQg4KvMGiUHfDGSYzB5QlX5JnJrndJ2pmrIMA0GCSqG
# SIb3DQEBAQUABIIBgKsVofJgVH6fhPHqN2ZfIRgakBY0egBTk/jYWhpRSF1le2Rr
# CH3yUWpOdAmpR8FNgcDRhT+HW2ZK3AuIKHz/gqq1fJb14NKa89jIGZ0gyelJ9KUy
# erdUzC+BZLfyMYgrgha7GAdYb6efx56qtNHHwnvTm0wPRy7T8duPkVBTHQ10wxkf
# TYY0KlymX2p9/m9tRhd4MVQ3x5DxMwy8oJoyaTPdWO4OYFXfWI2a6udfNetKG0cx
# 0Oj/YKQsxXOmEBYoxtuB6GMcwWNstD+C9snjK05IVEKQxrj5sCYqLX4Gdc5Cttj4
# e87ckUdyVFs5wW6uFBjkS/Aj8iIsJGjCqV8boXUDOpgYKTOSwoS/N0h81cqycSC0
# fnODqhIUYtajUls6xUT5En3RYps3k9eBzbAiNhcRzyREmgduemOGY1m1uvZwgzW5
# bC+TA/oiH5qTiRN0E1d5XjxgopMQhUWXestMuRe0aPEDtYy+TKURbiG9Vjoh/woo
# E8GhYaElw90WbseERKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDkxNDIy
# MTBaMC8GCSqGSIb3DQEJBDEiBCCRujd8YCHuyR5QTDM/PnV3F3mii+sBZTio8gFI
# XruoQjANBgkqhkiG9w0BAQEFAASCAgAUV3BGzJL2cn8ecFh6DcXjfAfh6t5l1MUB
# Hr3dv6LH9qiKA8zjQIRda7vWJV6XGAUgUCRcEXHWc1Igg1Lzu3NJTrT5jn6Okz48
# HH788JIvVE8a5aWzCM8y/Yelp2tkoeHnfxyzvkKmQsv8qLywJbldTfNm5ypk6aZT
# SAJzTrmlsZE9uCTokmz821BsTU7aPYB0wth7KtzC6Qtph3CIUJxp3OCYX1wjDBMQ
# gOBRSozVMExtyxwk35TblhK/QaoslSZWltLMu7fHLwuVX6OcafPbgtMpZzYoPAe1
# LtpA9mzyYhOxodoqIxyZO631KEeccwLBjHxHUmxvseHXR6zVvpVXmLcNf1L3nBFw
# CzNlVMuosACGvsNSv9CDBE1MoOWpbOrjixJyyxVZmEgTL9JToMHJMMFRXRmdbNJf
# DC/PHrddHlr6n+4hegEIB23i/XN2G0lZ6wy+zDefeN7TvygqGXR9UTaRAmdHg6/e
# cYOzVCII4LzAyk40GZoQP76DJWdHWav+M0bT3qXye0O0lRfYTSgcIqMpniPz3t5Y
# XHQuoPLIMz0UjYXx432IQUAkGC6OuqiItDJvc6/sKYmd01L/EgwJ6/fex2fz2UVk
# oELEXtDCDG+Kguf4qVFYdGyGsPCrt9kIDpNqcBo+X2moBjvVgveDq2KGCBdIzM4g
# r92gyFdG9w==
# SIG # End signature block
