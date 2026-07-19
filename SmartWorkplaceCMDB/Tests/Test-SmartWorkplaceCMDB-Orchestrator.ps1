<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB orchestrator and launcher tests.

.VERSION
0.1.1
#>
[CmdletBinding()]
param()

$ScriptVersion = '0.1.1'
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
                $result.DataRootPath -eq [IO.Path]::GetFullPath($boundedRoot)) `
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCx1htILlI2hr8e
# l+bM4NcdT6uePrs+P0oL8Y/5nPT4jaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPbVE6hfYd2ZxkI1GOaEoekOoDj9cX+uGtePH3EJ96huMA0GCSqG
# SIb3DQEBAQUABIIBgC28eHn8t8n6dB44uqb6aKg0zWUH9/fj3Glmhq7NZ8q2FtnC
# qELG85QLeEBBcpE7KDeqZDzw/3K3Hd/FFPC89/lTH4staAgzWLJUx50bXa6tDnxu
# qCHgR6cvWP0UCuUwdvv4zAmjGpE3VqUQErrvPaQWI8MMNkg7LWc+TSptXtroTtn7
# VCMo4lVYsD6ZHNbGry7/TTVwLrcvLe1h3k6Nj2ZYFX5fQl5gtknDBaHEZardMLCg
# FFYBL7+qxxss1N2UsOyylsHvIo+2dCCAmu7zaobkyfTtF46AHBluhLe66eLuVcGr
# GCKQitMATAQUkbyM7oolcG3tOhNUOE8P/BFEovQUSmrbhE5+ZsdbVNWqGaXVcClu
# 0/OSrcVU1FENikYnm3Ot+HtU2g7Sa6bvcUTemVtL7gQCMWd72kRWBuI8pPqYeEv1
# hMauEKLoP4Kp5Uk9V55PR5kbLSwmpGqrf28qNA3537l5jgBSTQxwKmS1R1mKsaHu
# Vtm+b5DI4a5ySxAd46GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMTEy
# NDFaMC8GCSqGSIb3DQEJBDEiBCADN01Gg7sCJIA4Nl10uv79TKIch+ZT1hWEn4i0
# YRdhFzANBgkqhkiG9w0BAQEFAASCAgAFPPuuNwoqKGHoPIjDee+r+8xhqpifgtYx
# qmgox9HyjI1xQz56IvOYLa4/A7A4+RAEnw5C9xUM/XQCJd5KNOTZreQgUJeD+sgQ
# 9Pzho6aZc0T78DYR3RpZz88uTtotn63SqYSqotJ7EUdQb20WdqqGpFJprwp6m1+H
# iF/Y7RKc5aFLn4s82W29sGSPwIaoXR3tf968uRNj0isrhEcIhjRyawbK5OcGG0ZG
# jOSwgVVtP5WjgsSv5gsvDKK7HcpFzK82QcQ23hJp/eeJBK5n3s5tTBou68Td4+pF
# K4ihvY9ZoXFBTOuX2irjEwJ60E+F5Ofgfg9RptDsdtlF5c2VlWlPrr/hhyLJMDXB
# HXxJMcuiHH1fj3Hawu8vUl5acN3Q8AmIwaipCl8o68o65c/AUtEIpqLzw9iBWZnP
# XTl2trGXkY+uRtvLGLhLKP528fGqP6BzFreUgyVQX+1Wq0IzeyZS0nLJVNVTVN6i
# LM4KvrC4hrQEB74fFW8+WWbvm4EG7Xx2/D0ZdiqUECCgAx0TsxBXqjAQ67mpc28K
# 5LdWKL3HNkSwiIa0Hnqlqg5I6VLAnlhcppF1OoB4Hm7cgy7xkpzmrFhDX8wf3jz2
# +9vaw0SjEI4+iO1LU0xNTbZtuKwp7W7UgpM6XJa5nTHcHFs7ldOR1/5k4GDP9rS0
# 3FtlNxvalg==
# SIG # End signature block
