<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB Intune managed-device pipeline tests.

.VERSION
0.1.0
#>
[CmdletBinding()]
param()

$ScriptVersion = '0.1.1'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:Passed = 0
$script:Failed = 0

function Invoke-SmartWorkplaceCMDBIntuneDeviceTest {
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

function Assert-SmartWorkplaceCMDBIntuneDeviceTrue {
    [CmdletBinding()]
    param([Parameter(Mandatory)][bool]$Condition, [Parameter(Mandatory)][string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBIntuneDeviceThrow {
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Action, [Parameter(Mandatory)][string]$MessagePattern)
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
$entraCollector = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraDevices-Collect.ps1'
$collector = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneManagedDevices-Collect.ps1'
$normalizer = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneDevices-Normalize.ps1'
$coreModule = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContract = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContract = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$entraFixture = Join-Path $PSScriptRoot 'Fixtures\EntraDevices.sample.json'
$intuneFixture = Join-Path $PSScriptRoot 'Fixtures\IntuneManagedDevices.sample.json'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$tempRoot = Join-Path $tempBase ('SmartWorkplaceCMDB-IntuneDevices-Tests-{0}' -f [guid]::NewGuid().ToString('N'))
Import-Module $coreModule -Force

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $identity = @{
        Tenant = 'audit'; OrganizationKey = 'contoso'; EnvironmentKey = 'prod'
        TenantKey = 'contoso-prod'; TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        NoConfigWrite = $true
    }

    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Collector ValidateOnly stays read-only' {
        $root = Join-Path $tempRoot 'ValidateCollector'
        & $collector @identity -DataRootPath $root -InputJsonPath $intuneFixture -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (-not (Test-Path $root)) 'Collector ValidateOnly created output.'
    }

    $runtime = Join-Path $tempRoot 'Runtime'
    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Collect Entra and Intune fixture devices' {
        $script:EntraCollection = & $entraCollector @identity -DataRootPath $runtime -InputJsonPath $entraFixture
        $script:IntuneCollection = & $collector @identity -DataRootPath $runtime -InputJsonPath $intuneFixture
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue ($script:EntraCollection.DeviceCount -eq 2) 'Expected two Entra devices.'
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue ($script:IntuneCollection.DeviceCount -eq 3) 'Expected three Intune devices.'
    }

    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Validate Intune raw contract and MaxItems' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtime 'DATA-LAST') `
            -ContractPath $rawContract)
        $intune = @($results | Where-Object Name -eq 'Intune_ManagedDevices.csv')
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $intune.Count -eq 1 -and $intune[0].Status -eq 'Valid'
        ) 'Raw Intune contract is invalid.'
        $bounded = & $collector @identity `
            -DataRootPath (Join-Path $tempRoot 'Bounded') `
            -InputJsonPath $intuneFixture `
            -MaxItems 1
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            @(Import-Csv $bounded.RawLatestOutputPath).Count -eq 1
        ) 'MaxItems was not honored.'
    }

    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Enrich Entra devices and retain Intune-only devices' {
        $script:Normalization = & $normalizer @identity -DataRootPath $runtime
        $cmdb = @(Import-Csv $script:Normalization.CmdbDeviceOutputPath)
        $dim = @(Import-Csv $script:Normalization.DimDeviceOutputPath)
        $compliance = @(Import-Csv $script:Normalization.FactDeviceComplianceOutputPath)
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $cmdb.Count -eq 4 -and $dim.Count -eq 4 -and
            $compliance.Count -eq 4
        ) 'Expected four unioned device records and compliance facts.'
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $script:Normalization.EnrichedDeviceCount -eq 1 -and
            $script:Normalization.IntuneOnlyDeviceCount -eq 2
        ) 'Enrichment counters are invalid.'
        $matched = @($cmdb | Where-Object SourceDeviceId -eq '66666666-6666-6666-6666-666666666666')
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $matched.Count -eq 1 -and
            $matched[0].Ownership -eq 'Corporate' -and
            $matched[0].ComplianceState -eq 'NonCompliant' -and
            $matched[0].PrimaryUserId -eq '11111111-1111-1111-1111-111111111111'
        ) 'Matched Intune enrichment is invalid.'
    }

    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Publish one compliance fact per device' {
        $cmdb = @(Import-Csv $script:Normalization.CmdbDeviceOutputPath)
        $facts = @(Import-Csv $script:Normalization.FactDeviceComplianceOutputPath)
        $matched = @($facts | Where-Object {
                $_.CmdbDeviceId -eq 'contoso-prod|device|66666666-6666-6666-6666-666666666666'
            })
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $facts.Count -eq $cmdb.Count -and
            $script:Normalization.FactDeviceComplianceCount -eq $cmdb.Count -and
            @($facts | Group-Object TenantDeviceKey | Where-Object Count -gt 1).Count -eq 0
        ) 'Compliance fact count or keys are invalid.'
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $matched.Count -eq 1 -and
            $matched[0].ComplianceState -eq 'NonCompliant' -and
            $matched[0].LastSyncDateTime -eq '2026-07-19T08:30:00.0000000Z' -and
            $matched[0].SourceSystem -eq 'MicrosoftEntraID+MicrosoftIntune'
        ) 'Compliance fact mapping is invalid.'
    }

    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Use stable fallback for missing Azure AD device ID' {
        $cmdb = @(Import-Csv $script:Normalization.CmdbDeviceOutputPath)
        $fallback = @($cmdb | Where-Object SourceDeviceId -like 'intune:*')
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $fallback.Count -eq 1 -and
            $fallback[0].CmdbDeviceId -like 'contoso-prod|device|intune:*'
        ) 'Missing Azure AD device ID fallback is invalid.'
    }

    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Select the freshest duplicate correlation deterministically' {
        $rows = @(Import-Csv $script:IntuneCollection.RawLatestOutputPath)
        $newer = $rows[0].PSObject.Copy()
        $newer.ManagedDeviceId = '99999999-9999-9999-9999-999999999994'
        $newer.DeviceName = 'EXAMPLE-WIN-LATEST'
        $newer.LastSyncDateTime = '2026-07-20T08:30:00Z'
        $duplicateCorrelationPath = Join-Path $tempRoot 'DuplicateCorrelation.csv'
        @($rows + $newer) | Export-Csv $duplicateCorrelationPath -NoTypeInformation -Encoding UTF8
        $result = & $normalizer @identity `
            -DataRootPath (Join-Path $tempRoot 'DuplicateCorrelation') `
            -EntraRawInputPath $script:EntraCollection.RawLatestOutputPath `
            -IntuneRawInputPath $duplicateCorrelationPath
        $cmdb = @(Import-Csv $result.CmdbDeviceOutputPath)
        $matched = @($cmdb | Where-Object SourceDeviceId -eq '66666666-6666-6666-6666-666666666666')
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $result.DuplicateCorrelationCount -eq 1 -and
            $matched.Count -eq 1 -and
            $matched[0].DeviceName -eq 'EXAMPLE-WIN-LATEST'
        ) 'Freshest duplicate correlation was not selected.'
    }

    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Validate curated device contracts and keys' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtime 'DATA-LAST') `
            -ContractPath $curatedContract)
        $devices = @($results | Where-Object Name -in @('CMDB_Devices.csv', 'DimDevice.csv', 'FactDeviceCompliance.csv'))
        $cmdb = @(Import-Csv $script:Normalization.CmdbDeviceOutputPath)
        $dim = @(Import-Csv $script:Normalization.DimDeviceOutputPath)
        $facts = @(Import-Csv $script:Normalization.FactDeviceComplianceOutputPath)
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $devices.Count -eq 3 -and
            @($devices | Where-Object Status -ne 'Valid').Count -eq 0
        ) 'Curated device contracts are invalid.'
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            @($cmdb | Group-Object CmdbDeviceId | Where-Object Count -gt 1).Count -eq 0 -and
            @($dim | Group-Object TenantDeviceKey | Where-Object Count -gt 1).Count -eq 0 -and
            @($facts | Group-Object TenantDeviceKey | Where-Object Count -gt 1).Count -eq 0
        ) 'Curated device keys are not unique.'
    }

    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Reject Intune identity mismatch and duplicate managed device IDs' {
        $rows = @(Import-Csv $script:IntuneCollection.RawLatestOutputPath)
        $mismatch = Join-Path $tempRoot 'Mismatch.csv'
        $rows[0].TenantKey = 'other-prod'
        $rows | Export-Csv $mismatch -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBIntuneDeviceThrow {
            & $normalizer @identity `
                -DataRootPath (Join-Path $tempRoot 'Mismatch') `
                -EntraRawInputPath $script:EntraCollection.RawLatestOutputPath `
                -IntuneRawInputPath $mismatch | Out-Null
        } 'identity mismatch'

        $rows = @(Import-Csv $script:IntuneCollection.RawLatestOutputPath)
        $duplicate = Join-Path $tempRoot 'Duplicate.csv'
        $rows[1].ManagedDeviceId = $rows[0].ManagedDeviceId
        $rows | Export-Csv $duplicate -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBIntuneDeviceThrow {
            & $normalizer @identity `
                -DataRootPath (Join-Path $tempRoot 'Duplicate') `
                -EntraRawInputPath $script:EntraCollection.RawLatestOutputPath `
                -IntuneRawInputPath $duplicate | Out-Null
        } 'duplicate ManagedDeviceId'
    }

    Invoke-SmartWorkplaceCMDBIntuneDeviceTest 'Normalizer ValidateOnly preserves outputs' {
        $beforeCmdb = (Get-FileHash $script:Normalization.CmdbDeviceOutputPath -Algorithm SHA256).Hash
        $beforeDim = (Get-FileHash $script:Normalization.DimDeviceOutputPath -Algorithm SHA256).Hash
        $beforeCompliance = (Get-FileHash $script:Normalization.FactDeviceComplianceOutputPath -Algorithm SHA256).Hash
        & $normalizer @identity -DataRootPath $runtime -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBIntuneDeviceTrue (
            $beforeCmdb -eq (Get-FileHash $script:Normalization.CmdbDeviceOutputPath -Algorithm SHA256).Hash -and
            $beforeDim -eq (Get-FileHash $script:Normalization.DimDeviceOutputPath -Algorithm SHA256).Hash -and
            $beforeCompliance -eq (Get-FileHash $script:Normalization.FactDeviceComplianceOutputPath -Algorithm SHA256).Hash
        ) 'ValidateOnly modified curated device outputs.'
    }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolved) -like 'SmartWorkplaceCMDB-IntuneDevices-Tests-*' -and
        (Test-Path $resolved)) {
        Remove-Item $resolved -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB Intune devices tests completed. Version={0}; Passed={1}; Failed={2}" -f
    $ScriptVersion, $script:Passed, $script:Failed
) -InformationAction Continue
if ($script:Failed -gt 0) {
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBoGLpC/hsGpzdA
# kYE6RhwsVQeAoyeupilcF3F3PnmzbaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIhpymXEYPY3vZxngMWy9f1sMeBjL83Wm5XpDSYvMBIKMA0GCSqG
# SIb3DQEBAQUABIIBgGssIQxAFm90vufBud1c990bwS6M1ZdOJ1fi9x1UhoFzyTAp
# YeUAPc4Wjw3fUYw/lOZNnNuy7N7WrMnbsin7LHs5+Ek/rAXZsRVJye9x+F8L0e6d
# 6hcuSlhX279dZl9IaKTKmV3nRltuX4h8dTA+FsVSXRwIfNsmpVYBmydKkKbmxcaD
# X/JdHLI4ahGBjAtvvt3e1tRrb273DkAqb8OcHSD8UKBtIJCGtipMo07lbsEGrHdO
# GvZWva1Uz9EYx+yxF0ZiDERrkz8aVI6nuC9zta2l+uZZu1MU/kmxrbO6ybT1R8gC
# hfQvh1MhFjqEyREuJUVzKZabgjefMe5hkXLfPGQa2dokNlbmdfUHdKAaxTTsMKF2
# JniXPLc9GPOmi3eavIW7Z+Cf0XkkLWmntz0OJFgo7tyEmApjKa2y2DcRJG5YD39Y
# b85UamO2O+wV9kE1cAInj+PNEbUJyRBDQM54lQ2QB5veqbNKepk24428Luj7pp23
# 6YEfRa7TkzaG5ngAZ6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxODEx
# MTJaMC8GCSqGSIb3DQEJBDEiBCAzhbv1C7q6UeZX7Qw6wRd6tyCXOk8v8yFISWzS
# 9y1jHTANBgkqhkiG9w0BAQEFAASCAgDORJBX69V/Y0vCSxKyBk7qhK9x0KD56IEK
# GPfWGohmVueYokkfU/Sh/QaEKIVQfqLc7C1ZGGTFmk6yg05ds7hrD37nX7PebZrK
# OWTNBjiqgi69XvCUFCfhv2hfZsV5mfq65hnb7TYoauQxsDsVFFIzErbphm6fGlNT
# 18U6UVgz5YvB8R8jpeL24WTRzi4eXV7c8zCsPN87vkna8zLfB5QPtmg+8Z4trHD7
# nGPCpYndkiC512wRDhvzBfUl5g+bBtpmoWYnSLEw2PMhrIHfV468u2Mc/2JPfa2O
# iotFitYylRWcbBBH1gavwpE3j9K5wjth04PGGNfncRxcA34QvG04yLKft2w1Qpp+
# PJYmWFgnz82/zv+eNgyS/d/TOn0Y0QVFcYUDdCF3y+tThtF0MAjxWOylzob/Ah96
# 2YZmYYBz1mmhUUhL4nSxNG3cO8wt6MXnAWA5reIGHdppYxisqOPh3Xa3Hf7bkY63
# R0EeJIuV9vrb+LtVMpP2GlrNlOv3EfQzbwPU0fhsCchmBqZ0UcM06WM5DHUDOtYp
# 7TbIcq5sUXsVjD5nHgh3+sr6L6yxLNhIjuIsu5cZKE4CokkubhrB3qXKG9DAK2xm
# ExJOiOGm/9XpR2AzjsK0DPahNSXvkT1OSCQimhvk3JULFSVvGN5vrwk7+1IbETIz
# XkZbu7fIvw==
# SIG # End signature block
