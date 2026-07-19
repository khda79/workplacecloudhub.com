<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB primary user-device relationship tests.

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

function Invoke-SmartWorkplaceCMDBUserDeviceTest {
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

function Assert-SmartWorkplaceCMDBUserDeviceTrue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBUserDeviceThrow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [Parameter(Mandatory)][string]$MessagePattern
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
$userCollector = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Collect.ps1'
$userNormalizer = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Normalize.ps1'
$entraDeviceCollector = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraDevices-Collect.ps1'
$intuneCollector = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneManagedDevices-Collect.ps1'
$deviceNormalizer = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneDevices-Normalize.ps1'
$relationshipNormalizer = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneUserDeviceRelationships-Normalize.ps1'
$coreModule = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$contractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$userFixture = Join-Path $PSScriptRoot 'Fixtures\EntraUsers.sample.json'
$entraDeviceFixture = Join-Path $PSScriptRoot 'Fixtures\EntraDevices.sample.json'
$intuneFixture = Join-Path $PSScriptRoot 'Fixtures\IntuneManagedDevices.sample.json'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$tempRoot = Join-Path $tempBase (
    'SmartWorkplaceCMDB-UserDevice-Tests-{0}' -f
    [guid]::NewGuid().ToString('N')
)
Import-Module $coreModule -Force

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $identity = @{
        Tenant = 'audit'
        OrganizationKey = 'contoso'
        EnvironmentKey = 'prod'
        TenantKey = 'contoso-prod'
        TenantId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        NoConfigWrite = $true
    }

    $runtime = Join-Path $tempRoot 'Runtime'
    Invoke-SmartWorkplaceCMDBUserDeviceTest 'Prepare curated users and devices' {
        & $userCollector @identity `
            -DataRootPath $runtime `
            -InputJsonPath $userFixture | Out-Null
        & $userNormalizer @identity -DataRootPath $runtime | Out-Null
        & $entraDeviceCollector @identity `
            -DataRootPath $runtime `
            -InputJsonPath $entraDeviceFixture | Out-Null
        & $intuneCollector @identity `
            -DataRootPath $runtime `
            -InputJsonPath $intuneFixture | Out-Null
        $script:DeviceNormalization = & $deviceNormalizer @identity `
            -DataRootPath $runtime
        Assert-SmartWorkplaceCMDBUserDeviceTrue `
            ($script:DeviceNormalization.DeviceCount -eq 4) `
            'Expected four curated devices.'
    }

    Invoke-SmartWorkplaceCMDBUserDeviceTest 'ValidateOnly stays read-only' {
        $root = Join-Path $tempRoot 'Validate'
        & $relationshipNormalizer @identity `
            -DataRootPath $root `
            -UserInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Users.csv') `
            -DeviceInputPath $script:DeviceNormalization.CmdbDeviceOutputPath `
            -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBUserDeviceTrue `
            (-not (Test-Path $root)) `
            'ValidateOnly created relationship output.'
    }

    Invoke-SmartWorkplaceCMDBUserDeviceTest 'Publish primary user relationships' {
        $script:Normalization = & $relationshipNormalizer @identity `
            -DataRootPath $runtime
        $relationships = @(Import-Csv `
            $script:Normalization.RelationshipOutputPath)
        $facts = @(Import-Csv $script:Normalization.FactOutputPath)
        Assert-SmartWorkplaceCMDBUserDeviceTrue `
            ($script:Normalization.CandidatePrimaryUsers -eq 2 -and
                $script:Normalization.RelationshipCount -eq 2 -and
                $script:Normalization.OrphanPrimaryUserCount -eq 0 -and
                $relationships.Count -eq 2 -and
                $facts.Count -eq 2) `
            'Primary user relationship counts are invalid.'
    }

    Invoke-SmartWorkplaceCMDBUserDeviceTest 'Validate mapping and stable keys' {
        $relationships = @(Import-Csv `
            $script:Normalization.RelationshipOutputPath)
        $facts = @(Import-Csv $script:Normalization.FactOutputPath)
        $primary = @($relationships | Where-Object {
                $_.CmdbUserId -eq
                'contoso-prod|entra-user|11111111-1111-1111-1111-111111111111'
            })
        Assert-SmartWorkplaceCMDBUserDeviceTrue `
            ($primary.Count -eq 1 -and
                $primary[0].RelationshipType -eq 'PrimaryUser' -and
                $primary[0].SourceSystem -eq 'MicrosoftIntune' -and
                $primary[0].Evidence -eq 'managedDevices.userId') `
            'Primary user relationship mapping is invalid.'
        Assert-SmartWorkplaceCMDBUserDeviceTrue `
            (@($relationships |
                    Group-Object CmdbRelationshipId |
                    Where-Object Count -gt 1).Count -eq 0 -and
                @($facts |
                    Group-Object TenantRelationshipKey |
                    Where-Object Count -gt 1).Count -eq 0 -and
                $facts[0].TenantRelationshipKey -eq
                $relationships[0].CmdbRelationshipId) `
            'Relationship keys are invalid.'
    }

    Invoke-SmartWorkplaceCMDBUserDeviceTest 'Validate curated contracts' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtime 'DATA-LAST') `
            -ContractPath $contractPath)
        $tables = @($results | Where-Object Name -in @(
                'CMDB_UserDeviceRelationships.csv',
                'FactUserDeviceRelationship.csv'
            ))
        Assert-SmartWorkplaceCMDBUserDeviceTrue `
            ($tables.Count -eq 2 -and
                @($tables | Where-Object Status -ne 'Valid').Count -eq 0) `
            'User-device relationship contracts are invalid.'
    }

    Invoke-SmartWorkplaceCMDBUserDeviceTest 'Count and skip orphan primary users' {
        $devices = @(Import-Csv `
            $script:DeviceNormalization.CmdbDeviceOutputPath)
        $candidate = @($devices | Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.PrimaryUserId)
            })[0]
        $candidate.PrimaryUserId = 'ffffffff-ffff-ffff-ffff-ffffffffffff'
        $orphanPath = Join-Path $tempRoot 'OrphanDevices.csv'
        $devices | Export-Csv $orphanPath -NoTypeInformation -Encoding UTF8
        $result = & $relationshipNormalizer @identity `
            -DataRootPath (Join-Path $tempRoot 'Orphan') `
            -UserInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Users.csv') `
            -DeviceInputPath $orphanPath
        Assert-SmartWorkplaceCMDBUserDeviceTrue `
            ($result.CandidatePrimaryUsers -eq 2 -and
                $result.RelationshipCount -eq 1 -and
                $result.OrphanPrimaryUserCount -eq 1) `
            'Orphan primary user handling is invalid.'
    }

    Invoke-SmartWorkplaceCMDBUserDeviceTest 'Reject identity mismatch' {
        $devices = @(Import-Csv `
            $script:DeviceNormalization.CmdbDeviceOutputPath)
        $devices[0].TenantKey = 'other-prod'
        $mismatchPath = Join-Path $tempRoot 'Mismatch.csv'
        $devices | Export-Csv $mismatchPath -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBUserDeviceThrow {
            & $relationshipNormalizer @identity `
                -DataRootPath (Join-Path $tempRoot 'Mismatch') `
                -UserInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Users.csv') `
                -DeviceInputPath $mismatchPath | Out-Null
        } 'identity mismatch'
    }

    Invoke-SmartWorkplaceCMDBUserDeviceTest 'Publish schema-only empty outputs' {
        $devices = @(Import-Csv `
            $script:DeviceNormalization.CmdbDeviceOutputPath)
        foreach ($device in $devices) {
            $device.PrimaryUserId = ''
        }
        $emptyPath = Join-Path $tempRoot 'EmptyDevices.csv'
        $devices | Export-Csv $emptyPath -NoTypeInformation -Encoding UTF8
        $emptyRoot = Join-Path $tempRoot 'Empty'
        $result = & $relationshipNormalizer @identity `
            -DataRootPath $emptyRoot `
            -UserInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Users.csv') `
            -DeviceInputPath $emptyPath
        $contracts = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $emptyRoot 'DATA-LAST') `
            -ContractPath $contractPath |
            Where-Object Name -in @(
                'CMDB_UserDeviceRelationships.csv',
                'FactUserDeviceRelationship.csv'
            ))
        Assert-SmartWorkplaceCMDBUserDeviceTrue `
            ($result.RelationshipCount -eq 0 -and
                @($contracts | Where-Object Status -ne 'Valid').Count -eq 0) `
            'Empty relationship outputs are invalid.'
    }

    Invoke-SmartWorkplaceCMDBUserDeviceTest 'ValidateOnly preserves outputs' {
        $beforeRelationship = (Get-FileHash `
            $script:Normalization.RelationshipOutputPath `
            -Algorithm SHA256).Hash
        $beforeFact = (Get-FileHash `
            $script:Normalization.FactOutputPath `
            -Algorithm SHA256).Hash
        & $relationshipNormalizer @identity `
            -DataRootPath $runtime `
            -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBUserDeviceTrue `
            ($beforeRelationship -eq (Get-FileHash `
                    $script:Normalization.RelationshipOutputPath `
                    -Algorithm SHA256).Hash -and
                $beforeFact -eq (Get-FileHash `
                    $script:Normalization.FactOutputPath `
                    -Algorithm SHA256).Hash) `
            'ValidateOnly modified relationship outputs.'
    }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    if ($resolved.StartsWith(
            $tempBase,
            [StringComparison]::OrdinalIgnoreCase
        ) -and
        (Split-Path -Leaf $resolved) -like
        'SmartWorkplaceCMDB-UserDevice-Tests-*' -and
        (Test-Path $resolved)) {
        Remove-Item $resolved -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB user-device relationship tests completed. Version={0}; Passed={1}; Failed={2}" -f
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBR1iQVws4GRzH2
# GIWvhYjupymQxGKF++Hm16ocue4ckqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBvfPCdkTF9Qp3UUFjyV+pLNND1Gy9LNiAFGdAflkUEaMA0GCSqG
# SIb3DQEBAQUABIIBgBxNDu0N2qnjy22S9JR4xtPzINpnoGvLSRHTSGNcSdluw75I
# 9nn6RfLpYkS2d2DJ6l0jF6/DyNiUwmj8/6Fl6yKjvVftcuHf8f1OBMZAGuw/Z4qV
# 18tfp49yWwMn37xxAQYtehR5NXbHawZskj4QeixFyK9rqV1jef+eB3huRDF+PpDv
# hkxGOVpQU8KgtsbK7LPZu/hr8x65QkjP3NytRzfdEjGYnr3eceEmyxhdto6ZRjj8
# 4BhjErQAgD0FDSh8EDC1LLeAge5kNdyJ4wVASEVT+hDLutrbmqiU7d3G3sLSIPEf
# 9u5k3nVch92hqwaCzuqmi5yt3csvulBtD0rKXfxkzaHMSuJuzV1rcW16k8VNKV38
# G6ucJ/Tv8TY1ktELTdXb4EdxMzTD35CW0GE0GKHZJ2ynXnPMZA0AjcX1DOxDyxQq
# 9AhSFA0qxurrpy5JUQMQ8qFw8O2/4XKCiSoDPXWOPzsEtQFC1HhTsAvgbqTx/gb4
# p+r2uw2iOPd++RKxEqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxOTQy
# MzRaMC8GCSqGSIb3DQEJBDEiBCC6aYWTuswJ2Rm4dj7UrRCSQ/ezpHhpjVh3LVg4
# xkrdzzANBgkqhkiG9w0BAQEFAASCAgBpqfRUFc7GKjQqpasgVwIhcOx2TPZ4IKLk
# 2Xgi8clTk/A/wSvGkCXY9LqCBS0qdBufN3Q66k+trDze4GDrEOUcxB28pcanky2+
# Ub5l2NcgzjkI5+49EX1ZdJE/nyPosP2swidsJ7griH299aoBflRCDyDhyy0FWOYf
# K8WtMZm1bfaMp4/V/TaWmgsVglCBb7hrWy2KvPsx7rPcXwU9PHgGUGYVUdDKX32t
# h2rraKjWw++YM8tw1liK4P5995A/pd3l81q9Xb6MdUy/+uqcx+ZxdBTeBmwrKqWH
# E4bw7ovG85EHaKHGnjogaErGlsWg2fuY5axGeqdVKEyxOzGs9Nh9MaufmB6ljISP
# D7GFyJj3BTUwSYTTAqyOf4lXxdOcUVCp5g55fINbQvsM0PzKAKsBieY+phbUkUtJ
# VTUJyxgZpmxZMaQ7Xkmxlz89BKtCu1EXjdrstUPt9kL5NDaVrr8yQn/jbhCNZ8Hv
# lxlNqxBVPHSUA1McUdtRE1SmbX4E5dnLriahHfFx7J5qByqPXs633s18CxCChDmH
# jsovoS6K3cPCy5P7md0G7OkHTELeZSnkeQAzaG7Ugm0qHfrovcZguinnlGH7R6jJ
# DASq9FW/hGSLUgS+8KVNvZl/ZBaX/el+A962muDDbSeO/Px1gvIOqUar71rp8dzJ
# t/Aof4+ffQ==
# SIG # End signature block
