<#
.SYNOPSIS
Runs offline SmartWorkplaceCMDB general relationship consolidation tests.

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

function Invoke-SmartWorkplaceCMDBRelationshipTest {
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

function Assert-SmartWorkplaceCMDBRelationshipTrue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-SmartWorkplaceCMDBRelationshipThrow {
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

function Write-SmartWorkplaceCMDBHeaderOnlyFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$TargetPath
    )
    $header = Get-Content -LiteralPath $SourcePath -TotalCount 1
    Set-Content -LiteralPath $TargetPath -Value $header -Encoding UTF8
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$userCollector = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Collect.ps1'
$userNormalizer = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Normalize.ps1'
$deviceCollector = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraDevices-Collect.ps1'
$intuneCollector = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneManagedDevices-Collect.ps1'
$deviceNormalizer = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneDevices-Normalize.ps1'
$userDeviceNormalizer = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneUserDeviceRelationships-Normalize.ps1'
$skuCollector = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Collect.ps1'
$skuNormalizer = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Normalize.ps1'
$licenseCollector = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Collect.ps1'
$licenseNormalizer = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Normalize.ps1'
$mailboxCollector = Join-Path $projectRoot 'Collectors\ExchangeOnline\SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Collect.ps1'
$mailboxNormalizer = Join-Path $projectRoot 'Collectors\ExchangeOnline\SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Normalize.ps1'
$normalizer = Join-Path $projectRoot 'Collectors\SmartWorkplaceCMDB-Relationships-Normalize.ps1'
$coreModule = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$contractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$fixtures = @{
    Users = Join-Path $PSScriptRoot 'Fixtures\EntraUsers.sample.json'
    Devices = Join-Path $PSScriptRoot 'Fixtures\EntraDevices.sample.json'
    Intune = Join-Path $PSScriptRoot 'Fixtures\IntuneManagedDevices.sample.json'
    Skus = Join-Path $PSScriptRoot 'Fixtures\M365SubscribedSkus.sample.json'
    Licenses = Join-Path $PSScriptRoot 'Fixtures\M365UserLicenseAssignments.sample.json'
    Mailboxes = Join-Path $PSScriptRoot 'Fixtures\ExchangeOnlineMailboxes.sample.json'
}
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$tempRoot = Join-Path $tempBase (
    'SmartWorkplaceCMDB-Relationships-Tests-{0}' -f
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

    Invoke-SmartWorkplaceCMDBRelationshipTest 'Prepare all relationship sources' {
        & $userCollector @identity -DataRootPath $runtime `
            -InputJsonPath $fixtures.Users | Out-Null
        & $userNormalizer @identity -DataRootPath $runtime | Out-Null
        & $deviceCollector @identity -DataRootPath $runtime `
            -InputJsonPath $fixtures.Devices | Out-Null
        & $intuneCollector @identity -DataRootPath $runtime `
            -InputJsonPath $fixtures.Intune | Out-Null
        & $deviceNormalizer @identity -DataRootPath $runtime | Out-Null
        & $userDeviceNormalizer @identity -DataRootPath $runtime | Out-Null
        & $skuCollector @identity -DataRootPath $runtime `
            -InputJsonPath $fixtures.Skus | Out-Null
        & $skuNormalizer @identity -DataRootPath $runtime | Out-Null
        & $licenseCollector @identity -DataRootPath $runtime `
            -InputJsonPath $fixtures.Licenses | Out-Null
        & $licenseNormalizer @identity -DataRootPath $runtime | Out-Null
        & $mailboxCollector @identity -DataRootPath $runtime `
            -InputJsonPath $fixtures.Mailboxes | Out-Null
        & $mailboxNormalizer @identity -DataRootPath $runtime | Out-Null
        Assert-SmartWorkplaceCMDBRelationshipTrue `
            (Test-Path (Join-Path $runtime 'DATA-LAST\PowerBI\FactMailbox.csv')) `
            'Relationship source preparation failed.'
    }

    Invoke-SmartWorkplaceCMDBRelationshipTest 'ValidateOnly stays read-only' {
        $root = Join-Path $tempRoot 'Validate'
        & $normalizer @identity `
            -DataRootPath $root `
            -UserInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Users.csv') `
            -DeviceInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Devices.csv') `
            -LicenseInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Licenses.csv') `
            -MailboxInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Mailboxes.csv') `
            -UserDeviceInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_UserDeviceRelationships.csv') `
            -MailboxFactInputPath (Join-Path $runtime 'DATA-LAST\PowerBI\FactMailbox.csv') `
            -UserLicenseFactInputPath (Join-Path $runtime 'DATA-LAST\PowerBI\FactUserLicense.csv') `
            -RawUserLicenseInputPath (Join-Path $runtime 'DATA-LAST\Raw\M365\M365_UserLicenseAssignments.csv') `
            -ValidateOnly | Out-Null
        Assert-SmartWorkplaceCMDBRelationshipTrue `
            (-not (Test-Path $root)) `
            'ValidateOnly created relationship output.'
    }

    Invoke-SmartWorkplaceCMDBRelationshipTest 'Consolidate three relationship types' {
        $script:Normalization = & $normalizer @identity -DataRootPath $runtime
        $rows = @(Import-Csv $script:Normalization.OutputPath)
        Assert-SmartWorkplaceCMDBRelationshipTrue `
            ($script:Normalization.UserDeviceCount -eq 2 -and
                $script:Normalization.MailboxCount -eq 2 -and
                $script:Normalization.UnlinkedMailboxCount -eq 1 -and
                $script:Normalization.UserLicenseCount -eq 2 -and
                $script:Normalization.RelationshipCount -eq 6 -and
                $rows.Count -eq 6) `
            'Relationship type counts are invalid.'
    }

    Invoke-SmartWorkplaceCMDBRelationshipTest 'Validate relationship mappings' {
        $rows = @(Import-Csv $script:Normalization.OutputPath)
        Assert-SmartWorkplaceCMDBRelationshipTrue `
            (@($rows | Where-Object RelationshipType -eq 'PrimaryUser').Count -eq 2 -and
                @($rows | Where-Object RelationshipType -eq 'HasMailbox').Count -eq 2 -and
                @($rows | Where-Object RelationshipType -eq 'AssignedLicense').Count -eq 2) `
            'Relationship type mapping is invalid.'
        Assert-SmartWorkplaceCMDBRelationshipTrue `
            (@($rows | Where-Object {
                    [string]::IsNullOrWhiteSpace($_.SourceCollectedDateTime)
                }).Count -eq 0 -and
                @($rows |
                    Group-Object CmdbRelationshipId |
                    Where-Object Count -gt 1).Count -eq 0) `
            'Relationship dates or stable keys are invalid.'
    }

    Invoke-SmartWorkplaceCMDBRelationshipTest 'Validate curated relationship contract' {
        $results = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $runtime 'DATA-LAST') `
            -ContractPath $contractPath)
        $relationship = @($results |
            Where-Object Name -eq 'CMDB_Relationships.csv')
        Assert-SmartWorkplaceCMDBRelationshipTrue `
            ($relationship.Count -eq 1 -and
                $relationship[0].Status -eq 'Valid') `
            'General relationship contract is invalid.'
    }

    Invoke-SmartWorkplaceCMDBRelationshipTest 'Reject broken user references' {
        $path = Join-Path $runtime 'DATA-LAST\CMDB\CMDB_UserDeviceRelationships.csv'
        $rows = @(Import-Csv $path)
        $rows[0].CmdbUserId = 'contoso-prod|entra-user|ffffffff-ffff-ffff-ffff-ffffffffffff'
        $brokenPath = Join-Path $tempRoot 'BrokenUserDevice.csv'
        $rows | Export-Csv $brokenPath -NoTypeInformation -Encoding UTF8
        Assert-SmartWorkplaceCMDBRelationshipThrow {
            & $normalizer @identity `
                -DataRootPath (Join-Path $tempRoot 'Broken') `
                -UserInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Users.csv') `
                -DeviceInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Devices.csv') `
                -LicenseInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Licenses.csv') `
                -MailboxInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Mailboxes.csv') `
                -UserDeviceInputPath $brokenPath `
                -MailboxFactInputPath (Join-Path $runtime 'DATA-LAST\PowerBI\FactMailbox.csv') `
                -UserLicenseFactInputPath (Join-Path $runtime 'DATA-LAST\PowerBI\FactUserLicense.csv') `
                -RawUserLicenseInputPath (Join-Path $runtime 'DATA-LAST\Raw\M365\M365_UserLicenseAssignments.csv') | Out-Null
        } 'unknown CMDB user'
    }

    Invoke-SmartWorkplaceCMDBRelationshipTest 'Publish schema-only empty output' {
        $userDevice = Join-Path $runtime 'DATA-LAST\CMDB\CMDB_UserDeviceRelationships.csv'
        $mailboxFact = Join-Path $runtime 'DATA-LAST\PowerBI\FactMailbox.csv'
        $licenseFact = Join-Path $runtime 'DATA-LAST\PowerBI\FactUserLicense.csv'
        $emptyUserDevice = Join-Path $tempRoot 'EmptyUserDevice.csv'
        $emptyMailbox = Join-Path $tempRoot 'EmptyMailbox.csv'
        $emptyLicense = Join-Path $tempRoot 'EmptyLicense.csv'
        Write-SmartWorkplaceCMDBHeaderOnlyFile $userDevice $emptyUserDevice
        Write-SmartWorkplaceCMDBHeaderOnlyFile $mailboxFact $emptyMailbox
        Write-SmartWorkplaceCMDBHeaderOnlyFile $licenseFact $emptyLicense
        $emptyRoot = Join-Path $tempRoot 'Empty'
        $result = & $normalizer @identity `
            -DataRootPath $emptyRoot `
            -UserInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Users.csv') `
            -DeviceInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Devices.csv') `
            -LicenseInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Licenses.csv') `
            -MailboxInputPath (Join-Path $runtime 'DATA-LAST\CMDB\CMDB_Mailboxes.csv') `
            -UserDeviceInputPath $emptyUserDevice `
            -MailboxFactInputPath $emptyMailbox `
            -UserLicenseFactInputPath $emptyLicense `
            -RawUserLicenseInputPath (Join-Path $runtime 'DATA-LAST\Raw\M365\M365_UserLicenseAssignments.csv')
        $contracts = @(Test-SmartWorkplaceCMDBCsvContract `
            -LatestOutputRootPath (Join-Path $emptyRoot 'DATA-LAST') `
            -ContractPath $contractPath |
            Where-Object Name -eq 'CMDB_Relationships.csv')
        Assert-SmartWorkplaceCMDBRelationshipTrue `
            ($result.RelationshipCount -eq 0 -and
                $contracts.Count -eq 1 -and
                $contracts[0].Status -eq 'Valid') `
            'Empty general relationship output is invalid.'
    }

    Invoke-SmartWorkplaceCMDBRelationshipTest 'ValidateOnly preserves output' {
        $before = (Get-FileHash `
            $script:Normalization.OutputPath `
            -Algorithm SHA256).Hash
        & $normalizer @identity -DataRootPath $runtime -ValidateOnly | Out-Null
        $after = (Get-FileHash `
            $script:Normalization.OutputPath `
            -Algorithm SHA256).Hash
        Assert-SmartWorkplaceCMDBRelationshipTrue `
            ($before -eq $after) `
            'ValidateOnly modified the general relationship output.'
    }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    if ($resolved.StartsWith(
            $tempBase,
            [StringComparison]::OrdinalIgnoreCase
        ) -and
        (Split-Path -Leaf $resolved) -like
        'SmartWorkplaceCMDB-Relationships-Tests-*' -and
        (Test-Path $resolved)) {
        Remove-Item $resolved -Recurse -Force
    }
}

Write-Information (
    "SmartWorkplaceCMDB relationship tests completed. Version={0}; Passed={1}; Failed={2}" -f
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCUXSSBEAgfLGXd
# 3N6WnM0vVIX2Kc+J7Y2wmxiap+QruaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHham0rvykYriKpan4eaTmG2x4LsvVZF5EjyaO+FV985MA0GCSqG
# SIb3DQEBAQUABIIBgJfwzhspaoxol2lR0fsHOtj7pyMEnKipUrPpE0XZap5blh57
# fNXbdyZHaINJrv/dOGKsDMrWWZgeg0TSGyFuPh0+XmEOb8JCIUGpoCtaTyOs9R9a
# /937KBhm0oM78i05sdOwo0eAzOzBVPBLf4U5onnd0qk7GjVBZuuoRv+n4vhO0bF0
# A6OjQVjFG9UzDtwi/6tVQVpNKAtrcOdb4q+W+fxH/FucaslUNKD1OaBKFnTlM0vC
# K7qQH67wIZsFmf3f8VuYqC1qsqEzttdfew3BdDxwCstK0r6GH5gxho8FGD0Y6L24
# EAYK5CziGKoh+UrhueJzpbnNmg6RIpTnYH8AMn//IhJc94UXYszZ/FooRG8Z3NfX
# ogWDpWs1s3JE84Kw9pQn+oPLNYeozutgHAPklddPX3uhUOSXamFh+Wp7cMU9LAMa
# GrNhOP34QB6+erouwxA9UOYHsZwQekBnVcE0xvJ+6/2gG2PBTLGXLNlMtYwwNInD
# JZHWvswC8he4vBlq4KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTExNTE1
# MjFaMC8GCSqGSIb3DQEJBDEiBCAurS51qWJ8WV9o6w1t2XWhASgKi4vngbvzeG37
# qXxTdTANBgkqhkiG9w0BAQEFAASCAgAvT8HlfQc4YhrZxWwAlZ/kX0vSwTcg5OGY
# K6qffhXAp/kZKQw6brQ5cqm3JEkT90OBJSqU4dN7irsa0yOTuQuc67IqWcLuasY+
# AalFNv93LQ2+M9bfMTP5dYcFL5qsP2TqjEuCqy703mq/lK4FotcBu1iis4+raOm/
# 2HzmoGBnr6iZ4k93T2SQCLsapW1wa8hVAV8qdYCY1j6lwNp6dU70bv5/JmmmpLmY
# 4xW4y48GEzbyp5ek5Md0rO2xDLCFBAMEOB1QMEyDqZ+P+5vc1mbZ8n7na/eYKsUA
# S3Ydo+kUXAMeyL1HKqHRK7rDp+IgbBQSog5Hf0h0wx1w89T+7IpLw6UdRKw4PvAp
# 9nWavmEwrmWUU39iZ/pomV1LGaJOHMEN2cdd/6k6Nubs1s8XPWqRu7w3NNVnOogg
# NHfY60I8p3lob7oQQojsb2rEYqvE8gkiO8tp1ucPC53DYqL+9PmEwz4YmEEppUy1
# r8YmQUTDXdY72IWGPdwtYWs1TsS+iEWbr5fu9Wqd0fdlRpiVmOsMXrtjt+PTUoB/
# 1w4USsx/whXd9VFwE/E1+2z+7DaQfWK+PIJMEVu+kApe3lY+wjGeZt5TOmfZyay3
# 4BJyQZHxIelSk10USLuMt/AIrokYQrl63gDrXiLoyohV5i6S1eaYGpmpublMwWnH
# i/f5HL0L9Q==
# SIG # End signature block
