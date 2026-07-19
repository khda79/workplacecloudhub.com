<#
.SYNOPSIS
Publishes primary user-to-device relationships for SmartWorkplaceCMDB.

.DESCRIPTION
Correlates CMDB_Devices.PrimaryUserId with CMDB_Users.SourceUserId and publishes
the specialized CMDB relationship table plus its Power BI fact. No tenant
connection is required because the normalizer consumes existing curated data.

.VERSION
0.1.0
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')][string]$Tenant = 'default',
    [string]$OrganizationKey,
    [string]$EnvironmentKey,
    [string]$TenantKey,
    [string]$TenantId,
    [string]$DataRootPath,
    [string]$DataAllRootPath,
    [string]$LatestOutputRootPath,
    [string]$LogRootPath,
    [string]$GlobalConfigPath,
    [string]$TenantConfigPath,
    [string]$UserInputPath,
    [string]$DeviceInputPath,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '0.1.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Test-SmartWorkplaceCMDBExactCsvHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$ExpectedColumns
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{
            Status = 'Missing'
            MissingColumns = ($ExpectedColumns -join ', ')
            UnexpectedColumns = ''
            OrderMatches = $false
        }
    }
    $headerLine = Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop
    $actualColumns = if ([string]::IsNullOrWhiteSpace($headerLine)) {
        @()
    }
    else {
        @($headerLine.Split(',') | ForEach-Object { $_.Trim().Trim('"') })
    }
    $missing = @($ExpectedColumns | Where-Object { $_ -notin $actualColumns })
    $unexpected = @($actualColumns | Where-Object { $_ -notin $ExpectedColumns })
    $orderMatches = (
        ($actualColumns -join [char]31) -ceq
        ($ExpectedColumns -join [char]31)
    )
    return [pscustomobject]@{
        Status = if ($missing.Count -eq 0 -and
            $unexpected.Count -eq 0 -and
            $orderMatches) { 'Valid' } else { 'Incompatible' }
        MissingColumns = ($missing -join ', ')
        UnexpectedColumns = ($unexpected -join ', ')
        OrderMatches = $orderMatches
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$contractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
Import-Module $coreModulePath -Force

$boundParameterCopy = @{}
foreach ($key in $PSBoundParameters.Keys) {
    $boundParameterCopy[$key] = $PSBoundParameters[$key]
}
$context = Resolve-SmartWorkplaceCMDBContext `
    -BoundParameters $boundParameterCopy `
    -GlobalConfigPath $GlobalConfigPath `
    -TenantConfigPath $TenantConfigPath `
    -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths = $context.Paths
$contract = Get-SmartWorkplaceCMDBTableContract -Path $contractPath
$userTable = @($contract.tables | Where-Object name -eq 'CMDB_Users.csv')
$deviceTable = @($contract.tables | Where-Object name -eq 'CMDB_Devices.csv')
$relationshipTable = @($contract.tables |
    Where-Object name -eq 'CMDB_UserDeviceRelationships.csv')
$factTable = @($contract.tables |
    Where-Object name -eq 'FactUserDeviceRelationship.csv')
if ($userTable.Count -ne 1 -or
    $deviceTable.Count -ne 1 -or
    $relationshipTable.Count -ne 1 -or
    $factTable.Count -ne 1) {
    throw 'User-device relationship contracts are incomplete.'
}
$userTable = $userTable[0]
$deviceTable = $deviceTable[0]
$relationshipTable = $relationshipTable[0]
$factTable = $factTable[0]

if ([string]::IsNullOrWhiteSpace($UserInputPath)) {
    $UserInputPath = Join-Path $paths.CmdbLatestPath ([string]$userTable.name)
}
if ([string]::IsNullOrWhiteSpace($DeviceInputPath)) {
    $DeviceInputPath = Join-Path $paths.CmdbLatestPath ([string]$deviceTable.name)
}
$UserInputPath = [IO.Path]::GetFullPath($UserInputPath)
$DeviceInputPath = [IO.Path]::GetFullPath($DeviceInputPath)
$relationshipOutputPath = Join-Path $paths.CmdbLatestPath (
    [string]$relationshipTable.name
)
$factOutputPath = Join-Path $paths.PowerBILatestPath ([string]$factTable.name)

$userHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $UserInputPath `
    -ExpectedColumns @($userTable.columns | ForEach-Object { [string]$_ })
$deviceHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $DeviceInputPath `
    -ExpectedColumns @($deviceTable.columns | ForEach-Object { [string]$_ })
$relationshipHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $relationshipOutputPath `
    -ExpectedColumns @($relationshipTable.columns | ForEach-Object { [string]$_ })
$factHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $factOutputPath `
    -ExpectedColumns @($factTable.columns | ForEach-Object { [string]$_ })

foreach ($sourceInput in @(
        [pscustomobject]@{ Name = $userTable.name; Header = $userHeader; Path = $UserInputPath },
        [pscustomobject]@{ Name = $deviceTable.name; Header = $deviceHeader; Path = $DeviceInputPath }
    )) {
    if ($sourceInput.Header.Status -eq 'Incompatible') {
        throw "Curated input '$($sourceInput.Name)' is incompatible."
    }
    if ($sourceInput.Header.Status -eq 'Missing' -and -not $ValidateOnly) {
        throw "Curated input '$($sourceInput.Name)' was not found: $($sourceInput.Path)"
    }
}
foreach ($target in @(
        [pscustomobject]@{ Name = $relationshipTable.name; Header = $relationshipHeader },
        [pscustomobject]@{ Name = $factTable.name; Header = $factHeader }
    )) {
    if ($target.Header.Status -eq 'Incompatible') {
        throw "Existing curated table '$($target.Name)' is incompatible."
    }
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status                   = 'Valid'
        ScriptVersion            = $ScriptVersion
        ContractVersion          = [string]$contract.contractVersion
        UserInputStatus          = $userHeader.Status
        DeviceInputStatus        = $deviceHeader.Status
        RelationshipTargetStatus = $relationshipHeader.Status
        FactTargetStatus         = $factHeader.Status
        UserInputPath            = $UserInputPath
        DeviceInputPath          = $DeviceInputPath
        RelationshipOutputPath   = $relationshipOutputPath
        FactOutputPath           = $factOutputPath
        TenantKey                = $paths.TenantKey
    } | Format-List
    return
}

$users = @(Import-Csv -LiteralPath $UserInputPath -ErrorAction Stop)
$devices = @(Import-Csv -LiteralPath $DeviceInputPath -ErrorAction Stop)
$identityFields = [ordered]@{
    TenantKey       = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey  = $paths.EnvironmentKey
    TenantId        = $paths.TenantId
}
foreach ($source in @(
        [pscustomobject]@{ Name = 'CMDB users'; Rows = $users; Key = 'CmdbUserId' },
        [pscustomobject]@{ Name = 'CMDB devices'; Rows = $devices; Key = 'CmdbDeviceId' }
    )) {
    foreach ($row in @($source.Rows)) {
        foreach ($identityName in $identityFields.Keys) {
            if ([string]$row.$identityName -ne
                [string]$identityFields[$identityName]) {
                throw "$($source.Name) identity mismatch for '$identityName'."
            }
        }
        if ([string]::IsNullOrWhiteSpace([string]$row.($source.Key))) {
            throw "$($source.Name) contains an empty $($source.Key)."
        }
    }
    $duplicates = @($source.Rows |
        Group-Object -Property $source.Key |
        Where-Object Count -gt 1)
    if ($duplicates.Count -gt 0) {
        throw "$($source.Name) contains duplicate $($source.Key) values."
    }
}

$userBySourceId = @{}
foreach ($user in $users) {
    $sourceUserId = ([string]$user.SourceUserId).Trim()
    if (-not [string]::IsNullOrWhiteSpace($sourceUserId)) {
        if ($userBySourceId.ContainsKey($sourceUserId)) {
            throw "CMDB users contains duplicate SourceUserId '$sourceUserId'."
        }
        $userBySourceId[$sourceUserId] = $user
    }
}

$candidateCount = 0
$orphanCount = 0
$relationshipRows = New-Object System.Collections.Generic.List[object]
foreach ($device in @($devices | Sort-Object CmdbDeviceId)) {
    $primaryUserId = ([string]$device.PrimaryUserId).Trim()
    if ([string]::IsNullOrWhiteSpace($primaryUserId)) {
        continue
    }
    $candidateCount++
    if (-not $userBySourceId.ContainsKey($primaryUserId)) {
        $orphanCount++
        continue
    }
    $user = $userBySourceId[$primaryUserId]
    $sourceDeviceId = ([string]$device.SourceDeviceId).Trim()
    if ([string]::IsNullOrWhiteSpace($sourceDeviceId)) {
        throw "CMDB device '$($device.CmdbDeviceId)' has no SourceDeviceId."
    }
    $relationshipId = '{0}|user-device|{1}|{2}' -f
        $paths.TenantKey,
        $primaryUserId.ToLowerInvariant(),
        $sourceDeviceId.ToLowerInvariant()
    $relationshipRows.Add([pscustomobject][ordered]@{
        CmdbRelationshipId       = $relationshipId
        CmdbUserId               = [string]$user.CmdbUserId
        CmdbDeviceId             = [string]$device.CmdbDeviceId
        RelationshipType         = 'PrimaryUser'
        SourceSystem             = 'MicrosoftIntune'
        ConfidenceScore          = [string]$device.ConfidenceScore
        Evidence                 = 'managedDevices.userId'
        SourceCollectedDateTime  = [string]$device.SourceCollectedDateTime
    })
}
$relationships = @($relationshipRows.ToArray() |
    Sort-Object CmdbUserId, CmdbDeviceId)
$duplicateRelationships = @($relationships |
    Group-Object CmdbRelationshipId |
    Where-Object Count -gt 1)
if ($duplicateRelationships.Count -gt 0) {
    throw 'User-device normalization produced duplicate relationship IDs.'
}

$facts = @($relationships | ForEach-Object {
    [pscustomobject][ordered]@{
        TenantRelationshipKey = $_.CmdbRelationshipId
        TenantUserKey         = $_.CmdbUserId
        TenantDeviceKey       = $_.CmdbDeviceId
        CmdbRelationshipId    = $_.CmdbRelationshipId
        CmdbUserId            = $_.CmdbUserId
        CmdbDeviceId          = $_.CmdbDeviceId
        RelationshipType      = $_.RelationshipType
        ConfidenceScore       = $_.ConfidenceScore
        SourceSystem          = $_.SourceSystem
    }
})

$identityExport = @{
    TenantKey       = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey
    TenantId        = $paths.TenantId
}
Export-SmartWorkplaceCMDBCsv `
    -InputObject $relationships `
    -Path $relationshipOutputPath `
    -Columns @($relationshipTable.columns | ForEach-Object { [string]$_ }) `
    @identityExport
Export-SmartWorkplaceCMDBCsv `
    -InputObject $facts `
    -Path $factOutputPath `
    -Columns @($factTable.columns | ForEach-Object { [string]$_ }) `
    @identityExport

$relationshipValidation = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $relationshipOutputPath `
    -ExpectedColumns @($relationshipTable.columns | ForEach-Object { [string]$_ })
$factValidation = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $factOutputPath `
    -ExpectedColumns @($factTable.columns | ForEach-Object { [string]$_ })
if ($relationshipValidation.Status -ne 'Valid' -or
    $factValidation.Status -ne 'Valid') {
    throw 'User-device relationship outputs do not satisfy the curated contracts.'
}

Write-Information (
    "SmartWorkplaceCMDB user-device normalization completed. Candidates={0}; Relationships={1}; Orphans={2}." -f
    $candidateCount,
    $relationships.Count,
    $orphanCount
) -InformationAction Continue

[pscustomobject]@{
    Status                 = 'Completed'
    ScriptVersion          = $ScriptVersion
    UserCount              = $users.Count
    DeviceCount            = $devices.Count
    CandidatePrimaryUsers  = $candidateCount
    RelationshipCount      = $relationships.Count
    OrphanPrimaryUserCount = $orphanCount
    RelationshipOutputPath = $relationshipOutputPath
    FactOutputPath         = $factOutputPath
    ContractVersion        = [string]$contract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCkeyGgMbv9n8Z2
# rUgEJ7XlMZ04z+xuHuwrU6FlwlT18KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJ0r5G9ENdkWLQQ17WZorQ9V0pucE3HEWIAatvYk0WWEMA0GCSqG
# SIb3DQEBAQUABIIBgBYLlkdu7BK+zgGox3lJyDANZBj2nVHWTaR80dUvzV9AGAC/
# UAF/wWHaCTUaN35YnQ0NPxXaEx6A3+hXfRCtJo8NSsiSdWTzMcI4nawUHF3b9Pgv
# 75ijkO4sIH70w2dLAIEdthNlx+Ym9StI0nHCIhBqBPBkU3S0MQGGR6IU1KyachCM
# XdkpnGtVu+uan0HpAX8uksHfeBjwM2KcTFOTLGa4zTJ/CuPYN2lilMUxgwuTajd0
# ww2d1NfI7szatVntRyahVdQEqZltlkLIBbKFGWRpjQaksELbppE6bh3ToHgi9K9T
# 8c3IN2xNjZ8T1Pud6H5RbzCt9aeO1UASUMyyafd/eUwGLrqlLOLkmWYbTuOcTnwI
# eiwFaynO2o0v9XvZqcb7qj6lFJX/SGQ+YVLllZWZ8meIQyIXz9uKIbBcoBpCSY1y
# CMzNlVS5YQF99T7VPvZhwoazE8RS/sB9TjS1RTqEmqI0E15GZnI4QWjw4/ERQ7NF
# 3sCpC4FS7oUcTayq2aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxOTQy
# MjlaMC8GCSqGSIb3DQEJBDEiBCDPgP4x1rigPbR4dK2kbV6O/KE3AJ0rR5KhmZND
# CXHoMjANBgkqhkiG9w0BAQEFAASCAgBIXbD+ePj1tioRJ4s7RDyTgZTQ1tb9C0nc
# aaloApNhpeuHLsw4sC0tXtU7VQWS70qYxWXtqFySMf3XqaGRbPcQVR8QrIuYX6Ol
# 41xSezJodOreyE8eXUdburDgtmBXwBahKaBkWlB/4cNj/9LZwVZsV6vbHqJLzJb/
# keRAq6sANpRI7bcUYzTwVe9m38blAQx1huu8VBUJPvdjzdk3Zd4ikJ398SI2mabU
# 31TF9DGHmmhVAQSfq+6XKORcyPTKaBosqpj8gwutmas5Lmzu9R+Fx5fLAqVwbfof
# riDwVgUVub+KmSxaghskAupRl6wLirjGuRMXf0Ek8ueewu1z9YwW4BWomDOVlJYR
# VoIQuhOGJFbAPUa01STboJnolEydgxY9ltWtgOFJAge1IoV8P7QclTF9wUaww+F1
# Mh2QWC5AQUP2cm2xiC5OI7AyuE05xFtKIWDIJ8uLLzafHeKPym86Zn6S0YvG+GRL
# uaHl8fpzUL1kcGF6dacmEJoB25X72pkiEeClHV0yXfuENAdi9Hs6gefIYcy5xQf0
# Fy4tRT4kzBBoi8XJsA+5BFYMrQZhr4FQpw9KqVV2gZdTblyazNbVMwBfwEO+g2Yg
# ++F0O953Lh8q6zWOp6JIvGWA/oH6mnRoi8A8woWNOBPV1xQukKmemn49B1LK0dyG
# VlsP3brYcg==
# SIG # End signature block
