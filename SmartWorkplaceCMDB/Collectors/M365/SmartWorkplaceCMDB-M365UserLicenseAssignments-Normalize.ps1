<#
.SYNOPSIS
Normalizes Microsoft 365 user license assignments for Power BI.

.DESCRIPTION
Publishes one FactUserLicense.csv row per user and SKU. When Graph returns
multiple direct or group-derived states for the same user and SKU, the most
critical state is selected first (Error, ActiveWithError, Disabled, Active),
then the most recently updated state. AssignedDateTime contains Graph
lastUpdatedDateTime because the original assignment creation time is not
available on licenseAssignmentState.

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
    [string]$RawInputPath,
    [string]$CmdbUsersPath,
    [string]$DimLicenseSkuPath,
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
            Status = 'Missing'; MissingColumns = ($ExpectedColumns -join ', ')
            UnexpectedColumns = ''; OrderMatches = $false
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
    $orderMatches = (($actualColumns -join [char]31) -ceq ($ExpectedColumns -join [char]31))
    return [pscustomobject]@{
        Status = if ($missing.Count -eq 0 -and $unexpected.Count -eq 0 -and $orderMatches) { 'Valid' } else { 'Incompatible' }
        MissingColumns = ($missing -join ', ')
        UnexpectedColumns = ($unexpected -join ', ')
        OrderMatches = $orderMatches
    }
}

function Get-SmartWorkplaceCMDBAssignmentStatePriority {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$State)
    switch ($State.Trim().ToLowerInvariant()) {
        'error' { return 4 }
        'activewitherror' { return 3 }
        'disabled' { return 2 }
        'active' { return 1 }
        default { return 0 }
    }
}

function Get-SmartWorkplaceCMDBAssignmentDateValue {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Value,
        [Parameter(Mandatory)][string]$RawAssignmentKey
    )
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return [datetimeoffset]::MinValue
    }
    $parsed = [datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse(
            $Value,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal,
            [ref]$parsed
        )) {
        throw "LastUpdatedDateTime '$Value' is invalid for assignment '$RawAssignmentKey'."
    }
    return $parsed.ToUniversalTime()
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$modulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
Import-Module $modulePath -Force

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
$rawContract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$curatedContract = Get-SmartWorkplaceCMDBTableContract -Path $curatedContractPath
$rawTable = @($rawContract.tables | Where-Object name -eq 'M365_UserLicenseAssignments.csv')
$factTable = @($curatedContract.tables | Where-Object name -eq 'FactUserLicense.csv')
if ($rawTable.Count -ne 1 -or $factTable.Count -ne 1) {
    throw 'The contracts must contain one user license assignments and FactUserLicense definition.'
}
$rawTable = $rawTable[0]; $factTable = $factTable[0]
if ([string]::IsNullOrWhiteSpace($RawInputPath)) {
    $RawInputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$rawTable.area) ([string]$rawTable.name)
    )
}
if ([string]::IsNullOrWhiteSpace($CmdbUsersPath)) {
    $CmdbUsersPath = Join-Path $paths.CmdbLatestPath 'CMDB_Users.csv'
}
if ([string]::IsNullOrWhiteSpace($DimLicenseSkuPath)) {
    $DimLicenseSkuPath = Join-Path $paths.PowerBILatestPath 'DimLicenseSku.csv'
}
$RawInputPath = [IO.Path]::GetFullPath($RawInputPath)
$CmdbUsersPath = [IO.Path]::GetFullPath($CmdbUsersPath)
$DimLicenseSkuPath = [IO.Path]::GetFullPath($DimLicenseSkuPath)
$rawHeader = Test-SmartWorkplaceCMDBExactCsvHeader $RawInputPath @($rawTable.columns | ForEach-Object { [string]$_ })
if ($rawHeader.Status -eq 'Incompatible') { throw 'Raw user license assignments CSV is incompatible.' }
if ($rawHeader.Status -eq 'Missing' -and -not $ValidateOnly) {
    throw "Raw user license assignments CSV was not found: $RawInputPath"
}
$factOutputPath = Join-Path $paths.PowerBILatestPath ([string]$factTable.name)
$factHeader = Test-SmartWorkplaceCMDBExactCsvHeader $factOutputPath @($factTable.columns | ForEach-Object { [string]$_ })
if ($factHeader.Status -eq 'Incompatible') {
    throw "Existing curated table '$($factTable.name)' is incompatible."
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status = 'Valid'; ScriptVersion = $ScriptVersion
        RawContractVersion = [string]$rawContract.contractVersion
        CuratedContractVersion = [string]$curatedContract.contractVersion
        RawInputStatus = $rawHeader.Status; RawInputPath = $RawInputPath
        FactUserLicenseTargetStatus = $factHeader.Status
        FactUserLicenseOutputPath = $factOutputPath
        CmdbUsersReferenceExists = Test-Path -LiteralPath $CmdbUsersPath -PathType Leaf
        DimLicenseSkuReferenceExists = Test-Path -LiteralPath $DimLicenseSkuPath -PathType Leaf
    } | Format-List
    return
}

$rawRows = @(Import-Csv -LiteralPath $RawInputPath -ErrorAction Stop)
$identityFields = [ordered]@{
    TenantKey = $paths.TenantKey; OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey; TenantId = $paths.TenantId
}
foreach ($row in $rawRows) {
    foreach ($identityName in $identityFields.Keys) {
        if ([string]$row.$identityName -ne [string]$identityFields[$identityName]) {
            throw "Raw user license assignments identity mismatch for '$identityName'."
        }
    }
    foreach ($keyName in @('RawAssignmentKey', 'SourceUserId', 'SkuId', 'AssignmentState')) {
        if ([string]::IsNullOrWhiteSpace([string]$row.$keyName)) {
            throw "Raw user license assignments data contains an empty $keyName."
        }
    }
}
$duplicateRawKeys = @($rawRows | Group-Object RawAssignmentKey | Where-Object Count -gt 1)
if ($duplicateRawKeys.Count -gt 0) {
    throw "Raw user license assignments contain duplicate RawAssignmentKey values: $($duplicateRawKeys.Name -join ', ')"
}

$factRows = @($rawRows | Group-Object {
        '{0}|{1}' -f ([string]$_.SourceUserId).Trim().ToLowerInvariant(), ([string]$_.SkuId).Trim().ToLowerInvariant()
    } | ForEach-Object {
        $selected = @($_.Group | Sort-Object `
            @{ Expression = { Get-SmartWorkplaceCMDBAssignmentStatePriority ([string]$_.AssignmentState) }; Descending = $true },
            @{ Expression = { Get-SmartWorkplaceCMDBAssignmentDateValue ([string]$_.LastUpdatedDateTime) ([string]$_.RawAssignmentKey) }; Descending = $true },
            @{ Expression = { [string]$_.RawAssignmentKey }; Descending = $false })[0]
        $sourceUserId = ([string]$selected.SourceUserId).Trim()
        $skuId = ([string]$selected.SkuId).Trim()
        [pscustomobject][ordered]@{
            TenantUserKey = ('{0}|entra-user|{1}' -f $paths.TenantKey, $sourceUserId.ToLowerInvariant())
            TenantSkuKey = ('{0}|sku|{1}' -f $paths.TenantKey, $skuId.ToLowerInvariant())
            CmdbUserId = ('{0}|entra-user|{1}' -f $paths.TenantKey, $sourceUserId.ToLowerInvariant())
            SkuId = $skuId
            AssignmentState = [string]$selected.AssignmentState
            AssignedDateTime = [string]$selected.LastUpdatedDateTime
            SourceSystem = 'MicrosoftEntraID'
        }
    } | Sort-Object TenantUserKey, TenantSkuKey)

$knownUsers = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
if (Test-Path -LiteralPath $CmdbUsersPath -PathType Leaf) {
    foreach ($user in @(Import-Csv -LiteralPath $CmdbUsersPath)) {
        [void]$knownUsers.Add([string]$user.SourceUserId)
    }
}
$knownSkus = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
if (Test-Path -LiteralPath $DimLicenseSkuPath -PathType Leaf) {
    foreach ($sku in @(Import-Csv -LiteralPath $DimLicenseSkuPath)) {
        [void]$knownSkus.Add([string]$sku.SkuId)
    }
}
$orphanUserCount = if ($knownUsers.Count -gt 0) {
    @($rawRows | Select-Object -ExpandProperty SourceUserId -Unique | Where-Object { -not $knownUsers.Contains([string]$_) }).Count
}
else { 0 }
$orphanSkuCount = if ($knownSkus.Count -gt 0) {
    @($factRows | Select-Object -ExpandProperty SkuId -Unique | Where-Object { -not $knownSkus.Contains([string]$_) }).Count
}
else { 0 }

$identityExport = @{
    TenantKey = $paths.TenantKey; OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey; TenantId = $paths.TenantId
}
Export-SmartWorkplaceCMDBCsv -InputObject $factRows -Path $factOutputPath `
    -Columns @($factTable.columns | ForEach-Object { [string]$_ }) @identityExport
$factValidation = Test-SmartWorkplaceCMDBExactCsvHeader $factOutputPath @($factTable.columns | ForEach-Object { [string]$_ })
if ($factValidation.Status -ne 'Valid') {
    throw 'Normalized FactUserLicense output did not satisfy the curated contract.'
}

Write-Information (
    "SmartWorkplaceCMDB user license assignments normalization completed. Raw={0}; Facts={1}; Collapsed={2}; OrphanUsers={3}; OrphanSKUs={4}." -f
    $rawRows.Count, $factRows.Count, ($rawRows.Count - $factRows.Count), $orphanUserCount, $orphanSkuCount
) -InformationAction Continue
[pscustomobject]@{
    Status = 'Completed'; ScriptVersion = $ScriptVersion
    RawAssignmentCount = $rawRows.Count; FactAssignmentCount = $factRows.Count
    CollapsedAssignmentCount = $rawRows.Count - $factRows.Count
    DirectAssignmentCount = @($rawRows | Where-Object { [string]::IsNullOrWhiteSpace($_.AssignedByGroupId) }).Count
    GroupAssignmentCount = @($rawRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.AssignedByGroupId) }).Count
    OrphanUserCount = $orphanUserCount; OrphanSkuCount = $orphanSkuCount
    RawInputPath = $RawInputPath; FactUserLicenseOutputPath = $factOutputPath
    RawContractVersion = [string]$rawContract.contractVersion
    CuratedContractVersion = [string]$curatedContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDs8wiQ3WLSslJe
# 1aMChuVsz89UOlGQzcPyc4dnr7E1V6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILMgXxEGLAioA4Kaj5gFNMcElevNLqafCY3cuwDMsL47MA0GCSqG
# SIb3DQEBAQUABIIBgGCxvItwoyzffOusuGazXPFOIIqcgSoBNXKHqATLyEbHzXO3
# 3Ou+E66tAK0alDZVtBdtQlEzx+KrYDSWTibQ24ZTV4tqFJPJdnjsP2VGW4cr3KSc
# ascczKC1gni9bx2XjPHoDV7NhNOlJ5lZA2ltOk1dGnbVwqMskjQYTSCvay8StPbt
# qzJD5qriW0dCALAiiMC0pqep/IZ3pW8ZcxdmNqBF14jICVTG8EI+k8dd28A59EgH
# mAI49zFNQxMLoKfIBrGXXZ02WtK9AmGOZjG/ZsDyGt4NQVVXRBYDE6k3o1MRZ7xf
# wJ1mB/8bwqrENCShhL5ecYhtK8bhVTN6dK9J8oO/W5isLeWBXpT/xjYvRbdgEfGj
# YGcCeO94ll0ZzH2wmRVkmQ8HMiE9wvFX2KImj8x0BrwWs705yFZwlrK3GI5yutjy
# 9YIkgj934gHLo5tcgZcMAuUfMBVtjA+6rYggmkIyxo+cle6NoiT3gUqPxRHNB58I
# 5VC20ifXCTffb5jmUqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxOTA0
# MzFaMC8GCSqGSIb3DQEJBDEiBCAwGxACyxv3r5eGXD31AY+Wcp/Hth9DahJr8s5V
# bO4UtjANBgkqhkiG9w0BAQEFAASCAgDOwuapxq5rte5pfLP1Yqt7AEPWc6sc8YIj
# o96hs17g7X7mWEIn6ZwDqsjszpcOWxVNFxakdmU6b2mGicWOiJLncWQsDxRI/K/x
# 0Fr2nXINNGTnm1fsFp+d1KYczyvPQnO7y4AtPa+XS5Ke8o3xPQgLTyR1mJLqR75F
# bW+l/Iuq7bhg/X17dhjC6CWE6b3q5eAeuK0Yd05EdQibsBzAzCeOEsu8LVi/JbeD
# ffPHEtqvLD08yb/i5C+M0bvSRMS51O8p7AI2bZsDSuI9wh+5oamAYO0GMbBMgxaQ
# PCm3s7v/Mv8RWSKW/DQa4wBHyhh66baXqKtOoAR5FsWHcc+Zo9hqccX+3KVoGQKe
# lbztq3yqkogFYUoy8ZnaPf32QtRJvll3pQqvhJPOejYNkSnHF4+PJnJZR8+xctsR
# FqNi0MiKzI4ezSaQYfcEJuW0230KWg2spBysyYeBAlyWD//DHcE/zc9lNr494Kiu
# 3k22tE33mdjD5aIX0+JtRz7uezDJ26nb1b7rm2v8c4ko7LIhXWowjWUbNs9GJfh2
# vfuOl5MUk0oKyIeC2wNiLsTGt0PmMp7CD5vosL81aHif5xxoh0CIv7nTmxTTyO/t
# vDPeTkldeAvFmvvsyCxws0g9GYFyzNingyDm0SdTxxsnADS9FGes+IJmanXw+sBi
# lhO1jQPbkg==
# SIG # End signature block
