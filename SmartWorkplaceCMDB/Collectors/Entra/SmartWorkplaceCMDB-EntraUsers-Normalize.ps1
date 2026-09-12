<#
.SYNOPSIS
Normalizes the latest raw Entra users snapshot for SmartWorkplaceCMDB.

.DESCRIPTION
Validates the autonomous Entra users raw contract, enforces tenant identity,
and publishes curated CMDB_Users.csv and Power BI DimUser.csv tables. No
Microsoft Graph connection is performed by this script.

.VERSION
1.0.0
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')]
    [string]$Tenant = 'default',
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
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.0.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Test-SmartWorkplaceCMDBExactCsvHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string[]]$ExpectedColumns
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{
            Status            = 'Missing'
            MissingColumns    = ($ExpectedColumns -join ', ')
            UnexpectedColumns = ''
            OrderMatches      = $false
        }
    }

    $headerLine = Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop
    $actualColumns = if ([string]::IsNullOrWhiteSpace($headerLine)) {
        @()
    }
    else {
        @($headerLine.Split(',') | ForEach-Object { $_.Trim().Trim('"') })
    }
    $missingColumns = @($ExpectedColumns | Where-Object { $_ -notin $actualColumns })
    $unexpectedColumns = @($actualColumns | Where-Object { $_ -notin $ExpectedColumns })
    $orderMatches = (($actualColumns -join [char]31) -ceq ($ExpectedColumns -join [char]31))
    $status = if ($missingColumns.Count -eq 0 -and
        $unexpectedColumns.Count -eq 0 -and
        $orderMatches) {
        'Valid'
    }
    else {
        'Incompatible'
    }

    return [pscustomobject]@{
        Status            = $status
        MissingColumns    = ($missingColumns -join ', ')
        UnexpectedColumns = ($unexpectedColumns -join ', ')
        OrderMatches      = $orderMatches
    }
}

function ConvertTo-SmartWorkplaceCMDBNormalizedBoolean {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value,

        [Parameter(Mandatory)]
        [string]$FieldName,

        [Parameter(Mandatory)]
        [string]$SourceUserId
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $parsed = $false
    if (-not [bool]::TryParse($Value, [ref]$parsed)) {
        throw "$FieldName '$Value' is invalid for source user '$SourceUserId'."
    }
    return $parsed
}

function ConvertTo-SmartWorkplaceCMDBUsageLocation {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    $cleanValue = $Value.Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($cleanValue)) {
        return [pscustomobject]@{
            UsageLocation = ''
            CountryCode = ''
            CountryLabel = 'Unknown / unassigned'
            CountryStatus = 'Not provided'
        }
    }
    if ($cleanValue -notmatch '^[A-Z]{2}$') {
        return [pscustomobject]@{
            UsageLocation = $cleanValue
            CountryCode = ''
            CountryLabel = 'Unknown / unassigned'
            CountryStatus = 'Invalid'
        }
    }
    return [pscustomobject]@{
        UsageLocation = $cleanValue
        CountryCode = $cleanValue
        CountryLabel = $cleanValue
        CountryStatus = 'Reported'
    }
}

function Get-SmartWorkplaceCMDBDefaultConfidenceScore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Configuration
    )

    $value = 0.5
    if ($Configuration.Contains('DataQuality') -and
        $Configuration['DataQuality'] -is [System.Collections.IDictionary] -and
        $Configuration['DataQuality'].Contains('DefaultConfidenceScore')) {
        $candidate = 0.0
        if ([double]::TryParse(
                [string]$Configuration['DataQuality']['DefaultConfidenceScore'],
                [System.Globalization.NumberStyles]::Float,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [ref]$candidate
            )) {
            $value = $candidate
        }
    }
    if ($value -lt 0 -or $value -gt 1) {
        throw 'DataQuality.DefaultConfidenceScore must be between 0 and 1.'
    }
    return $value
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$modulePath = Join-Path -Path $projectRoot -ChildPath 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath = Join-Path -Path $projectRoot -ChildPath 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath = Join-Path -Path $projectRoot -ChildPath 'Schema\SmartWorkplaceCMDB.tables.json'

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

$rawTable = @($rawContract.tables | Where-Object name -eq 'Entra_Users.csv')
$cmdbUserTable = @($curatedContract.tables | Where-Object name -eq 'CMDB_Users.csv')
$dimUserTable = @($curatedContract.tables | Where-Object name -eq 'DimUser.csv')
if ($rawTable.Count -ne 1 -or $cmdbUserTable.Count -ne 1 -or $dimUserTable.Count -ne 1) {
    throw 'The SmartWorkplaceCMDB contracts must contain exactly one raw Entra users, CMDB users, and DimUser definition.'
}
$rawTable = $rawTable[0]
$cmdbUserTable = $cmdbUserTable[0]
$dimUserTable = $dimUserTable[0]

if ([string]::IsNullOrWhiteSpace($RawInputPath)) {
    $RawInputPath = Join-Path -Path $paths.LatestOutputRootPath -ChildPath (
        Join-Path -Path ([string]$rawTable.area) -ChildPath ([string]$rawTable.name)
    )
}
$RawInputPath = [System.IO.Path]::GetFullPath($RawInputPath)
$rawHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $RawInputPath `
    -ExpectedColumns @($rawTable.columns | ForEach-Object { [string]$_ })

if ($rawHeader.Status -eq 'Incompatible') {
    throw (
        "Raw Entra users CSV is incompatible. Missing=[{0}] Unexpected=[{1}] OrderMatches={2}. Path='{3}'" -f
        $rawHeader.MissingColumns,
        $rawHeader.UnexpectedColumns,
        $rawHeader.OrderMatches,
        $RawInputPath
    )
}
if ($rawHeader.Status -eq 'Missing' -and -not $ValidateOnly) {
    throw "Raw Entra users CSV was not found: $RawInputPath"
}

$cmdbOutputPath = Join-Path -Path $paths.CmdbLatestPath -ChildPath ([string]$cmdbUserTable.name)
$dimUserOutputPath = Join-Path -Path $paths.PowerBILatestPath -ChildPath ([string]$dimUserTable.name)
$existingCmdbHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $cmdbOutputPath `
    -ExpectedColumns @($cmdbUserTable.columns | ForEach-Object { [string]$_ })
$existingDimHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $dimUserOutputPath `
    -ExpectedColumns @($dimUserTable.columns | ForEach-Object { [string]$_ })

foreach ($target in @(
        [pscustomobject]@{ Name = $cmdbUserTable.name; Header = $existingCmdbHeader },
        [pscustomobject]@{ Name = $dimUserTable.name; Header = $existingDimHeader }
    )) {
    if ($target.Header.Status -eq 'Incompatible') {
        throw (
            "Existing curated table '{0}' is incompatible. Missing=[{1}] Unexpected=[{2}] OrderMatches={3}." -f
            $target.Name,
            $target.Header.MissingColumns,
            $target.Header.UnexpectedColumns,
            $target.Header.OrderMatches
        )
    }
}

if ($ValidateOnly) {
    if (Test-Path -LiteralPath $RawInputPath) { Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $RawInputPath -Paths $paths | Out-Null }
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status                 = 'Valid'
        ScriptVersion          = $ScriptVersion
        RawContractVersion     = [string]$rawContract.contractVersion
        CuratedContractVersion = [string]$curatedContract.contractVersion
        RawInputStatus         = $rawHeader.Status
        RawInputPath           = $RawInputPath
        CmdbUserTargetStatus   = $existingCmdbHeader.Status
        DimUserTargetStatus    = $existingDimHeader.Status
        CmdbUserOutputPath     = $cmdbOutputPath
        DimUserOutputPath      = $dimUserOutputPath
        TenantKey              = $paths.TenantKey
    } | Format-List
    return
}

$rawRows = @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $RawInputPath -Paths $paths -ErrorAction Stop)
$identityFields = [ordered]@{
    TenantKey       = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey  = $paths.EnvironmentKey
    TenantId        = $paths.TenantId
}

foreach ($row in $rawRows) {
    foreach ($identityName in $identityFields.Keys) {
        $expectedValue = [string]$identityFields[$identityName]
        $actualValue = [string]$row.$identityName
        if ($actualValue -ne $expectedValue) {
            throw (
                "Raw Entra users identity mismatch for '{0}'. Expected='{1}' Actual='{2}'." -f
                $identityName,
                $expectedValue,
                $actualValue
            )
        }
    }
    if ([string]::IsNullOrWhiteSpace([string]$row.SourceUserId)) {
        throw 'Raw Entra users data contains an empty SourceUserId.'
    }
}

$duplicateSourceIds = @($rawRows |
    Group-Object SourceUserId |
    Where-Object Count -gt 1 |
    Select-Object -ExpandProperty Name)
if ($duplicateSourceIds.Count -gt 0) {
    throw "Raw Entra users data contains duplicate SourceUserId values: $($duplicateSourceIds -join ', ')"
}

$confidenceScore = Get-SmartWorkplaceCMDBDefaultConfidenceScore -Configuration $context.Configuration
$confidenceText = $confidenceScore.ToString(
    '0.################',
    [System.Globalization.CultureInfo]::InvariantCulture
)

$cmdbRows = @($rawRows |
    Sort-Object UserPrincipalName, SourceUserId |
    ForEach-Object {
        $sourceUserId = [string]$_.SourceUserId
        $usageLocation = ConvertTo-SmartWorkplaceCMDBUsageLocation -Value ([string]$_.UsageLocation)
        $cmdbUserId = '{0}|entra-user|{1}' -f $paths.TenantKey, $sourceUserId.ToLowerInvariant()
        [pscustomobject][ordered]@{
            CmdbUserId             = $cmdbUserId
            SourceSystem           = [string]$_.SourceSystem
            SourceUserId           = $sourceUserId
            UserPrincipalName      = [string]$_.UserPrincipalName
            DisplayName            = [string]$_.DisplayName
            AccountEnabled         = ConvertTo-SmartWorkplaceCMDBNormalizedBoolean `
                -Value ([string]$_.AccountEnabled) `
                -FieldName 'AccountEnabled' `
                -SourceUserId $sourceUserId
            UserType               = [string]$_.UserType
            Department             = [string]$_.Department
            JobTitle               = [string]$_.JobTitle
            UsageLocation          = $usageLocation.UsageLocation
            UsageLocationStatus    = $usageLocation.CountryStatus
            ManagerUserId          = ''
            CreatedDateTime        = [string]$_.CreatedDateTime
            LastSignInDateTime     = ''
            ConfidenceScore        = $confidenceText
            SourceCollectedDateTime = [string]$_.SourceCollectedDateTime
        }
    })

$dimRows = @($cmdbRows | ForEach-Object {
    [pscustomobject][ordered]@{
        TenantUserKey    = $_.CmdbUserId
        CmdbUserId       = $_.CmdbUserId
        UserPrincipalName = $_.UserPrincipalName
        DisplayName      = $_.DisplayName
        AccountEnabled   = $_.AccountEnabled
        UserType         = $_.UserType
        Department       = $_.Department
        JobTitle         = $_.JobTitle
        CountryCode      = if ($_.UsageLocationStatus -eq 'Reported') { $_.UsageLocation } else { '' }
        CountryLabel     = if ($_.UsageLocationStatus -eq 'Reported') { $_.UsageLocation } else { 'Unknown / unassigned' }
        CountryStatus    = $_.UsageLocationStatus
        ConfidenceScore  = $_.ConfidenceScore
    }
})

$identityExportParameters = @{
    TenantKey       = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey
    TenantId        = $paths.TenantId
}
Export-SmartWorkplaceCMDBCsv `
    -InputObject $cmdbRows `
    -Path $cmdbOutputPath `
    -Columns @($cmdbUserTable.columns | ForEach-Object { [string]$_ }) `
    @identityExportParameters
Export-SmartWorkplaceCMDBCsv `
    -InputObject $dimRows `
    -Path $dimUserOutputPath `
    -Columns @($dimUserTable.columns | ForEach-Object { [string]$_ }) `
    @identityExportParameters

$cmdbValidation = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $cmdbOutputPath `
    -ExpectedColumns @($cmdbUserTable.columns | ForEach-Object { [string]$_ })
$dimValidation = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $dimUserOutputPath `
    -ExpectedColumns @($dimUserTable.columns | ForEach-Object { [string]$_ })
if ($cmdbValidation.Status -ne 'Valid' -or $dimValidation.Status -ne 'Valid') {
    throw 'The normalized Entra users outputs did not satisfy the curated SmartWorkplaceCMDB contracts.'
}

Write-Information (
    "SmartWorkplaceCMDB Entra users normalization completed. Users={0}; CMDB='{1}'; DimUser='{2}'." -f
    $cmdbRows.Count,
    $cmdbOutputPath,
    $dimUserOutputPath
) -InformationAction Continue

[pscustomobject]@{
    Status                 = 'Completed'
    ScriptVersion          = $ScriptVersion
    UserCount              = $cmdbRows.Count
    RawInputPath           = $RawInputPath
    CmdbUserOutputPath     = $cmdbOutputPath
    DimUserOutputPath      = $dimUserOutputPath
    RawContractVersion     = [string]$rawContract.contractVersion
    CuratedContractVersion = [string]$curatedContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBYQGAaH/owrzJZ
# XQNjRZrtIoanrvl1pk4z7y1d4TyGfKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIObQWQsx4jGYEmgjZIJJmHJzT+quR5yVVa24uaHOlVfAMA0GCSqG
# SIb3DQEBAQUABIIBgGZHk4RFz65XezuTLtZ/GNVUv/lrlzFJq4duVdq8GQo65E2R
# BKEmTdSkdvK6q2desrDbhzf8ZJnVgld6xUoLS8I/77TRf1JTcTGUrYE9v2vM8bT6
# 8f+ls4vR1+LlDOe/SmZjrHW4YwVcQvp4Y0p604uZABNLECq9AMyqUXw2VTLdxr00
# sm4RTvJYCTj1LVRcQs98ZYCCeGUQtHE0XiRb66c1W2ORBMTF7xHIfF4pfnXWH3St
# XB9t+ONejG6rwzsXMHfkOx7sseVpfBKAEbWaGuKPQHcmvHeLpq+vUEHR8u26ZnOe
# EOlCGx1SJOYKDW7EjYSp4PSbN1Py6Cd9ctk0nIZM4i3IdFVG3Is820jIDkzUWZhG
# H4Ghhb/yB44QFUrRWqy9GbPJx60c4Q9jVEg3op6mL38c6Ali/czGQT3o2OotFza4
# uARCRIvoDIOhFuGeuOm0j/aICCqOUDDJicH2R6IvfZQCYsDpbNB0TAfOnUUtcY4/
# g81nVsHoh9tkGUW8u6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTExNTE1
# MTJaMC8GCSqGSIb3DQEJBDEiBCA2DiFTUPpSXmnLiAHZ8/Eg6HCfYQqLuETIU+1b
# PhDQpDANBgkqhkiG9w0BAQEFAASCAgAO2VulPUiNqZU4k0mdKQ1tkWBO1XW6FBMx
# VMqjaNWTviOV/RQ9IwY3j8SF3Pyd9eGL+iyR7MWEPYvjg7y03Kh7SKID37+JNTIL
# 5h/O0TCoMc7jP1xan0pZmK5hEio81zE4MwLffx0Gihv4dxI852k/VToklsLKoKAy
# GWW2dLDQt/kKqXq2rFoeRfFRNd2rwPlyFAyI17N0uZCiyPEJ/GwDVremWWbnyFJf
# VwWT7yrzaNrttmo72lRNCMx0dtnXdEpP47yncnrm7cNeZJkqd6+zAkMZo6HnTpFB
# zmtGMmIjwE8Acs4YqZ288Uk4ZS2+ovpC0+2Q+kiMzRu3DoFkuwcdaUlFXyBar3bz
# VGJlrWDeuGVhduHdutyMP2b17QqZpE37ExzjPGkjYpK/D39bZpxoK2GJniPNVTun
# ibg6/FSn0haZpkf8+fR6+QN27vgFr0CFyFqcyj077GI83lZDNSyDWTFprxdMq4Xd
# hYbmutnDBVxSH4mPZKGfBHwmTrgnU19tQU8IrNWw1noulC+sWW+oF7lrMT35JlEz
# +yVCnF6I6dI7zqWNVKVUJpST8mH5zCwXvIHWOi1M8Yo8pOHOv36tjPYwpKmsINDu
# OR2S3CBEKdQMa8XOnknkxy90M3M5ncHN8DFgRfMBYhiaESi3O5AbVF0iO0+/LMtB
# ogwhrn7aXg==
# SIG # End signature block
