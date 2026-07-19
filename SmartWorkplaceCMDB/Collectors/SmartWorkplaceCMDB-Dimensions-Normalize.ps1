<#
.SYNOPSIS
Publishes autonomous SmartWorkplaceCMDB tenant and date dimensions.

.DESCRIPTION
Builds DimTenant.csv from the resolved local tenant profile and DimDate.csv from
the full calendar-year range found in local curated CSV date columns. Optional
StartDate and EndDate parameters can override the source-derived boundaries.
The normalizer performs no tenant connection.

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
    [string]$DateSourceRootPath,
    [datetime]$StartDate,
    [datetime]$EndDate,
    [datetimeoffset]$ReferenceDateTime = [datetimeoffset]::UtcNow,
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

function Get-SmartWorkplaceCMDBDateRange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceRootPath,
        [Parameter(Mandatory)][datetimeoffset]$FallbackDateTime
    )
    $dates = New-Object System.Collections.Generic.List[datetimeoffset]
    $excludedDateCount = 0
    $maximumSupportedYear = $FallbackDateTime.ToUniversalTime().Year + 10
    $files = @()
    if (Test-Path -LiteralPath $SourceRootPath -PathType Container) {
        $files = @(Get-ChildItem `
            -LiteralPath $SourceRootPath `
            -Recurse `
            -Filter '*.csv' `
            -File |
            Where-Object Name -notin @('DimTenant.csv', 'DimDate.csv') |
            Where-Object FullName -notmatch '[\\/]Raw[\\/]')
    }
    foreach ($file in $files) {
        $rows = @(Import-Csv -LiteralPath $file.FullName)
        if ($rows.Count -eq 0) {
            continue
        }
        $dateColumns = @($rows[0].PSObject.Properties.Name |
            Where-Object { $_ -match '(?i)(Date|DateTime)$' })
        foreach ($row in $rows) {
            foreach ($column in $dateColumns) {
                $value = [string]$row.$column
                if ([string]::IsNullOrWhiteSpace($value)) {
                    continue
                }
                $parsed = [datetimeoffset]::MinValue
                if ([datetimeoffset]::TryParse(
                        $value,
                        [Globalization.CultureInfo]::InvariantCulture,
                        [Globalization.DateTimeStyles]::AssumeUniversal,
                        [ref]$parsed
                    )) {
                    $parsed = $parsed.ToUniversalTime()
                    if ($parsed.Year -lt 1900 -or
                        $parsed.Year -gt $maximumSupportedYear) {
                        $excludedDateCount++
                    }
                    else {
                        $dates.Add($parsed)
                    }
                }
            }
        }
    }
    if ($dates.Count -eq 0) {
        $fallbackYear = $FallbackDateTime.ToUniversalTime().Year
        return [pscustomobject]@{
            StartDate = [datetime]::new($fallbackYear, 1, 1)
            EndDate = [datetime]::new($fallbackYear, 12, 31)
            ParsedDateCount = 0
            SourceFileCount = $files.Count
            ExcludedDateCount = $excludedDateCount
        }
    }
    $ordered = @($dates | Sort-Object)
    return [pscustomobject]@{
        StartDate = [datetime]::new($ordered[0].Year, 1, 1)
        EndDate = [datetime]::new($ordered[-1].Year, 12, 31)
        ParsedDateCount = $dates.Count
        SourceFileCount = $files.Count
        ExcludedDateCount = $excludedDateCount
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
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
$dimTenantTable = @($contract.tables | Where-Object name -eq 'DimTenant.csv')
$dimDateTable = @($contract.tables | Where-Object name -eq 'DimDate.csv')
if ($dimTenantTable.Count -ne 1 -or $dimDateTable.Count -ne 1) {
    throw 'The curated contract must contain DimTenant.csv and DimDate.csv.'
}
$dimTenantTable = $dimTenantTable[0]
$dimDateTable = $dimDateTable[0]

if ([string]::IsNullOrWhiteSpace($DateSourceRootPath)) {
    $DateSourceRootPath = $paths.LatestOutputRootPath
}
$derivedRange = Get-SmartWorkplaceCMDBDateRange `
    -SourceRootPath $DateSourceRootPath `
    -FallbackDateTime $ReferenceDateTime
$effectiveStartDate = if ($PSBoundParameters.ContainsKey('StartDate')) {
    $StartDate.Date
}
else {
    $derivedRange.StartDate
}
$effectiveEndDate = if ($PSBoundParameters.ContainsKey('EndDate')) {
    $EndDate.Date
}
else {
    $derivedRange.EndDate
}
if ($effectiveEndDate -lt $effectiveStartDate) {
    throw 'EndDate must be greater than or equal to StartDate.'
}
if (($effectiveEndDate - $effectiveStartDate).TotalDays -gt 36525) {
    throw 'The DimDate range cannot exceed 100 years.'
}

$dimTenantOutputPath = Join-Path $paths.PowerBILatestPath 'DimTenant.csv'
$dimDateOutputPath = Join-Path $paths.PowerBILatestPath 'DimDate.csv'
$dimTenantHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $dimTenantOutputPath `
    -ExpectedColumns @($dimTenantTable.columns |
        ForEach-Object { [string]$_ })
$dimDateHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $dimDateOutputPath `
    -ExpectedColumns @($dimDateTable.columns |
        ForEach-Object { [string]$_ })
if ($dimTenantHeader.Status -eq 'Incompatible' -or
    $dimDateHeader.Status -eq 'Incompatible') {
    throw 'Existing dimension output is incompatible with the curated contract.'
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status = 'Valid'
        ScriptVersion = $ScriptVersion
        ContractVersion = [string]$contract.contractVersion
        StartDate = $effectiveStartDate.ToString('yyyy-MM-dd')
        EndDate = $effectiveEndDate.ToString('yyyy-MM-dd')
        ParsedDateCount = $derivedRange.ParsedDateCount
        SourceFileCount = $derivedRange.SourceFileCount
        ExcludedDateCount = $derivedRange.ExcludedDateCount
        DimTenantTargetStatus = $dimTenantHeader.Status
        DimDateTargetStatus = $dimDateHeader.Status
        DimTenantOutputPath = $dimTenantOutputPath
        DimDateOutputPath = $dimDateOutputPath
    }
    return
}

$displayName = [string]$paths.TenantKey
if ($context.Configuration.Contains('DisplayName') -and
    -not [string]::IsNullOrWhiteSpace(
        [string]$context.Configuration['DisplayName']
    )) {
    $displayName = [string]$context.Configuration['DisplayName']
}
$lastRefreshDateTime = $ReferenceDateTime.ToUniversalTime().ToString(
    'yyyy-MM-ddTHH:mm:ss.fffffffZ',
    [Globalization.CultureInfo]::InvariantCulture
)
$tenantRow = [pscustomobject][ordered]@{
    TenantKey = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey
    TenantId = $paths.TenantId
    TenantDisplayName = $displayName
    Environment = $paths.EnvironmentKey
    LastRefreshDateTime = $lastRefreshDateTime
}

$dateRows = New-Object System.Collections.Generic.List[object]
$currentDate = $effectiveStartDate
while ($currentDate -le $effectiveEndDate) {
    $quarter = [math]::Ceiling($currentDate.Month / 3.0)
    $dateRows.Add([pscustomobject][ordered]@{
        Date = $currentDate.ToString(
            'yyyy-MM-dd',
            [Globalization.CultureInfo]::InvariantCulture
        )
        Year = $currentDate.Year
        Quarter = "Q$quarter"
        Month = $currentDate.Month
        MonthName = $currentDate.ToString(
            'MMMM',
            [Globalization.CultureInfo]::InvariantCulture
        )
        Day = $currentDate.Day
    })
    $currentDate = $currentDate.AddDays(1)
}

Initialize-SmartWorkplaceCMDBTenantFolder -Paths $paths
Export-SmartWorkplaceCMDBCsv `
    -InputObject @($tenantRow) `
    -Path $dimTenantOutputPath `
    -Columns @($dimTenantTable.columns |
        ForEach-Object { [string]$_ }) | Out-Null
Export-SmartWorkplaceCMDBCsv `
    -InputObject @($dateRows.ToArray()) `
    -Path $dimDateOutputPath `
    -Columns @($dimDateTable.columns |
        ForEach-Object { [string]$_ }) | Out-Null

$dimTenantValidation = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $dimTenantOutputPath `
    -ExpectedColumns @($dimTenantTable.columns |
        ForEach-Object { [string]$_ })
$dimDateValidation = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $dimDateOutputPath `
    -ExpectedColumns @($dimDateTable.columns |
        ForEach-Object { [string]$_ })
if ($dimTenantValidation.Status -ne 'Valid' -or
    $dimDateValidation.Status -ne 'Valid') {
    throw 'Published dimension output does not satisfy the curated contract.'
}

Write-Information (
    "SmartWorkplaceCMDB dimension normalization completed. Tenant=1; Dates={0}; Range={1}..{2}." -f
    $dateRows.Count,
    $effectiveStartDate.ToString('yyyy-MM-dd'),
    $effectiveEndDate.ToString('yyyy-MM-dd')
) -InformationAction Continue

[pscustomobject]@{
    Status = 'Completed'
    ScriptVersion = $ScriptVersion
    TenantCount = 1
    DateCount = $dateRows.Count
    StartDate = $effectiveStartDate.ToString('yyyy-MM-dd')
    EndDate = $effectiveEndDate.ToString('yyyy-MM-dd')
    ParsedDateCount = $derivedRange.ParsedDateCount
    SourceFileCount = $derivedRange.SourceFileCount
    ExcludedDateCount = $derivedRange.ExcludedDateCount
    DimTenantOutputPath = $dimTenantOutputPath
    DimDateOutputPath = $dimDateOutputPath
    ContractVersion = [string]$contract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDfr2n1yCGZQ040
# BkoQKeTYV7CvQ3fKIl91GvjH7KGRZ6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOE5xn9+WEv2h63sHrNL15RusUPih9LzbCTjFKnc12T6MA0GCSqG
# SIb3DQEBAQUABIIBgDkTpTZYZk5FMRlrFP6XsDqNTBVeifxEuRK+JeNGXN+hxK4p
# SnGlRrXWv86OpQaTd/z+GPcpUuAAU0mVg2XxetTZNXQXdsIrL8zdFcNUHTR4Ezv3
# dIx+6WqN4t8b9tbS2k7KXbuKqWPMHItj02chs89YAj2MB/UXxEO7hsHjRgN1J+Mx
# 4lLSXYzcfvEyn3Rm+Z9ZCT8XToo6AzARpAyy99fmDJJe3XIDhm19VTNFYoAwXAaD
# f1s7htJ4yGL+MoWTtleUugdt+seByRsO6tWgaiSFI/SOJ0hAOuAmcSkmAb1Pvh67
# 6vcKDGVc5sAgoixMXGFzx9Fk7F/frFcpneJ+VKr9PHuELnxBrpFaH1xWQPH8sa+P
# mjZ5wvoU087gWPpgeA27sHsuxdnFkK6KDwQdcA3g2vGSd8zc5InQ69MZTYOHhPW4
# jpgFRq4xrF/Fi9QErUGyrmU7nsuCIPCrju1tQVIEzw1vStYpGBPZpQ6fZSJchPWz
# O6i9LBeVmgTd8I+7BKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMTA5
# NTVaMC8GCSqGSIb3DQEJBDEiBCDIRPkX2i7czHWq1ZzWrjgLMEvHkIIRSw8/0JVE
# mTtgMTANBgkqhkiG9w0BAQEFAASCAgCFAY28SnyLEawsQgt42Nr4Jaq7jAYgL8es
# /W1ayCC8CnEvJ+3JHIVn9+XohkmaVNiyqJ+DvkmoDyPRY8Q/HuGPDyplxTkwTiNr
# inkbwQgX8mPRsC2P8slJFGsyvvHJqe889kyS8usm3Tjfz3COu8LqhPLoxKKOL3ND
# i7wrO46T+qCOt2jCLZ1vOWBLyP/YRCaZHkT+xiToLPY2rzLf8kyNq+qwry3t+79E
# lhXaicwQiGQFF4/AJvJ5JBx0GOlr/2ihBz5TN1Bc7v7omQ3K/Q3rlOHAMhLgJ1Zs
# 6boS8EBJUuYXyQUSi/ApDVeSSCddknMGPATvnpA8UoGfc7N8sTZSrW8/zTOQQtEI
# 3fdBgbeQA/ihICPxT9RQemTnNhD4xinp9LDL5iOciIxeEp/2wMsyv3Hnt16XjiL6
# jaEYXlpTWYI3F2OvBAU7J3b8UecBn83BFcQca+0U3yX8WQCvnU88rqCpfY0iBrqW
# NZ7Sd1jeXXETRld/i1NyW1bx2SpZ/O2+oczf9uloRExvP6HLkSRRny4e6YB2LirE
# ym1cbQ11VEvNw6HQ3uwwWzTA+XtuNnxIBnRcDMqBcy/ZzZqU5n9gO1WnUaNf4Srr
# M0WTEwSoBSAbcCXLU0eXcfOl/lPuni02dY/nmftmCOFuRYOczdcFiybpX1XylYTb
# bWV6Md2duQ==
# SIG # End signature block
