<#
.SYNOPSIS
Normalizes Microsoft 365 subscribed SKUs for SmartWorkplaceCMDB.

.DESCRIPTION
Validates the raw subscribed SKU contract and tenant identity, aggregates any
duplicate SKU subscriptions by skuId, and publishes CMDB_Licenses.csv plus
Power BI DimLicenseSku.csv and DimLicenseServicePlan.csv.

.VERSION
1.1.0
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
    [string]$RawServicePlansInputPath,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.1.0'
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

function ConvertTo-SmartWorkplaceCMDBLicenseInteger {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Value,
        [Parameter(Mandatory)][string]$FieldName,
        [Parameter(Mandatory)][string]$SubscribedSkuId
    )
    $parsed = [long]0
    if (-not [long]::TryParse(
            $Value,
            [Globalization.NumberStyles]::Integer,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$parsed
        )) {
        throw "$FieldName '$Value' is invalid for subscribed SKU '$SubscribedSkuId'."
    }
    return $parsed
}

function Get-SmartWorkplaceCMDBLicenseDateValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Value,
        [Parameter(Mandatory)][string]$SubscribedSkuId
    )
    $parsed = [datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse(
            $Value,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal,
            [ref]$parsed
        )) {
        throw "SourceCollectedDateTime '$Value' is invalid for subscribed SKU '$SubscribedSkuId'."
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
$rawTable = @($rawContract.tables | Where-Object name -eq 'M365_SubscribedSkus.csv')
$rawServicePlanTable = @($rawContract.tables | Where-Object name -eq 'M365_ServicePlans.csv')
$cmdbTable = @($curatedContract.tables | Where-Object name -eq 'CMDB_Licenses.csv')
$dimTable = @($curatedContract.tables | Where-Object name -eq 'DimLicenseSku.csv')
$dimServicePlanTable = @($curatedContract.tables | Where-Object name -eq 'DimLicenseServicePlan.csv')
if ($rawTable.Count -ne 1 -or $rawServicePlanTable.Count -ne 1 -or $cmdbTable.Count -ne 1 -or
    $dimTable.Count -ne 1 -or $dimServicePlanTable.Count -ne 1) {
    throw 'The contracts must contain the subscribed SKU and service-plan source and dimension definitions.'
}
$rawTable = $rawTable[0]; $rawServicePlanTable = $rawServicePlanTable[0]
$cmdbTable = $cmdbTable[0]; $dimTable = $dimTable[0]; $dimServicePlanTable = $dimServicePlanTable[0]

if ([string]::IsNullOrWhiteSpace($RawInputPath)) {
    $RawInputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$rawTable.area) ([string]$rawTable.name)
    )
}
$RawInputPath = [IO.Path]::GetFullPath($RawInputPath)
$RawServicePlansInputPath = if ([string]::IsNullOrWhiteSpace($RawServicePlansInputPath)) {
    Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$rawServicePlanTable.area) ([string]$rawServicePlanTable.name))
}
else { [IO.Path]::GetFullPath($RawServicePlansInputPath) }
$rawHeader = Test-SmartWorkplaceCMDBExactCsvHeader $RawInputPath @($rawTable.columns | ForEach-Object { [string]$_ })
$rawServicePlanHeader = Test-SmartWorkplaceCMDBExactCsvHeader $RawServicePlansInputPath @($rawServicePlanTable.columns | ForEach-Object { [string]$_ })
if ($rawHeader.Status -eq 'Incompatible' -or $rawServicePlanHeader.Status -eq 'Incompatible') {
    throw 'Raw subscribed SKUs or service-plan CSV is incompatible.'
}
if ($rawHeader.Status -eq 'Missing' -and -not $ValidateOnly) {
    throw 'The raw subscribed SKU CSV was not found.'
}

$cmdbOutputPath = Join-Path $paths.CmdbLatestPath ([string]$cmdbTable.name)
$dimOutputPath = Join-Path $paths.PowerBILatestPath ([string]$dimTable.name)
$dimServicePlanOutputPath = Join-Path $paths.PowerBILatestPath ([string]$dimServicePlanTable.name)
$cmdbHeader = Test-SmartWorkplaceCMDBExactCsvHeader $cmdbOutputPath @($cmdbTable.columns | ForEach-Object { [string]$_ })
$dimHeader = Test-SmartWorkplaceCMDBExactCsvHeader $dimOutputPath @($dimTable.columns | ForEach-Object { [string]$_ })
$dimServicePlanHeader = Test-SmartWorkplaceCMDBExactCsvHeader $dimServicePlanOutputPath @($dimServicePlanTable.columns | ForEach-Object { [string]$_ })
foreach ($target in @(
        [pscustomobject]@{ Name = $cmdbTable.name; Header = $cmdbHeader },
        [pscustomobject]@{ Name = $dimTable.name; Header = $dimHeader },
        [pscustomobject]@{ Name = $dimServicePlanTable.name; Header = $dimServicePlanHeader }
    )) {
    if ($target.Header.Status -eq 'Incompatible') {
        throw "Existing curated table '$($target.Name)' is incompatible."
    }
}

if ($ValidateOnly) {
    if (Test-Path -LiteralPath $RawInputPath) { Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $RawInputPath -Paths $paths | Out-Null }
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status = 'Valid'; ScriptVersion = $ScriptVersion
        RawContractVersion = [string]$rawContract.contractVersion
        CuratedContractVersion = [string]$curatedContract.contractVersion
        RawInputStatus = $rawHeader.Status; RawInputPath = $RawInputPath
        RawServicePlansInputStatus = $rawServicePlanHeader.Status
        CmdbLicenseTargetStatus = $cmdbHeader.Status
        DimLicenseSkuTargetStatus = $dimHeader.Status
        DimLicenseServicePlanTargetStatus = $dimServicePlanHeader.Status
        CmdbLicenseOutputPath = $cmdbOutputPath; DimLicenseSkuOutputPath = $dimOutputPath
        DimLicenseServicePlanOutputPath = $dimServicePlanOutputPath
    } | Format-List
    return
}

$rawRows = @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $RawInputPath -Paths $paths -ErrorAction Stop)
$rawServicePlanRows = if (Test-Path -LiteralPath $RawServicePlansInputPath -PathType Leaf) {
    @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $RawServicePlansInputPath -Paths $paths -ErrorAction Stop)
}
else { @() }
$identityFields = [ordered]@{
    TenantKey = $paths.TenantKey; OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey; TenantId = $paths.TenantId
}
foreach ($row in $rawServicePlanRows) {
    foreach ($identityName in $identityFields.Keys) {
        if ([string]$row.$identityName -ne [string]$identityFields[$identityName]) {
            throw "Raw service-plan identity mismatch for '$identityName'."
        }
    }
}
$duplicateServicePlanKeys = @($rawServicePlanRows | Group-Object {
    '{0}|{1}' -f ([string]$_.SkuId).ToLowerInvariant(),([string]$_.ServicePlanId).ToLowerInvariant()
} | Where-Object Count -gt 1)
if ($duplicateServicePlanKeys.Count -gt 0) {
    throw "Raw service-plan data contains duplicate SKU/plan keys: $($duplicateServicePlanKeys.Name -join ', ')"
}
foreach ($row in $rawRows) {
    foreach ($identityName in $identityFields.Keys) {
        if ([string]$row.$identityName -ne [string]$identityFields[$identityName]) {
            throw "Raw subscribed SKUs identity mismatch for '$identityName'."
        }
    }
    foreach ($keyName in @('SubscribedSkuId', 'SkuId', 'SkuPartNumber')) {
        if ([string]::IsNullOrWhiteSpace([string]$row.$keyName)) {
            throw "Raw subscribed SKUs data contains an empty $keyName."
        }
    }
}
$duplicateIds = @($rawRows | Group-Object SubscribedSkuId | Where-Object Count -gt 1)
if ($duplicateIds.Count -gt 0) {
    throw "Raw subscribed SKUs data contains duplicate SubscribedSkuId values: $($duplicateIds.Name -join ', ')"
}

$cmdbRows = @($rawRows | Group-Object {
        ([string]$_.SkuId).Trim().ToLowerInvariant()
    } | ForEach-Object {
        $groupRows = @($_.Group)
        $skuPartNumbers = @($groupRows | Select-Object -ExpandProperty SkuPartNumber -Unique)
        if ($skuPartNumbers.Count -ne 1) {
            throw "SKU '$($_.Name)' has inconsistent skuPartNumber values."
        }
        $consumed = [long]0; $enabled = [long]0; $suspended = [long]0; $warning = [long]0
        foreach ($row in $groupRows) {
            $consumed += ConvertTo-SmartWorkplaceCMDBLicenseInteger ([string]$row.ConsumedUnits) 'ConsumedUnits' ([string]$row.SubscribedSkuId)
            $enabled += ConvertTo-SmartWorkplaceCMDBLicenseInteger ([string]$row.EnabledUnits) 'EnabledUnits' ([string]$row.SubscribedSkuId)
            $suspended += ConvertTo-SmartWorkplaceCMDBLicenseInteger ([string]$row.SuspendedUnits) 'SuspendedUnits' ([string]$row.SubscribedSkuId)
            $warning += ConvertTo-SmartWorkplaceCMDBLicenseInteger ([string]$row.WarningUnits) 'WarningUnits' ([string]$row.SubscribedSkuId)
        }
        $latestRow = @($groupRows | Sort-Object `
            @{ Expression = { Get-SmartWorkplaceCMDBLicenseDateValue ([string]$_.SourceCollectedDateTime) ([string]$_.SubscribedSkuId) }; Descending = $true },
            @{ Expression = { [string]$_.SubscribedSkuId }; Descending = $false })[0]
        [pscustomobject][ordered]@{
            CmdbLicenseId = ('{0}|sku|{1}' -f $paths.TenantKey, [string]$_.Name)
            SourceSystem = 'MicrosoftGraph'
            SkuId = [string]$latestRow.SkuId
            SkuPartNumber = [string]$skuPartNumbers[0]
            ConsumedUnits = $consumed
            EnabledUnits = $enabled
            SuspendedUnits = $suspended
            WarningUnits = $warning
            SourceCollectedDateTime = [string]$latestRow.SourceCollectedDateTime
        }
    } | Sort-Object SkuPartNumber, SkuId)
$dimRows = @($cmdbRows | ForEach-Object {
    [pscustomobject][ordered]@{
        TenantSkuKey = $_.CmdbLicenseId; SkuId = $_.SkuId
        SkuPartNumber = $_.SkuPartNumber; ConsumedUnits = $_.ConsumedUnits
        EnabledUnits = $_.EnabledUnits
    }
})
$dimServicePlanRows = @($rawServicePlanRows | Sort-Object SkuPartNumber,ServicePlanName,ServicePlanId | ForEach-Object {
    $skuId = ([string]$_.SkuId).Trim()
    $servicePlanId = ([string]$_.ServicePlanId).Trim()
    if ([string]::IsNullOrWhiteSpace($skuId) -or [string]::IsNullOrWhiteSpace($servicePlanId)) {
        throw 'Raw service-plan data contains an empty SkuId or ServicePlanId.'
    }
    [pscustomobject][ordered]@{
        TenantServicePlanKey = ('{0}|service-plan|{1}|{2}' -f $paths.TenantKey,$skuId.ToLowerInvariant(),$servicePlanId.ToLowerInvariant())
        TenantSkuKey = ('{0}|sku|{1}' -f $paths.TenantKey,$skuId.ToLowerInvariant())
        SkuId = $skuId
        SkuPartNumber = [string]$_.SkuPartNumber
        ServicePlanId = $servicePlanId
        ServicePlanName = [string]$_.ServicePlanName
        ProvisioningStatus = [string]$_.ProvisioningStatus
        AppliesTo = [string]$_.AppliesTo
    }
})
$identityExport = @{
    TenantKey = $paths.TenantKey; OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey; TenantId = $paths.TenantId
}
Export-SmartWorkplaceCMDBCsv -InputObject $cmdbRows -Path $cmdbOutputPath `
    -Columns @($cmdbTable.columns | ForEach-Object { [string]$_ }) @identityExport
Export-SmartWorkplaceCMDBCsv -InputObject $dimRows -Path $dimOutputPath `
    -Columns @($dimTable.columns | ForEach-Object { [string]$_ }) @identityExport
Export-SmartWorkplaceCMDBCsv -InputObject $dimServicePlanRows -Path $dimServicePlanOutputPath `
    -Columns @($dimServicePlanTable.columns | ForEach-Object { [string]$_ }) @identityExport

$cmdbValidation = Test-SmartWorkplaceCMDBExactCsvHeader $cmdbOutputPath @($cmdbTable.columns | ForEach-Object { [string]$_ })
$dimValidation = Test-SmartWorkplaceCMDBExactCsvHeader $dimOutputPath @($dimTable.columns | ForEach-Object { [string]$_ })
$dimServicePlanValidation = Test-SmartWorkplaceCMDBExactCsvHeader $dimServicePlanOutputPath @($dimServicePlanTable.columns | ForEach-Object { [string]$_ })
if ($cmdbValidation.Status -ne 'Valid' -or $dimValidation.Status -ne 'Valid' -or $dimServicePlanValidation.Status -ne 'Valid') {
    throw 'Normalized license outputs did not satisfy the curated contracts.'
}
Write-Information (
    "SmartWorkplaceCMDB Microsoft 365 subscribed SKUs normalization completed. RawSKUs={0}; CMDBLicenses={1}; ServicePlans={2}." -f
    $rawRows.Count, $cmdbRows.Count, $dimServicePlanRows.Count
) -InformationAction Continue
[pscustomobject]@{
    Status = 'Completed'; ScriptVersion = $ScriptVersion
    RawSkuCount = $rawRows.Count; LicenseCount = $cmdbRows.Count
    RawInputPath = $RawInputPath; CmdbLicenseOutputPath = $cmdbOutputPath
    DimLicenseSkuOutputPath = $dimOutputPath
    ServicePlanCount = $dimServicePlanRows.Count
    DimLicenseServicePlanOutputPath = $dimServicePlanOutputPath
    RawContractVersion = [string]$rawContract.contractVersion
    CuratedContractVersion = [string]$curatedContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCDbziGcR5lk3kP
# AHhrnzPiwCmBTuZaj5prK3IR3IeQMKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDhSiGeNxwUSdTFQMMuxiXePrlye1MCJzYPRbEqMd3yzMA0GCSqG
# SIb3DQEBAQUABIIBgFU0jq+bv/e9uQ5w0kWt3LswcjOsF7WzvQVuwG5UfiKMMN33
# GxEX2Hsw2KAu5bdtguUYNmQpPokQdrhUmL+n6Pi5iT6vNqIcV1vCjEN7ymUtj3+o
# Gt5lw9ScWfClyTffH81/nZUWO/q0aKcNLu7wKzxqBdgDrTzb10nB1h6c8sfvdWI7
# swiGP815sLtQB9BoTupVGaw+D2DDb+u2/RQIYoGeTrSJ+1GQLf2G0VoUVcDlGYEM
# mLB8SsgMLaKvCsj89Vpm+isFmZ2BkgiK+frP+YKdJE/TzHshcvo2pvDjaFF4cdho
# uV/S5Zh0r2tvEAFDsAgFkrF5JoLSkv69YIWS6r9Z45nwv/o4S53i99q/NxOV0zVY
# HvaU4B1IVPJDNeruIqhcgKgRnzml4X0oCYkzX+qCmRRP0ssi/N31xSxVhwoMJ3Tg
# Umy8JADy4oGY2eROFe9yYJD2o/q7XyT6dZD0qBFFwvVISRFuPh0YQTokD6qK60/u
# e1zj/Y2KMBVuPfs0J6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxNzQ0
# MjhaMC8GCSqGSIb3DQEJBDEiBCDtJzA5mnOhVcQkUpt+dNfi1ukdFkC6kB9ZTg/X
# Pv9aQDANBgkqhkiG9w0BAQEFAASCAgCHZOMi/JqCJkd+iDsIYuEVGBC3aeBI0rGx
# qfD87e5uZ9afYSKDV8VY1AbnL+TUkLxV3ilSA3DVhCGxfPY+K6PTpUxz+BWvDIZq
# 8s1c2BPyiH7uJ5YhoiH4SXJpYwSrIdz5+d3CxpOO974k8Q/fACWJfG1ZzwZ7j3nA
# u+vwwfDxfPLcmTz9U0qjzJqsNKEfcsvl3Gw478fD8pYFPUrxcVLsbPw1bIz3yVyn
# kCWk2wHhIq4cMzcS9HtHosw/YsNC/gOXBeYPiZ1RhWWEnlg3oAvh5wEHCU3u14VA
# d61M3DQu/1R7qEWTM8yyj58qUdsuvx2hUA1KwynB6HsKvFqfaxWAbD6jhAIIC7Zq
# hFhN4bv847aS0iucx+FKVTXS1g66noT9KMGGhlYqgrz5MnYh/T6awgyJ5c2siysj
# pXA4giB+X6iwg19E8xRDXIMU/biD7izJmWDeBqN5yv0sxgpLna76qWfml7xutVDe
# rmDZHA480ATNNKcgawWwjaUajKl4TV3Ad1zisoNz/Uzp9nK4nOaIa9J3cNzxuDW5
# WSMK/MepeyC82JLuDtETFXbFa02vJtktcZH1envSuoxRrMJb+P0pVYEhObrAvT/P
# FyNNUVRST2kryYysXBEWQwW5muSWg6aMdDCCF9vgjoRuP8Cb0N7NNaGWAn2Hr1rO
# IeXlxpaZ6A==
# SIG # End signature block
