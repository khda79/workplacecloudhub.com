<#
.SYNOPSIS
Normalizes Microsoft 365 user license assignments for Power BI.

.DESCRIPTION
Publishes one FactUserLicense.csv row per user and SKU and one compact
FactUserServicePlan.csv row per effective user/SKU/service-plan path. When Graph returns
multiple direct or group-derived states for the same user and SKU, the most
critical state is selected first (Error, ActiveWithError, Disabled, Active),
then the most recently updated state. AssignedDateTime contains Graph
lastUpdatedDateTime because the original assignment creation time is not
available on licenseAssignmentState.

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
    [string]$CmdbUsersPath,
    [string]$DimLicenseSkuPath,
    [string]$ServicePlansPath,
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
$servicePlanTable = @($rawContract.tables | Where-Object name -eq 'M365_ServicePlans.csv')
$servicePlanFactTable = @($curatedContract.tables | Where-Object name -eq 'FactUserServicePlan.csv')
if ($rawTable.Count -ne 1 -or $factTable.Count -ne 1 -or $servicePlanTable.Count -ne 1 -or $servicePlanFactTable.Count -ne 1) {
    throw 'The contracts must contain user license and service-plan source and fact definitions.'
}
$rawTable = $rawTable[0]; $factTable = $factTable[0]
$servicePlanTable = $servicePlanTable[0]; $servicePlanFactTable = $servicePlanFactTable[0]
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
if ([string]::IsNullOrWhiteSpace($ServicePlansPath)) {
    $ServicePlansPath = Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$servicePlanTable.area) ([string]$servicePlanTable.name))
}
$RawInputPath = [IO.Path]::GetFullPath($RawInputPath)
$CmdbUsersPath = [IO.Path]::GetFullPath($CmdbUsersPath)
$DimLicenseSkuPath = [IO.Path]::GetFullPath($DimLicenseSkuPath)
$ServicePlansPath = [IO.Path]::GetFullPath($ServicePlansPath)
$rawHeader = Test-SmartWorkplaceCMDBExactCsvHeader $RawInputPath @($rawTable.columns | ForEach-Object { [string]$_ })
if ($rawHeader.Status -eq 'Incompatible') { throw 'Raw user license assignments CSV is incompatible.' }
if ($rawHeader.Status -eq 'Missing' -and -not $ValidateOnly) {
    throw "Raw user license assignments CSV was not found: $RawInputPath"
}
$factOutputPath = Join-Path $paths.PowerBILatestPath ([string]$factTable.name)
$servicePlanFactOutputPath = Join-Path $paths.PowerBILatestPath ([string]$servicePlanFactTable.name)
$factHeader = Test-SmartWorkplaceCMDBExactCsvHeader $factOutputPath @($factTable.columns | ForEach-Object { [string]$_ })
$servicePlanFactHeader = Test-SmartWorkplaceCMDBExactCsvHeader $servicePlanFactOutputPath @($servicePlanFactTable.columns | ForEach-Object { [string]$_ })
if ($factHeader.Status -eq 'Incompatible' -or $servicePlanFactHeader.Status -eq 'Incompatible') {
    throw 'An existing curated user-license or service-plan fact table is incompatible.'
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
        FactUserLicenseTargetStatus = $factHeader.Status
        FactUserLicenseOutputPath = $factOutputPath
        FactUserServicePlanTargetStatus = $servicePlanFactHeader.Status
        FactUserServicePlanOutputPath = $servicePlanFactOutputPath
        CmdbUsersReferenceExists = Test-Path -LiteralPath $CmdbUsersPath -PathType Leaf
        DimLicenseSkuReferenceExists = Test-Path -LiteralPath $DimLicenseSkuPath -PathType Leaf
    } | Format-List
    return
}

$rawRows = @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $RawInputPath -Paths $paths -ErrorAction Stop)
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

$selectedAssignments = @($rawRows | Group-Object {
        '{0}|{1}' -f ([string]$_.SourceUserId).Trim().ToLowerInvariant(), ([string]$_.SkuId).Trim().ToLowerInvariant()
    } | ForEach-Object {
        @($_.Group | Sort-Object `
            @{ Expression = { Get-SmartWorkplaceCMDBAssignmentStatePriority ([string]$_.AssignmentState) }; Descending = $true },
            @{ Expression = { Get-SmartWorkplaceCMDBAssignmentDateValue ([string]$_.LastUpdatedDateTime) ([string]$_.RawAssignmentKey) }; Descending = $true },
            @{ Expression = { [string]$_.RawAssignmentKey }; Descending = $false })[0]
    })
$factRows = @($selectedAssignments | ForEach-Object {
        $selected = $_
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

$servicePlansBySku = @{}
if (Test-Path -LiteralPath $ServicePlansPath -PathType Leaf) {
    foreach ($plan in @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $ServicePlansPath -Paths $paths -ErrorAction Stop)) {
        $skuKey = ([string]$plan.SkuId).Trim().ToLowerInvariant()
        if (-not $servicePlansBySku.ContainsKey($skuKey)) { $servicePlansBySku[$skuKey] = New-Object 'Collections.Generic.List[object]' }
        $servicePlansBySku[$skuKey].Add($plan)
    }
}
$servicePlanFactRows = @($selectedAssignments | ForEach-Object {
    $selected = $_
    $sourceUserId = ([string]$selected.SourceUserId).Trim()
    $skuId = ([string]$selected.SkuId).Trim()
    $skuKey = $skuId.ToLowerInvariant()
    $disabledPlans = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($disabledPlanId in @(([string]$selected.DisabledPlanIds) -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        [void]$disabledPlans.Add($disabledPlanId.Trim())
    }
    if ($servicePlansBySku.ContainsKey($skuKey)) {
        foreach ($plan in $servicePlansBySku[$skuKey]) {
            $servicePlanId = ([string]$plan.ServicePlanId).Trim()
            [pscustomobject][ordered]@{
                TenantUserKey = ('{0}|entra-user|{1}' -f $paths.TenantKey,$sourceUserId.ToLowerInvariant())
                TenantSkuKey = ('{0}|sku|{1}' -f $paths.TenantKey,$skuKey)
                TenantServicePlanKey = ('{0}|service-plan|{1}|{2}' -f $paths.TenantKey,$skuKey,$servicePlanId.ToLowerInvariant())
                CmdbUserId = ('{0}|entra-user|{1}' -f $paths.TenantKey,$sourceUserId.ToLowerInvariant())
                SkuId = $skuId
                ServicePlanId = $servicePlanId
                ServicePlanName = [string]$plan.ServicePlanName
                IsEnabled = -not $disabledPlans.Contains($servicePlanId)
                AssignmentState = [string]$selected.AssignmentState
                SourceSystem = 'MicrosoftEntraID'
            }
        }
    }
} | Sort-Object TenantUserKey,TenantSkuKey,TenantServicePlanKey)

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
Export-SmartWorkplaceCMDBCsv -InputObject $servicePlanFactRows -Path $servicePlanFactOutputPath `
    -Columns @($servicePlanFactTable.columns | ForEach-Object { [string]$_ }) @identityExport
$factValidation = Test-SmartWorkplaceCMDBExactCsvHeader $factOutputPath @($factTable.columns | ForEach-Object { [string]$_ })
$servicePlanFactValidation = Test-SmartWorkplaceCMDBExactCsvHeader $servicePlanFactOutputPath @($servicePlanFactTable.columns | ForEach-Object { [string]$_ })
if ($factValidation.Status -ne 'Valid' -or $servicePlanFactValidation.Status -ne 'Valid') {
    throw 'Normalized user-license outputs did not satisfy the curated contracts.'
}

Write-Information (
    "SmartWorkplaceCMDB user license assignments normalization completed. Raw={0}; Facts={1}; ServicePlanFacts={2}; Collapsed={3}; OrphanUsers={4}; OrphanSKUs={5}." -f
    $rawRows.Count, $factRows.Count, $servicePlanFactRows.Count, ($rawRows.Count - $factRows.Count), $orphanUserCount, $orphanSkuCount
) -InformationAction Continue
[pscustomobject]@{
    Status = 'Completed'; ScriptVersion = $ScriptVersion
    RawAssignmentCount = $rawRows.Count; FactAssignmentCount = $factRows.Count
    CollapsedAssignmentCount = $rawRows.Count - $factRows.Count
    DirectAssignmentCount = @($rawRows | Where-Object { [string]::IsNullOrWhiteSpace($_.AssignedByGroupId) }).Count
    GroupAssignmentCount = @($rawRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.AssignedByGroupId) }).Count
    OrphanUserCount = $orphanUserCount; OrphanSkuCount = $orphanSkuCount
    RawInputPath = $RawInputPath; FactUserLicenseOutputPath = $factOutputPath
    ServicePlanFactCount = $servicePlanFactRows.Count
    FactUserServicePlanOutputPath = $servicePlanFactOutputPath
    RawContractVersion = [string]$rawContract.contractVersion
    CuratedContractVersion = [string]$curatedContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBgkXBnQAo67Eaj
# jYkTXEIlQ1Hh6+bUq6nBz6y2so8JUKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII9ur7gJWiL4yH0U+3Vl9Dyj9IVo8qrF5AYvnvQfS60BMA0GCSqG
# SIb3DQEBAQUABIIBgBDTduCVqLXmW/KD44fGBTzczribDGssXQAPgNh6sU+3B1Sw
# E4MWGr7gCAK68jPPZLFd7HleYvnhNVCvPxxdmIfErlzE6bKcxGAZI3n775i6MMAp
# Wjmsob/Mh3OIGEefDXPbdiksIOZQhgA7hn8Q1ibYHQi5JlmUYTbRyS0K9wJyrfBg
# NgUeKw7tC7EMTypvRQmV5L3WhsEVFcZhy9F/H5v9JABzxedovEe24bayjvCvhbL3
# 5il/OJFkvEspHGvLkQyuGSmKjrtfGZZ9tS3Wc6wAXj+YX1P8ChFcFsqK5hINYASt
# MGfJdmMiocZd53jafZ2Ow/m0Gz9X8656ojagmreMfmVW/HySosm7lG5ujBoFzEsY
# e6O5gB9YT88lensJdz9zBB1cHQRa7l1Jsvcil+DTyZ2Z0MjkU0Pg+LmxERAGGguy
# 4Ngk/9cNHPYDbHksPrKpSdc3Dyuim8+B7B72NJdwDayunCL+9kf1TSsWvy0P+Q9y
# tZQalYhH5/UbWLywoaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxNzQ0
# MjhaMC8GCSqGSIb3DQEJBDEiBCDCip02Y5O9fCM7I5R2JwzTkF6OhfSp06XwE91W
# rQ4exjANBgkqhkiG9w0BAQEFAASCAgBKTI68yJl6GU9tog71Xjyj+ZJzhMlbsfkk
# ypIwhqzSq2QNM0eP74j0S+jIAP3KUqOqt2EYTJ9oFiJ6+vyTqkOM2J2DbsB2pYr7
# T2Rc6znS0OUKSvtnJ9GfSOHClDMwHxXV4wIxR2kFJhZnS6FlTWWHMAHg2k/fhn2R
# 3rn5encZWwdyPA5d0Ao0FoEn+fRp1QAn6Sm3xESNClQ2XAq3R2NXGHKlCrHvLe1+
# Gv0Bm7BqYSwZDuZTeJ3UQGL62olMsyXBWFTOw0rgNblNEguKm6kVnZNjh53wxr9/
# hBNJ74P2ELzUJyBgwO/5W58NLqYi9I9DLu5bJ5KNbMmWTKHvAPqwflHjwhHPDHMt
# 6yuGdSYr3V1ozoaW4PB7TNbq1T2GXKE7lFK+iLOH5lgOQ912NVOZU9eheoWrobZc
# NKexZxQV9KTl8IzZW3sApxZCgaAVXHWDcbbe8AUbEMct1q/5e5cIn7GPEQiioF6q
# HhTuK10Z1JNuP196NZRoXqzVCX+2T+UCyPA9TkIxTYXn7jxy0IrjIb4LQfRGVQrQ
# c8TVznfs1VX0b0bEk5qSqjGUTmG5J/LJu98Dqkg8ITij3L54hL4TDLsTUAQ7Gk21
# 0ignAcZELIp8wVdsZ8+NiOWTW26xQcnDA3FiGcvGLxjd9yS+hULza1b63Poz9zvO
# E3aSjxZ9Jw==
# SIG # End signature block
