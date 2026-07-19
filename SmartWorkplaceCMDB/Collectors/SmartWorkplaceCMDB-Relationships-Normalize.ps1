<#
.SYNOPSIS
Consolidates SmartWorkplaceCMDB entity relationships.

.DESCRIPTION
Publishes CMDB_Relationships.csv from validated primary user-device links,
mailbox facts, and user-license facts. The normalizer performs no tenant
connection and rejects broken references in published relationships.

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
    [string]$LicenseInputPath,
    [string]$MailboxInputPath,
    [string]$UserDeviceInputPath,
    [string]$MailboxFactInputPath,
    [string]$UserLicenseFactInputPath,
    [string]$RawUserLicenseInputPath,
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

function ConvertTo-SmartWorkplaceCMDBUtcDateText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Value,
        [Parameter(Mandatory)][string]$Context
    )
    $parsed = [datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse(
            $Value,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal,
            [ref]$parsed
        )) {
        throw "Invalid date '$Value' for $Context."
    }
    return $parsed.ToUniversalTime().ToString(
        'yyyy-MM-ddTHH:mm:ss.fffffffZ',
        [Globalization.CultureInfo]::InvariantCulture
    )
}

function Get-SmartWorkplaceCMDBDefaultConfidenceText {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Configuration)
    $value = '0.5'
    if ($Configuration.Contains('DataQuality') -and
        $Configuration['DataQuality'] -is [System.Collections.IDictionary] -and
        $Configuration['DataQuality'].Contains('DefaultConfidenceScore')) {
        $value = [string]$Configuration['DataQuality']['DefaultConfidenceScore']
    }
    $parsed = [double]0
    if (-not [double]::TryParse(
            $value,
            [Globalization.NumberStyles]::Float,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$parsed
        ) -or
        $parsed -lt 0 -or
        $parsed -gt 1) {
        throw "DataQuality.DefaultConfidenceScore '$value' must be between 0 and 1."
    }
    return $parsed.ToString(
        '0.################',
        [Globalization.CultureInfo]::InvariantCulture
    )
}

function Add-SmartWorkplaceCMDBUniqueLookup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Lookup,
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)]$Value,
        [Parameter(Mandatory)][string]$Context
    )
    if ([string]::IsNullOrWhiteSpace($Key)) {
        throw "$Context contains an empty key."
    }
    if ($Lookup.ContainsKey($Key)) {
        throw "$Context contains duplicate key '$Key'."
    }
    $Lookup[$Key] = $Value
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$contractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
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
$rawContract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath

$tableNames = @(
    'CMDB_Users.csv',
    'CMDB_Devices.csv',
    'CMDB_Licenses.csv',
    'CMDB_Mailboxes.csv',
    'CMDB_UserDeviceRelationships.csv',
    'FactMailbox.csv',
    'FactUserLicense.csv',
    'CMDB_Relationships.csv'
)
$tables = @{}
foreach ($name in $tableNames) {
    $match = @($contract.tables | Where-Object name -eq $name)
    if ($match.Count -ne 1) {
        throw "The curated contract must contain exactly one '$name' table."
    }
    $tables[$name] = $match[0]
}
$rawAssignmentTable = @($rawContract.tables |
    Where-Object name -eq 'M365_UserLicenseAssignments.csv')
if ($rawAssignmentTable.Count -ne 1) {
    throw 'The raw contract must contain one M365_UserLicenseAssignments.csv table.'
}
$rawAssignmentTable = $rawAssignmentTable[0]

$defaultInputs = [ordered]@{
    UserInputPath = Join-Path $paths.CmdbLatestPath 'CMDB_Users.csv'
    DeviceInputPath = Join-Path $paths.CmdbLatestPath 'CMDB_Devices.csv'
    LicenseInputPath = Join-Path $paths.CmdbLatestPath 'CMDB_Licenses.csv'
    MailboxInputPath = Join-Path $paths.CmdbLatestPath 'CMDB_Mailboxes.csv'
    UserDeviceInputPath = Join-Path $paths.CmdbLatestPath 'CMDB_UserDeviceRelationships.csv'
    MailboxFactInputPath = Join-Path $paths.PowerBILatestPath 'FactMailbox.csv'
    UserLicenseFactInputPath = Join-Path $paths.PowerBILatestPath 'FactUserLicense.csv'
    RawUserLicenseInputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$rawAssignmentTable.area) ([string]$rawAssignmentTable.name)
    )
}
foreach ($parameterName in $defaultInputs.Keys) {
    $currentValue = Get-Variable -Name $parameterName -ValueOnly
    if ([string]::IsNullOrWhiteSpace([string]$currentValue)) {
        Set-Variable -Name $parameterName -Value $defaultInputs[$parameterName]
    }
}

$inputDefinitions = @(
    [pscustomobject]@{ Name = 'CMDB_Users.csv'; Path = [IO.Path]::GetFullPath($UserInputPath); Table = $tables['CMDB_Users.csv'] },
    [pscustomobject]@{ Name = 'CMDB_Devices.csv'; Path = [IO.Path]::GetFullPath($DeviceInputPath); Table = $tables['CMDB_Devices.csv'] },
    [pscustomobject]@{ Name = 'CMDB_Licenses.csv'; Path = [IO.Path]::GetFullPath($LicenseInputPath); Table = $tables['CMDB_Licenses.csv'] },
    [pscustomobject]@{ Name = 'CMDB_Mailboxes.csv'; Path = [IO.Path]::GetFullPath($MailboxInputPath); Table = $tables['CMDB_Mailboxes.csv'] },
    [pscustomobject]@{ Name = 'CMDB_UserDeviceRelationships.csv'; Path = [IO.Path]::GetFullPath($UserDeviceInputPath); Table = $tables['CMDB_UserDeviceRelationships.csv'] },
    [pscustomobject]@{ Name = 'FactMailbox.csv'; Path = [IO.Path]::GetFullPath($MailboxFactInputPath); Table = $tables['FactMailbox.csv'] },
    [pscustomobject]@{ Name = 'FactUserLicense.csv'; Path = [IO.Path]::GetFullPath($UserLicenseFactInputPath); Table = $tables['FactUserLicense.csv'] },
    [pscustomobject]@{ Name = 'M365_UserLicenseAssignments.csv'; Path = [IO.Path]::GetFullPath($RawUserLicenseInputPath); Table = $rawAssignmentTable }
)
foreach ($definition in $inputDefinitions) {
    $definition | Add-Member -NotePropertyName Header -NotePropertyValue (
        Test-SmartWorkplaceCMDBExactCsvHeader `
            -Path $definition.Path `
            -ExpectedColumns @($definition.Table.columns | ForEach-Object { [string]$_ })
    )
    if ($definition.Header.Status -eq 'Incompatible') {
        throw "Relationship input '$($definition.Name)' is incompatible."
    }
    if ($definition.Header.Status -eq 'Missing' -and -not $ValidateOnly) {
        throw "Relationship input '$($definition.Name)' was not found: $($definition.Path)"
    }
}

$targetTable = $tables['CMDB_Relationships.csv']
$outputPath = Join-Path $paths.CmdbLatestPath ([string]$targetTable.name)
$targetHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $outputPath `
    -ExpectedColumns @($targetTable.columns | ForEach-Object { [string]$_ })
if ($targetHeader.Status -eq 'Incompatible') {
    throw "Existing curated table '$($targetTable.name)' is incompatible."
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status              = 'Valid'
        ScriptVersion       = $ScriptVersion
        ContractVersion     = [string]$contract.contractVersion
        RawContractVersion  = [string]$rawContract.contractVersion
        ValidInputCount     = @($inputDefinitions |
            Where-Object { $_.Header.Status -eq 'Valid' }).Count
        MissingInputCount   = @($inputDefinitions |
            Where-Object { $_.Header.Status -eq 'Missing' }).Count
        TargetStatus        = $targetHeader.Status
        OutputPath          = $outputPath
        TenantKey           = $paths.TenantKey
    } | Format-List
    return
}

$inputRows = @{}
foreach ($definition in $inputDefinitions) {
    $inputRows[$definition.Name] = @(
        Import-Csv -LiteralPath $definition.Path -ErrorAction Stop
    )
}
$identityFields = [ordered]@{
    TenantKey       = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey  = $paths.EnvironmentKey
    TenantId        = $paths.TenantId
}
foreach ($definition in $inputDefinitions) {
    foreach ($row in @($inputRows[$definition.Name])) {
        foreach ($identityName in $identityFields.Keys) {
            if ([string]$row.$identityName -ne
                [string]$identityFields[$identityName]) {
                throw "Relationship input '$($definition.Name)' identity mismatch for '$identityName'."
            }
        }
    }
}

$usersById = @{}
foreach ($row in @($inputRows['CMDB_Users.csv'])) {
    Add-SmartWorkplaceCMDBUniqueLookup `
        -Lookup $usersById `
        -Key ([string]$row.CmdbUserId) `
        -Value $row `
        -Context 'CMDB users'
}
$devicesById = @{}
foreach ($row in @($inputRows['CMDB_Devices.csv'])) {
    Add-SmartWorkplaceCMDBUniqueLookup `
        -Lookup $devicesById `
        -Key ([string]$row.CmdbDeviceId) `
        -Value $row `
        -Context 'CMDB devices'
}
$licensesById = @{}
foreach ($row in @($inputRows['CMDB_Licenses.csv'])) {
    Add-SmartWorkplaceCMDBUniqueLookup `
        -Lookup $licensesById `
        -Key ([string]$row.CmdbLicenseId) `
        -Value $row `
        -Context 'CMDB licenses'
}
$mailboxesById = @{}
foreach ($row in @($inputRows['CMDB_Mailboxes.csv'])) {
    Add-SmartWorkplaceCMDBUniqueLookup `
        -Lookup $mailboxesById `
        -Key ([string]$row.CmdbMailboxId) `
        -Value $row `
        -Context 'CMDB mailboxes'
}

$assignmentCollectedByUserSku = @{}
$rawAssignmentIds = @{}
foreach ($row in @($inputRows['M365_UserLicenseAssignments.csv'])) {
    Add-SmartWorkplaceCMDBUniqueLookup `
        -Lookup $rawAssignmentIds `
        -Key ([string]$row.RawAssignmentKey) `
        -Value $row `
        -Context 'Raw user license assignments'
    $sourceUserId = ([string]$row.SourceUserId).Trim().ToLowerInvariant()
    $skuId = ([string]$row.SkuId).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($sourceUserId) -or
        [string]::IsNullOrWhiteSpace($skuId)) {
        throw 'Raw user license assignments contains an empty user or SKU ID.'
    }
    $key = '{0}{1}{2}' -f $sourceUserId, [char]31, $skuId
    $dateText = ConvertTo-SmartWorkplaceCMDBUtcDateText `
        -Value ([string]$row.SourceCollectedDateTime) `
        -Context "raw assignment '$($row.RawAssignmentKey)'"
    if (-not $assignmentCollectedByUserSku.ContainsKey($key) -or
        [datetimeoffset]$dateText -gt
        [datetimeoffset]$assignmentCollectedByUserSku[$key]) {
        $assignmentCollectedByUserSku[$key] = $dateText
    }
}

$defaultConfidence = Get-SmartWorkplaceCMDBDefaultConfidenceText `
    -Configuration $context.Configuration
$rows = New-Object System.Collections.Generic.List[object]
$userDeviceCount = 0
foreach ($relationship in @($inputRows['CMDB_UserDeviceRelationships.csv'])) {
    if (-not $usersById.ContainsKey([string]$relationship.CmdbUserId)) {
        throw "User-device relationship references an unknown CMDB user."
    }
    if (-not $devicesById.ContainsKey([string]$relationship.CmdbDeviceId)) {
        throw "User-device relationship references an unknown CMDB device."
    }
    $rows.Add([pscustomobject][ordered]@{
        CmdbRelationshipId      = [string]$relationship.CmdbRelationshipId
        FromEntityType          = 'User'
        FromEntityId            = [string]$relationship.CmdbUserId
        ToEntityType            = 'Device'
        ToEntityId              = [string]$relationship.CmdbDeviceId
        RelationshipType        = [string]$relationship.RelationshipType
        SourceSystem            = [string]$relationship.SourceSystem
        ConfidenceScore         = [string]$relationship.ConfidenceScore
        SourceCollectedDateTime = ConvertTo-SmartWorkplaceCMDBUtcDateText `
            -Value ([string]$relationship.SourceCollectedDateTime) `
            -Context "user-device relationship '$($relationship.CmdbRelationshipId)'"
    })
    $userDeviceCount++
}

$mailboxCount = 0
$unlinkedMailboxCount = 0
foreach ($fact in @($inputRows['FactMailbox.csv'])) {
    if ([string]::IsNullOrWhiteSpace([string]$fact.CmdbUserId)) {
        $unlinkedMailboxCount++
        continue
    }
    if (-not $usersById.ContainsKey([string]$fact.CmdbUserId)) {
        throw 'Mailbox fact references an unknown CMDB user.'
    }
    if (-not $mailboxesById.ContainsKey([string]$fact.CmdbMailboxId)) {
        throw 'Mailbox fact references an unknown CMDB mailbox.'
    }
    $mailbox = $mailboxesById[[string]$fact.CmdbMailboxId]
    $rows.Add([pscustomobject][ordered]@{
        CmdbRelationshipId      = ('{0}|relationship|user-mailbox|{1}' -f
            $paths.TenantKey,
            ([string]$fact.CmdbMailboxId).ToLowerInvariant())
        FromEntityType          = 'User'
        FromEntityId            = [string]$fact.CmdbUserId
        ToEntityType            = 'Mailbox'
        ToEntityId              = [string]$fact.CmdbMailboxId
        RelationshipType        = 'HasMailbox'
        SourceSystem            = [string]$fact.SourceSystem
        ConfidenceScore         = $defaultConfidence
        SourceCollectedDateTime = ConvertTo-SmartWorkplaceCMDBUtcDateText `
            -Value ([string]$mailbox.SourceCollectedDateTime) `
            -Context "mailbox '$($fact.CmdbMailboxId)'"
    })
    $mailboxCount++
}

$licenseCount = 0
foreach ($fact in @($inputRows['FactUserLicense.csv'])) {
    if (-not $usersById.ContainsKey([string]$fact.CmdbUserId)) {
        throw 'User-license fact references an unknown CMDB user.'
    }
    if (-not $licensesById.ContainsKey([string]$fact.TenantSkuKey)) {
        throw 'User-license fact references an unknown CMDB license.'
    }
    $user = $usersById[[string]$fact.CmdbUserId]
    $sourceUserId = ([string]$user.SourceUserId).Trim().ToLowerInvariant()
    $skuId = ([string]$fact.SkuId).Trim().ToLowerInvariant()
    $assignmentKey = '{0}{1}{2}' -f $sourceUserId, [char]31, $skuId
    if (-not $assignmentCollectedByUserSku.ContainsKey($assignmentKey)) {
        throw 'User-license fact has no matching raw assignment collection date.'
    }
    $rows.Add([pscustomobject][ordered]@{
        CmdbRelationshipId      = ('{0}|relationship|user-license|{1}|{2}' -f
            $paths.TenantKey,
            $sourceUserId,
            $skuId)
        FromEntityType          = 'User'
        FromEntityId            = [string]$fact.CmdbUserId
        ToEntityType            = 'License'
        ToEntityId              = [string]$fact.TenantSkuKey
        RelationshipType        = 'AssignedLicense'
        SourceSystem            = [string]$fact.SourceSystem
        ConfidenceScore         = $defaultConfidence
        SourceCollectedDateTime = $assignmentCollectedByUserSku[$assignmentKey]
    })
    $licenseCount++
}

$relationships = @($rows.ToArray() |
    Sort-Object RelationshipType, FromEntityId, ToEntityId)
$duplicateIds = @($relationships |
    Group-Object CmdbRelationshipId |
    Where-Object Count -gt 1)
if ($duplicateIds.Count -gt 0) {
    throw 'Relationship consolidation produced duplicate CMDB relationship IDs.'
}

$identityExport = @{
    TenantKey       = $paths.TenantKey
    OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey
    TenantId        = $paths.TenantId
}
Export-SmartWorkplaceCMDBCsv `
    -InputObject $relationships `
    -Path $outputPath `
    -Columns @($targetTable.columns | ForEach-Object { [string]$_ }) `
    @identityExport
$validation = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $outputPath `
    -ExpectedColumns @($targetTable.columns | ForEach-Object { [string]$_ })
if ($validation.Status -ne 'Valid') {
    throw 'CMDB relationships output does not satisfy its curated contract.'
}

Write-Information (
    "SmartWorkplaceCMDB relationship consolidation completed. UserDevice={0}; Mailbox={1}; License={2}; Total={3}." -f
    $userDeviceCount,
    $mailboxCount,
    $licenseCount,
    $relationships.Count
) -InformationAction Continue

[pscustomobject]@{
    Status                 = 'Completed'
    ScriptVersion          = $ScriptVersion
    UserDeviceCount        = $userDeviceCount
    MailboxCount           = $mailboxCount
    UnlinkedMailboxCount   = $unlinkedMailboxCount
    UserLicenseCount       = $licenseCount
    RelationshipCount      = $relationships.Count
    OutputPath             = $outputPath
    ContractVersion        = [string]$contract.contractVersion
    RawContractVersion     = [string]$rawContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDkGmbKfCYWdGdh
# kiRWHpuP8RqN6wRmVj8ftYNHhUizrKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEFJf/tdk09ZjgnzMcKKbL54tHXUbe+HtoKHsDY0pTcWMA0GCSqG
# SIb3DQEBAQUABIIBgGCe0IFoAhNWGnKhqO5husioAH6udhxqoysh8/zrThZ2Wgg1
# /0lD+O7En17h2oxKheRw88X1T6T3VW5aKgZx8g52HbughihClxsZpnZ/3MneSE0o
# gs1D1YBI/h/6OgYOQfDIh/ev8eFER+fCk1fK8REkVAKsmWAP9gonSickTGpvAu6R
# SE/F9Mdo+MZDsngbtK11K+UF+t5A+C4NoyqKCLC2rjQ/NbB1Kw9xt/FoUrcPt9yT
# ri8X1XVyq6TeMbV0T31xdHXfQ2ZN0Unt4sRExG34EsUm1Z8DuOf9gghkzCulnKgp
# kSRY+CAJTlfy4BI6BWFBAlHiDUhP+nDPh3Xx7h6l54+PDjD3J6e0dbIPyfe8rFxp
# dS1ekMPfGG5kA5O6l5LQl1iGNBhrvCynb2yQ7z/WcvQkBHiUsAwNy+lg7vEQGbya
# XuLxGpIJlIeFj7yagEmqGXfosSt/NVDLf9iWJac+teGWmqap8elYOGZ6G8NVbltL
# /qqlcRxv+XmDl71XIaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxOTQy
# MzFaMC8GCSqGSIb3DQEJBDEiBCDeSlUJSEyEDwPa7cGidZube5GMAC5ETuA0Qo2k
# 3wU97jANBgkqhkiG9w0BAQEFAASCAgB4VKxI/T+sBrtm6zaQp19i2qHactgmgE1G
# vQ6qgvrHZA+yjX2TDyMkkk3mBhfloxQWmR4NGy9LbmEzrAlJC6uQa6gW4UA3YzdG
# RChX4ydNxiPtV+oxjhKrWDXwNUiJ1NL7JLGUu7AiLFW3g9FwkBOtwyHo5keHPEv0
# EX8FYqxjGRTzT5xf4E2kojY51qsAmjQntz0OPjSmEJyLEWYrllY16wbFhksmcORa
# tsVHdu0H6+qsPh/MU96HxSAdSsy/zSf6tLq7q5oLdQr2hopOJwL98lQ3F3xT6VWS
# 2ZmUgFMGbN/Xxm2x/UrUo3pTr7TQ6m/qfV5iFib9WeYuJXhwd3juGyaYU6EVf6H0
# QZZaHYYMPF7SlxAyD7giu8K4ZFpIgaAAnTsoYfSDny6qLgulxaZDWNzxScqsqN0B
# ldloWofOtLK0ibmxtE12pfETNf63waq4uVgNLtUi29XGuOtE2fesZdLZCWafT4Fq
# 06z7q0aUtTnw9B/fzrdxMgb4EmCT1r0PLUyMp1xqJmIlCTFrg0kff2gfg32BqdnN
# 6S2iDIPQRiIV73eWNkoQX6YjfXJEkr+NXFCbOAKr4c/LQdqpOlkZCpeDaWbEGIlV
# qfpJzX2IMJ9LMjehKaRHkSx1x7HCnT37glwqj+lfhIT4GX6AEo/UiF1PkkQ2p19g
# HEd/c2wAtg==
# SIG # End signature block
