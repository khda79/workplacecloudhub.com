<#
.SYNOPSIS
Publishes autonomous SmartWorkplaceCMDB data-quality findings.

.DESCRIPTION
Builds CMDB_DataQuality.csv and FactDataQuality.csv from curated local CSV
outputs. The normalizer performs no tenant connection. It reports orphan
primary-user references, unlinked mailboxes, missing or invalid collection
dates, duplicate entity keys, and stale non-empty entity datasets.

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
    [string]$GroupInputPath,
    [string]$DeviceInputPath,
    [string]$LicenseInputPath,
    [string]$MailboxInputPath,
    [string]$MailboxFactInputPath,
    [datetimeoffset]$ReferenceDateTime = [datetimeoffset]::UtcNow,
    [int]$FreshnessWarningHours = -1,
    [int]$FreshnessCriticalHours = -1,
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
    param([Parameter(Mandatory)][datetimeoffset]$Value)
    return $Value.ToUniversalTime().ToString(
        'yyyy-MM-ddTHH:mm:ss.fffffffZ',
        [Globalization.CultureInfo]::InvariantCulture
    )
}

function ConvertFrom-SmartWorkplaceCMDBDateText {
    [CmdletBinding()]
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }
    $parsed = [datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse(
            $Value,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal,
            [ref]$parsed
        )) {
        return $null
    }
    return $parsed.ToUniversalTime()
}

function Get-SmartWorkplaceCMDBPositiveIntegerSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Configuration,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$DefaultValue,
        [Parameter(Mandatory)][int]$OverrideValue
    )
    if ($OverrideValue -gt 0) {
        return $OverrideValue
    }
    $value = $DefaultValue
    if ($Configuration.Contains('DataQuality') -and
        $Configuration['DataQuality'] -is [System.Collections.IDictionary] -and
        $Configuration['DataQuality'].Contains($Name)) {
        $candidate = 0
        if (-not [int]::TryParse(
                [string]$Configuration['DataQuality'][$Name],
                [ref]$candidate
            ) -or $candidate -le 0) {
            throw "DataQuality.$Name must be a positive integer."
        }
        $value = $candidate
    }
    return $value
}

function ConvertTo-SmartWorkplaceCMDBFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Paths,
        [Parameter(Mandatory)][string]$FindingKey,
        [Parameter(Mandatory)][string]$Severity,
        [Parameter(Mandatory)][string]$EntityType,
        [Parameter(Mandatory)][string]$EntityId,
        [Parameter(Mandatory)][string]$FindingType,
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][string]$SourceSystem,
        [Parameter(Mandatory)][string]$DetectedDateTime,
        [Parameter(Mandatory)][string]$RecommendedAction
    )
    return [pscustomobject][ordered]@{
        TenantKey        = $Paths.TenantKey
        OrganizationKey  = $Paths.OrganizationKey
        EnvironmentKey   = $Paths.EnvironmentKey
        TenantId         = $Paths.TenantId
        FindingId        = ('{0}|finding|{1}' -f
            $Paths.TenantKey,
            $FindingKey.ToLowerInvariant())
        Severity         = $Severity
        EntityType       = $EntityType
        EntityId         = $EntityId
        FindingType      = $FindingType
        Description      = $Description
        SourceSystem     = $SourceSystem
        DetectedDateTime = $DetectedDateTime
        RecommendedAction = $RecommendedAction
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

$tableNames = @(
    'CMDB_Users.csv',
    'CMDB_Groups.csv',
    'CMDB_Devices.csv',
    'CMDB_Licenses.csv',
    'CMDB_Mailboxes.csv',
    'FactMailbox.csv',
    'CMDB_DataQuality.csv',
    'FactDataQuality.csv'
)
$tables = @{}
foreach ($name in $tableNames) {
    $match = @($contract.tables | Where-Object name -eq $name)
    if ($match.Count -ne 1) {
        throw "The curated contract must contain exactly one '$name' table."
    }
    $tables[$name] = $match[0]
}

$inputDefinitions = @(
    [pscustomobject]@{
        Name = 'CMDB_Users.csv'
        Path = $UserInputPath
        DefaultPath = Join-Path $paths.CmdbLatestPath 'CMDB_Users.csv'
        KeyColumn = 'CmdbUserId'
        EntityType = 'User'
    },
    [pscustomobject]@{
        Name = 'CMDB_Groups.csv'
        Path = $GroupInputPath
        DefaultPath = Join-Path $paths.CmdbLatestPath 'CMDB_Groups.csv'
        KeyColumn = 'CmdbGroupId'
        EntityType = 'Group'
    },
    [pscustomobject]@{
        Name = 'CMDB_Devices.csv'
        Path = $DeviceInputPath
        DefaultPath = Join-Path $paths.CmdbLatestPath 'CMDB_Devices.csv'
        KeyColumn = 'CmdbDeviceId'
        EntityType = 'Device'
    },
    [pscustomobject]@{
        Name = 'CMDB_Licenses.csv'
        Path = $LicenseInputPath
        DefaultPath = Join-Path $paths.CmdbLatestPath 'CMDB_Licenses.csv'
        KeyColumn = 'CmdbLicenseId'
        EntityType = 'License'
    },
    [pscustomobject]@{
        Name = 'CMDB_Mailboxes.csv'
        Path = $MailboxInputPath
        DefaultPath = Join-Path $paths.CmdbLatestPath 'CMDB_Mailboxes.csv'
        KeyColumn = 'CmdbMailboxId'
        EntityType = 'Mailbox'
    },
    [pscustomobject]@{
        Name = 'FactMailbox.csv'
        Path = $MailboxFactInputPath
        DefaultPath = Join-Path $paths.PowerBILatestPath 'FactMailbox.csv'
        KeyColumn = 'CmdbMailboxId'
        EntityType = 'MailboxFact'
    }
)

$inputRows = @{}
foreach ($definition in $inputDefinitions) {
    if ([string]::IsNullOrWhiteSpace([string]$definition.Path)) {
        $definition.Path = $definition.DefaultPath
    }
    $expectedColumns = @($tables[$definition.Name].columns |
        ForEach-Object { [string]$_ })
    $header = Test-SmartWorkplaceCMDBExactCsvHeader `
        -Path $definition.Path `
        -ExpectedColumns $expectedColumns
    if ($header.Status -ne 'Valid') {
        throw "Input '$($definition.Name)' is missing or incompatible."
    }
    $rows = @(Import-Csv -LiteralPath $definition.Path)
    foreach ($row in $rows) {
        if ([string]$row.TenantKey -ne [string]$paths.TenantKey -or
            [string]$row.OrganizationKey -ne [string]$paths.OrganizationKey -or
            [string]$row.EnvironmentKey -ne [string]$paths.EnvironmentKey -or
            [string]$row.TenantId -ne [string]$paths.TenantId) {
            throw "Input '$($definition.Name)' contains a tenant identity mismatch."
        }
    }
    $inputRows[$definition.Name] = $rows
}

$cmdbOutputPath = Join-Path $paths.CmdbLatestPath 'CMDB_DataQuality.csv'
$factOutputPath = Join-Path $paths.PowerBILatestPath 'FactDataQuality.csv'
$cmdbTargetHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $cmdbOutputPath `
    -ExpectedColumns @($tables['CMDB_DataQuality.csv'].columns |
        ForEach-Object { [string]$_ })
$factTargetHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $factOutputPath `
    -ExpectedColumns @($tables['FactDataQuality.csv'].columns |
        ForEach-Object { [string]$_ })
if ($cmdbTargetHeader.Status -eq 'Incompatible' -or
    $factTargetHeader.Status -eq 'Incompatible') {
    throw 'Existing data-quality output is incompatible with the curated contract.'
}

$warningHours = Get-SmartWorkplaceCMDBPositiveIntegerSetting `
    -Configuration $context.Configuration `
    -Name 'FreshnessWarningHours' `
    -DefaultValue 48 `
    -OverrideValue $FreshnessWarningHours
$criticalHours = Get-SmartWorkplaceCMDBPositiveIntegerSetting `
    -Configuration $context.Configuration `
    -Name 'FreshnessCriticalHours' `
    -DefaultValue 168 `
    -OverrideValue $FreshnessCriticalHours
if ($criticalHours -le $warningHours) {
    throw 'FreshnessCriticalHours must be greater than FreshnessWarningHours.'
}
$referenceUtc = $ReferenceDateTime.ToUniversalTime()
$detectedDateTime = ConvertTo-SmartWorkplaceCMDBUtcDateText $referenceUtc

if ($ValidateOnly) {
    [pscustomobject]@{
        Status = 'Valid'
        ScriptVersion = $ScriptVersion
        ContractVersion = [string]$contract.contractVersion
        InputTableCount = $inputDefinitions.Count
        FreshnessWarningHours = $warningHours
        FreshnessCriticalHours = $criticalHours
        CmdbTargetStatus = $cmdbTargetHeader.Status
        FactTargetStatus = $factTargetHeader.Status
        CmdbOutputPath = $cmdbOutputPath
        FactOutputPath = $factOutputPath
    } | Format-List
    return
}

$findings = New-Object System.Collections.Generic.List[object]
$entityDefinitions = @($inputDefinitions |
    Where-Object Name -ne 'FactMailbox.csv')

foreach ($definition in $entityDefinitions) {
    $rows = @($inputRows[$definition.Name])
    $duplicates = @($rows |
        Group-Object -Property $definition.KeyColumn |
        Where-Object Count -gt 1)
    foreach ($duplicate in $duplicates) {
        $entityId = [string]$duplicate.Name
        $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
            -Paths $paths `
            -FindingKey ('duplicate-entity-key|{0}|{1}' -f
                $definition.EntityType,
                $entityId) `
            -Severity 'Critical' `
            -EntityType $definition.EntityType `
            -EntityId $entityId `
            -FindingType 'DuplicateEntityKey' `
            -Description 'The curated entity key occurs more than once.' `
            -SourceSystem 'SmartWorkplaceCMDB' `
            -DetectedDateTime $detectedDateTime `
            -RecommendedAction 'Review source normalization and republish the entity table.'))
    }

    $validDates = New-Object System.Collections.Generic.List[datetimeoffset]
    foreach ($row in $rows) {
        $entityId = [string]$row.($definition.KeyColumn)
        $dateText = [string]$row.SourceCollectedDateTime
        $parsedDate = ConvertFrom-SmartWorkplaceCMDBDateText $dateText
        if ($null -eq $parsedDate) {
            $findingType = if ([string]::IsNullOrWhiteSpace($dateText)) {
                'MissingSourceCollectedDateTime'
            }
            else {
                'InvalidSourceCollectedDateTime'
            }
            $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
                -Paths $paths `
                -FindingKey ('{0}|{1}|{2}' -f
                    $findingType,
                    $definition.EntityType,
                    $entityId) `
                -Severity 'Warning' `
                -EntityType $definition.EntityType `
                -EntityId $entityId `
                -FindingType $findingType `
                -Description 'The entity has no usable source collection date.' `
                -SourceSystem ([string]$row.SourceSystem) `
                -DetectedDateTime $detectedDateTime `
                -RecommendedAction 'Recollect and normalize the source entity.'))
        }
        else {
            $validDates.Add($parsedDate)
        }
    }

    if ($rows.Count -gt 0 -and $validDates.Count -gt 0) {
        $latestDate = @($validDates | Sort-Object -Descending)[0]
        $ageHours = ($referenceUtc - $latestDate).TotalHours
        if ($ageHours -gt $warningHours) {
            $severity = if ($ageHours -gt $criticalHours) {
                'Critical'
            }
            else {
                'Warning'
            }
            $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
                -Paths $paths `
                -FindingKey ('stale-dataset|{0}' -f $definition.Name) `
                -Severity $severity `
                -EntityType 'Dataset' `
                -EntityId $definition.Name `
                -FindingType 'StaleDataset' `
                -Description 'The latest source collection date exceeds the configured freshness threshold.' `
                -SourceSystem 'SmartWorkplaceCMDB' `
                -DetectedDateTime $detectedDateTime `
                -RecommendedAction 'Run the corresponding autonomous collector and normalizer.'))
        }
    }
}

$knownSourceUsers = New-Object `
    'System.Collections.Generic.HashSet[string]' `
    ([StringComparer]::OrdinalIgnoreCase)
foreach ($user in @($inputRows['CMDB_Users.csv'])) {
    if (-not [string]::IsNullOrWhiteSpace([string]$user.SourceUserId)) {
        [void]$knownSourceUsers.Add([string]$user.SourceUserId)
    }
}
foreach ($device in @($inputRows['CMDB_Devices.csv'])) {
    $primaryUserId = [string]$device.PrimaryUserId
    if (-not [string]::IsNullOrWhiteSpace($primaryUserId) -and
        -not $knownSourceUsers.Contains($primaryUserId)) {
        $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
            -Paths $paths `
            -FindingKey ('orphan-primary-user|{0}' -f
                [string]$device.CmdbDeviceId) `
            -Severity 'Warning' `
            -EntityType 'Device' `
            -EntityId ([string]$device.CmdbDeviceId) `
            -FindingType 'OrphanPrimaryUserReference' `
            -Description 'The device primary user is absent from the curated user inventory.' `
            -SourceSystem ([string]$device.SourceSystem) `
            -DetectedDateTime $detectedDateTime `
            -RecommendedAction 'Refresh Entra users and Intune devices, then rebuild relationships.'))
    }
}

foreach ($fact in @($inputRows['FactMailbox.csv'])) {
    if ([string]::IsNullOrWhiteSpace([string]$fact.CmdbUserId)) {
        $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
            -Paths $paths `
            -FindingKey ('unlinked-mailbox|{0}' -f
                [string]$fact.CmdbMailboxId) `
            -Severity 'Warning' `
            -EntityType 'Mailbox' `
            -EntityId ([string]$fact.CmdbMailboxId) `
            -FindingType 'UnlinkedMailbox' `
            -Description 'The mailbox could not be correlated to a curated Entra user.' `
            -SourceSystem ([string]$fact.SourceSystem) `
            -DetectedDateTime $detectedDateTime `
            -RecommendedAction 'Review mailbox external identifiers and refresh the Entra user inventory.'))
    }
}

$findingRows = @($findings.ToArray() |
    Sort-Object FindingType, EntityType, EntityId)
$duplicateFindingIds = @($findingRows |
    Group-Object FindingId |
    Where-Object Count -gt 1)
if ($duplicateFindingIds.Count -gt 0) {
    throw 'Data-quality normalization produced duplicate finding IDs.'
}
$factRows = @($findingRows | ForEach-Object {
    [pscustomobject][ordered]@{
        TenantKey         = $_.TenantKey
        OrganizationKey   = $_.OrganizationKey
        EnvironmentKey    = $_.EnvironmentKey
        TenantId          = $_.TenantId
        TenantFindingKey  = $_.FindingId
        FindingId         = $_.FindingId
        Severity          = $_.Severity
        EntityType        = $_.EntityType
        FindingType       = $_.FindingType
        DetectedDateTime  = $_.DetectedDateTime
        SourceSystem      = $_.SourceSystem
    }
})

Initialize-SmartWorkplaceCMDBTenantFolder -Paths $paths
Export-SmartWorkplaceCMDBCsv `
    -InputObject $findingRows `
    -Path $cmdbOutputPath `
    -Columns @($tables['CMDB_DataQuality.csv'].columns |
        ForEach-Object { [string]$_ }) | Out-Null
Export-SmartWorkplaceCMDBCsv `
    -InputObject $factRows `
    -Path $factOutputPath `
    -Columns @($tables['FactDataQuality.csv'].columns |
        ForEach-Object { [string]$_ }) | Out-Null

$cmdbValidation = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $cmdbOutputPath `
    -ExpectedColumns @($tables['CMDB_DataQuality.csv'].columns |
        ForEach-Object { [string]$_ })
$factValidation = Test-SmartWorkplaceCMDBExactCsvHeader `
    -Path $factOutputPath `
    -ExpectedColumns @($tables['FactDataQuality.csv'].columns |
        ForEach-Object { [string]$_ })
if ($cmdbValidation.Status -ne 'Valid' -or
    $factValidation.Status -ne 'Valid') {
    throw 'Published data-quality output does not satisfy the curated contract.'
}

$severityCounts = @{}
foreach ($severity in @('Critical', 'Warning', 'Information')) {
    $severityCounts[$severity] = @($findingRows |
        Where-Object Severity -eq $severity).Count
}
Write-Information (
    "SmartWorkplaceCMDB data-quality normalization completed. Findings={0}; Critical={1}; Warning={2}." -f
    $findingRows.Count,
    $severityCounts['Critical'],
    $severityCounts['Warning']
) -InformationAction Continue

[pscustomobject]@{
    Status = 'Completed'
    ScriptVersion = $ScriptVersion
    FindingCount = $findingRows.Count
    CriticalCount = $severityCounts['Critical']
    WarningCount = $severityCounts['Warning']
    InformationCount = $severityCounts['Information']
    CmdbOutputPath = $cmdbOutputPath
    FactOutputPath = $factOutputPath
    ContractVersion = [string]$contract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA0qCQVnsSZPKZj
# Vtm7WFdGb1zhyAlbql1jaN2WnCLUhKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPT/Bi1n+R1V1VYfhlskq4ta7qXCk5u3yHlH3u2HK5SbMA0GCSqG
# SIb3DQEBAQUABIIBgGBT6kfQqHIyi+JrWjQWijINRLoZEtX+GElbVmtZ1ryWOIID
# 4HR8EHCEm+7UvP5T7xmxVH88ZA7WRRHYrefJ1auijd4cyEc1SajX77nYax421mnf
# r1ngv9YWZ9mLyKeIszJ9UpEvd7g31RQUAZSC4flieFdtuzqGnpasuwoWxM4Fh+sP
# 7MvcUnizvT65qpHbHETR2x5MvodDyjRgboxhpRgt+YkEhNS6olw7TH2z1ctoJy94
# vYnH/u9HwwrGflYybHL9xE7OBr8hE/OCQt3vxDiRW+rLS/c4mm7mRe2XpdOiqNBU
# n20sUv6WnBJ5zSMieHD0Plp4h6cZ6IvlFZ1xqTLw4U2cjltUzSPXEgVf9F9nJNYz
# t2w9sfviyKUyAuBfWdk+JMKzo0N6qXc77ZJepEtMp487BWK5Fh9JN9XyOoLqkUoj
# J428m+58Uuhk/FQrUqKbt4mM22Yc3onMhqfnKWU/ecXi1WsWhr3D6B/Aqv1ZZauP
# Xz0fQU4ydAVrBEZ5saGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxOTQy
# MzBaMC8GCSqGSIb3DQEJBDEiBCDgqAQMO+Q5oJX94XbX7GKaj0j6WBN7ca/jsoUe
# JWCyHzANBgkqhkiG9w0BAQEFAASCAgAMOTXa+7Wf5uk9P8wHcNKXBMKgLDlG0Zsz
# t73SvNiJwAN3tTUBUJ7K2hd8tq/zYecdz3tqWmBQdDqTuBGlk8CfaYuQsG7b2jFu
# GXlSmtnVgC7VZmh2aEApPopL8MscPcdDHO53f027t9bbXJeMuLEqzOTMbw2xyqQN
# UcofquTc4XGK2wPuUnexXHpPwnm14RiDkyCCds6/OSRmsmqXrQE/SsIFNq4W3mGt
# 0G3Ts9JgwP7TlWZCB6ZBG0tqwO7eHvsSRBXbwWTmVie27srE7iOoowZ7Y7oRgr6b
# Rhl7tDuPBEh2FoYOgMgvI3QjnX2SB318giMjBb+VQxaXF2w7Z+Cb0HLD+h0ZGOGt
# GU3vuCjg+28rYi31IGc+TfeWHLu5VwQTgiFtt2C7CgzON2EFx/LiEFuuz9s47CnP
# 1f8Wu3KnJWqxXwf47r4zQjoEBDQIf8PiZa0YJn7VCKsv3G3jOCHMDXqbF7lPZ9JK
# +nn5qUKJ+RvdfmF5MRP5Rx+QXMR4Es7QsTIyj4uCGIROfKLatYO4b7oiDL4wMHqp
# xAmbyN7f3m/izSeLcx6j9vW/plJRsYtLMZyS1A/lWtwE8i8IV8fm+EdJgkzhNFr2
# SNHa3exBuQW6c4G81cxzPMENMsmljvvaf5s953sF9G9wUvze38TtaIA2BvgatvlB
# vXbTedf+sw==
# SIG # End signature block
