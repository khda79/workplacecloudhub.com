<#
.SYNOPSIS
Publishes autonomous SmartWorkplaceCMDB data-quality findings.

.DESCRIPTION
Builds CMDB_DataQuality.csv and FactDataQuality.csv from curated local CSV
outputs. The normalizer performs no tenant connection. It reports missing or
invalid source identities and dates, duplicate entity keys, primary-user and
country coverage gaps, observed license-assignment errors, unlinked mailboxes,
and stale non-empty entity datasets.

.VERSION
1.0.0
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
    [string]$UserLicenseAssignmentInputPath,
    [datetimeoffset]$ReferenceDateTime = [datetimeoffset]::UtcNow,
    [int]$FreshnessWarningHours = -1,
    [int]$FreshnessCriticalHours = -1,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.0.0'
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
        SourceIdColumn = 'SourceUserId'
        EntityType = 'User'
    },
    [pscustomobject]@{
        Name = 'CMDB_Groups.csv'
        Path = $GroupInputPath
        DefaultPath = Join-Path $paths.CmdbLatestPath 'CMDB_Groups.csv'
        KeyColumn = 'CmdbGroupId'
        SourceIdColumn = 'SourceGroupId'
        EntityType = 'Group'
    },
    [pscustomobject]@{
        Name = 'CMDB_Devices.csv'
        Path = $DeviceInputPath
        DefaultPath = Join-Path $paths.CmdbLatestPath 'CMDB_Devices.csv'
        KeyColumn = 'CmdbDeviceId'
        SourceIdColumn = 'SourceDeviceId'
        EntityType = 'Device'
    },
    [pscustomobject]@{
        Name = 'CMDB_Licenses.csv'
        Path = $LicenseInputPath
        DefaultPath = Join-Path $paths.CmdbLatestPath 'CMDB_Licenses.csv'
        KeyColumn = 'CmdbLicenseId'
        SourceIdColumn = 'SkuId'
        EntityType = 'License'
    },
    [pscustomobject]@{
        Name = 'CMDB_Mailboxes.csv'
        Path = $MailboxInputPath
        DefaultPath = Join-Path $paths.CmdbLatestPath 'CMDB_Mailboxes.csv'
        KeyColumn = 'CmdbMailboxId'
        SourceIdColumn = ''
        EntityType = 'Mailbox'
    },
    [pscustomobject]@{
        Name = 'FactMailbox.csv'
        Path = $MailboxFactInputPath
        DefaultPath = Join-Path $paths.PowerBILatestPath 'FactMailbox.csv'
        KeyColumn = 'CmdbMailboxId'
        SourceIdColumn = ''
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

$licenseAssignmentTable = @($rawContract.tables |
    Where-Object name -eq 'M365_UserLicenseAssignments.csv')
if ($licenseAssignmentTable.Count -ne 1) {
    throw "The raw contract must contain exactly one 'M365_UserLicenseAssignments.csv' table."
}
$licenseAssignmentPathWasExplicit = -not [string]::IsNullOrWhiteSpace(
    $UserLicenseAssignmentInputPath
)
if (-not $licenseAssignmentPathWasExplicit) {
    $UserLicenseAssignmentInputPath = Join-Path `
        $paths.LatestOutputRootPath `
        (Join-Path `
            ([string]$licenseAssignmentTable[0].area) `
            ([string]$licenseAssignmentTable[0].name))
}
$licenseAssignmentRows = @()
if (Test-Path -LiteralPath $UserLicenseAssignmentInputPath -PathType Leaf) {
    $assignmentHeader = Test-SmartWorkplaceCMDBExactCsvHeader `
        -Path $UserLicenseAssignmentInputPath `
        -ExpectedColumns @($licenseAssignmentTable[0].columns |
            ForEach-Object { [string]$_ })
    if ($assignmentHeader.Status -ne 'Valid') {
        throw "Input 'M365_UserLicenseAssignments.csv' is incompatible."
    }
    $licenseAssignmentRows = @(Import-Csv -LiteralPath $UserLicenseAssignmentInputPath)
    foreach ($row in $licenseAssignmentRows) {
        if ([string]$row.TenantKey -ne [string]$paths.TenantKey -or
            [string]$row.OrganizationKey -ne [string]$paths.OrganizationKey -or
            [string]$row.EnvironmentKey -ne [string]$paths.EnvironmentKey -or
            [string]$row.TenantId -ne [string]$paths.TenantId) {
            throw "Input 'M365_UserLicenseAssignments.csv' contains a tenant identity mismatch."
        }
    }
}
elseif ($licenseAssignmentPathWasExplicit) {
    throw "Input 'M365_UserLicenseAssignments.csv' is missing."
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
        LicenseAssignmentInputStatus = if ($licenseAssignmentRows.Count -gt 0) {
            'Valid'
        }
        else {
            'NotProvided'
        }
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
$sourceHealth = @(Get-SmartWorkplaceCMDBSourceHealth -Paths $paths -ReferenceDateTime $referenceUtc -WarningHours $warningHours -CriticalHours $criticalHours)
foreach ($source in @($sourceHealth | Where-Object { $_.HasEvidence -and $_.Status -ne 'Complete' })) {
    $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding -Paths $paths `
        -FindingKey ('source-health|{0}' -f $source.SourceName) `
        -Severity $source.Severity -EntityType 'Dataset' -EntityId $source.SourceName `
        -FindingType ('Source{0}' -f $source.Status) `
        -Description ('Source evidence status: {0}; coverage: {1}. This is not proof of a fresh complete live snapshot.' -f $source.Status, $source.Coverage) `
        -SourceSystem 'SmartWorkplaceCMDB' -DetectedDateTime $detectedDateTime `
        -RecommendedAction 'Review collection evidence and refresh the source before using it for coverage decisions.'))
}
$entityDefinitions = @($inputDefinitions |
    Where-Object Name -ne 'FactMailbox.csv')

foreach ($definition in $entityDefinitions) {
    $rows = @($inputRows[$definition.Name])
    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        $row = $rows[$rowIndex]
        $entityId = [string]$row.($definition.KeyColumn)
        if ([string]::IsNullOrWhiteSpace($entityId)) {
            $entityId = '{0}|row-{1}' -f $definition.Name, ($rowIndex + 1)
            $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
                -Paths $paths `
                -FindingKey ('missing-entity-key|{0}|{1}' -f
                    $definition.EntityType,
                    ($rowIndex + 1)) `
                -Severity 'Critical' `
                -EntityType $definition.EntityType `
                -EntityId $entityId `
                -FindingType 'MissingEntityKey' `
                -Description 'The curated entity has no CMDB entity key.' `
                -SourceSystem 'SmartWorkplaceCMDB' `
                -DetectedDateTime $detectedDateTime `
                -RecommendedAction 'Correct source normalization before using this entity downstream.'))
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$definition.SourceIdColumn) -and
            ([string]::IsNullOrWhiteSpace([string]$row.SourceSystem) -or
                [string]::IsNullOrWhiteSpace(
                    [string]$row.($definition.SourceIdColumn)
                ))) {
            $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
                -Paths $paths `
                -FindingKey ('missing-source-identity|{0}|{1}' -f
                    $definition.EntityType,
                    $entityId) `
                -Severity 'Warning' `
                -EntityType $definition.EntityType `
                -EntityId $entityId `
                -FindingType 'MissingSourceIdentity' `
                -Description 'SourceSystem plus the source-specific identifier is incomplete.' `
                -SourceSystem 'SmartWorkplaceCMDB' `
                -DetectedDateTime $detectedDateTime `
                -RecommendedAction 'Recollect and normalize the source identity without inventing a global SourceID.'))
        }
    }
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
        $oldestDate = @($validDates | Sort-Object)[0]
        $ageHours = ($referenceUtc - $oldestDate).TotalHours
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
                -Description 'At least one source collection date exceeds the configured freshness threshold.' `
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
$usersBySourceId = @{}
foreach ($user in @($inputRows['CMDB_Users.csv'])) {
    $sourceUserId = [string]$user.SourceUserId
    if (-not [string]::IsNullOrWhiteSpace($sourceUserId) -and
        -not $usersBySourceId.ContainsKey($sourceUserId)) {
        $usersBySourceId[$sourceUserId] = $user
    }
    if ([string]::IsNullOrWhiteSpace([string]$user.UsageLocation) -or
        [string]$user.UsageLocationStatus -ine 'Reported') {
        $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
            -Paths $paths `
            -FindingKey ('user-country-unknown|{0}' -f [string]$user.CmdbUserId) `
            -Severity 'Warning' `
            -EntityType 'User' `
            -EntityId ([string]$user.CmdbUserId) `
            -FindingType 'UserCountryUnknown' `
            -Description 'The user country is not reported by the authoritative user source.' `
            -SourceSystem ([string]$user.SourceSystem) `
            -DetectedDateTime $detectedDateTime `
            -RecommendedAction 'Populate UsageLocation in Entra ID or document the user as Not Reported.'))
    }
}
foreach ($device in @($inputRows['CMDB_Devices.csv'])) {
    $primaryUserId = [string]$device.PrimaryUserId
    if ([string]::IsNullOrWhiteSpace($primaryUserId)) {
        $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
            -Paths $paths `
            -FindingKey ('device-without-primary-user|{0}' -f
                [string]$device.CmdbDeviceId) `
            -Severity 'Warning' `
            -EntityType 'Device' `
            -EntityId ([string]$device.CmdbDeviceId) `
            -FindingType 'DeviceWithoutPrimaryUser' `
            -Description 'The device has no observed primary user.' `
            -SourceSystem ([string]$device.SourceSystem) `
            -DetectedDateTime $detectedDateTime `
            -RecommendedAction 'Review Intune primary-user assignment; do not infer ownership from device name.'))
    }
    elseif (
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
    $userCountryIsKnown = (
        -not [string]::IsNullOrWhiteSpace($primaryUserId) -and
        $usersBySourceId.ContainsKey($primaryUserId) -and
        -not [string]::IsNullOrWhiteSpace(
            [string]$usersBySourceId[$primaryUserId].UsageLocation
        ) -and
        [string]$usersBySourceId[$primaryUserId].UsageLocationStatus -ieq 'Reported'
    )
    if (-not $userCountryIsKnown) {
        $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
            -Paths $paths `
            -FindingKey ('device-country-unknown|{0}' -f
                [string]$device.CmdbDeviceId) `
            -Severity 'Warning' `
            -EntityType 'Device' `
            -EntityId ([string]$device.CmdbDeviceId) `
            -FindingType 'DeviceCountryUnknown' `
            -Description 'The device country cannot be derived from a reported primary-user country.' `
            -SourceSystem ([string]$device.SourceSystem) `
            -DetectedDateTime $detectedDateTime `
            -RecommendedAction 'Complete the primary-user link and user UsageLocation; keep country as Not Reported until then.'))
    }
}

foreach ($assignment in $licenseAssignmentRows) {
    $assignmentError = ([string]$assignment.AssignmentError).Trim()
    if ([string]::IsNullOrWhiteSpace($assignmentError) -or
        $assignmentError -ieq 'None') {
        continue
    }
    $assignmentKey = [string]$assignment.RawAssignmentKey
    if ([string]::IsNullOrWhiteSpace($assignmentKey)) {
        $assignmentKey = '{0}|{1}' -f
            [string]$assignment.SourceUserId,
            [string]$assignment.SkuId
    }
    $entityId = $assignmentKey
    if ($usersBySourceId.ContainsKey([string]$assignment.SourceUserId)) {
        $entityId = [string]$usersBySourceId[[string]$assignment.SourceUserId].CmdbUserId
    }
    $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
        -Paths $paths `
        -FindingKey ('license-assignment-error|{0}' -f $assignmentKey) `
        -Severity 'Warning' `
        -EntityType 'UserLicenseAssignment' `
        -EntityId $entityId `
        -FindingType 'ObservedLicenseAssignmentError' `
        -Description ('Microsoft reported license assignment error: {0}.' -f $assignmentError) `
        -SourceSystem ([string]$assignment.SourceSystem) `
        -DetectedDateTime $detectedDateTime `
        -RecommendedAction 'Review the reported license assignment state and source group; do not infer a financial compliance issue.'))
}

$mailboxesById = @{}
foreach ($mailbox in @($inputRows['CMDB_Mailboxes.csv'])) {
    $mailboxId = [string]$mailbox.CmdbMailboxId
    if (-not $mailboxesById.ContainsKey($mailboxId)) {
        $mailboxesById[$mailboxId] = @()
    }
    $mailboxesById[$mailboxId] += $mailbox
}
foreach ($fact in @($inputRows['FactMailbox.csv'])) {
    if ([string]::IsNullOrWhiteSpace([string]$fact.CmdbUserId)) {
        # Only a corroborated technical type without an external ID is informational.
        # Keep the record and stable finding key; unresolved ordinary mailboxes remain warnings.
        $mailboxId = [string]$fact.CmdbMailboxId
        $mailboxMatches = @(if ($mailboxesById.ContainsKey($mailboxId)) { $mailboxesById[$mailboxId] })
        $technical = ($mailboxMatches.Count -eq 1 -and
            ([string]$fact.RecipientTypeDetails).Trim() -ieq 'DiscoveryMailbox' -and
            ([string]$mailboxMatches[0].RecipientTypeDetails).Trim() -ieq 'DiscoveryMailbox' -and
            [string]::IsNullOrWhiteSpace([string]$mailboxMatches[0].ExternalDirectoryObjectId))
        $severity = if ($technical) { 'Information' } else { 'Warning' }
        $findingType = if ($technical) { 'TechnicalMailboxWithoutUser' } else { 'UnlinkedMailbox' }
        $description = if ($technical) {
            'A DiscoveryMailbox without an external directory identifier is retained as a technical mailbox without a user link.'
        } else { 'The mailbox could not be correlated to a curated Entra user.' }
        $action = if ($technical) {
            'Review as a technical mailbox; do not infer a missing user or create an automatic user association.'
        } else { 'Review mailbox external identifiers and refresh the Entra user inventory.' }
        $findings.Add((ConvertTo-SmartWorkplaceCMDBFinding `
            -Paths $paths `
            -FindingKey ('unlinked-mailbox|{0}' -f
                [string]$fact.CmdbMailboxId) `
            -Severity $severity `
            -EntityType 'Mailbox' `
            -EntityId ([string]$fact.CmdbMailboxId) `
            -FindingType $findingType `
            -Description $description `
            -SourceSystem ([string]$fact.SourceSystem) `
            -DetectedDateTime $detectedDateTime `
            -RecommendedAction $action))
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAgCZ6N1WL3RJoi
# Nnl/fhcLT83E9PSdVLxP2ZpVHSG83qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEdVlRu1+I7UCYuPLCKjpuqPHPGzW1//ayFMeRdeUMJ4MA0GCSqG
# SIb3DQEBAQUABIIBgA45G3jAIixaAnkhQ7+jx+EdVFgN5gT3WdMn4RlumIbGHuRp
# o5chNIFUYNyuPRPOB+vBKQcId7lEG/QDuPPi280CwjYU36nfv0RI1HnTbRfvjLuu
# HfuMGHcV7tDcd25IKCvXOpCdvE4emwBctSVqSCZ3opvDKnUggNQvH33X6pr/7a2P
# BK5u8QUsJN8LLRcr1G/pXGOdO/OMMDCzxNYQmGaNYwLkvCz5TlsmwD7K7xPU2Txt
# 2AbEUyCr53iIbcghxfu4AzyYru/Gtqa/3i+w08ruHxORGd1kXePpg6MA+uI7HZ9z
# PNef60TzDSD+v+HT0FcByFTHOUtNcoMYVoPNwMgUj23HaBH+YwYR9fxVA6JAxcyY
# tLAST7bdhp3AIVDPd2xOto1ygUkpQxAweDJz1svnXECE+skRHNc5WDLS59EJJtWS
# n/FqMu7Pe4Mhwhj7E3Z5Tt1Kxm59CP3sFIbMKc1dbyFujqSA8I5PR/x6xFZvSl8W
# IQt13bljpQVZmb3U3aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxMzI1
# MjBaMC8GCSqGSIb3DQEJBDEiBCAVRjzshTm6MA7LRasfMpWNUlkVwr40hnPorktr
# xQOHKzANBgkqhkiG9w0BAQEFAASCAgCTK//DaAngbPlEkVWUQU+0hrIBSQIbCWIK
# AeofBwvZRFaJX3fsaT5plWNirUTev1iyVTMZiRKR7okop6d9G4BXVpLFo/BExtlf
# EhzfQocrLn5Se7jExR+Fd4RJ/eSyZtXn5o7gmCLRZ7s67pQDRe5IPyeQVkroOhcA
# UUCO5l9SQXcBXSt4edE55VsHLkKJCnrfHYEpxBx1stQeepVicy+GbOBDTK047nO5
# nyYCkzAZJEt/RBnIAk2VLb1AyODVcyi63YHUuCVkvrmU8+wj7P1QvtVctUW9MYyd
# zzlfkMHeNPfpsDtqa3y8AbMZPbz5auO1EfNle2BFNyye+O7d5E2ruVRAiCOv/RHi
# TLCKOIUGU0WWx3s7+PWqV6ClCYutdyS+DHdcTMpE4KHQQvswov2gEklhCqMCNmdZ
# s3he7ZI57SuBGHCGtOwfGWhIWPc9pFJN1XdeFyx0BlUyYuAJF1+k5O2X9xMOi+3A
# maYv/agxdBBLMlsQ/JOptx9MsTM4Oi1ydvE+ktLu0WxCFlwFum2hrtV8ob0IK0no
# 0ZNd5yXwd6tJQlZUuJfBnppFExJhnEGeg4QTSpchvxtIcIG7UNftDSNIgRODRGSA
# FdtsK6ZDhA4HQ5Cc0VnNIGD08V9PZLzuLkMhSLNxXNVxkE8BJioHXk1ME/axW4i+
# fMrlaO30gg==
# SIG # End signature block
