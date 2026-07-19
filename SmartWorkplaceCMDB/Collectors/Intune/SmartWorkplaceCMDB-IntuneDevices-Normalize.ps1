<#
.SYNOPSIS
Enriches SmartWorkplaceCMDB devices with Microsoft Intune managed-device data.

.DESCRIPTION
Validates the latest raw Entra and Intune device contracts, joins records by
Intune azureADDeviceId to Entra deviceId, and republishes CMDB_Devices.csv and
DimDevice.csv. Intune-only records are retained. When azureADDeviceId is empty,
the stable fallback source key is "intune:<managedDeviceId>".

.VERSION
0.1.0
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
    [string]$EntraRawInputPath,
    [string]$IntuneRawInputPath,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '0.1.1'
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

function ConvertTo-SmartWorkplaceCMDBNormalizedBoolean {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Value,
        [Parameter(Mandatory)][string]$FieldName,
        [Parameter(Mandatory)][string]$SourceDeviceId
    )
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }
    $parsed = $false
    if (-not [bool]::TryParse($Value, [ref]$parsed)) {
        throw "$FieldName '$Value' is invalid for source device '$SourceDeviceId'."
    }
    return $parsed
}

function ConvertTo-SmartWorkplaceCMDBComplianceState {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Value)
    switch ($Value.Trim().ToLowerInvariant()) {
        '' { return '' }
        'compliant' { return 'Compliant' }
        'noncompliant' { return 'NonCompliant' }
        'ingraceperiod' { return 'InGracePeriod' }
        'configmanager' { return 'ConfigManager' }
        'conflict' { return 'Conflict' }
        'error' { return 'Error' }
        'unknown' { return 'Unknown' }
        default { return $Value.Trim() }
    }
}

function ConvertTo-SmartWorkplaceCMDBOwnership {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Value)
    switch ($Value.Trim().ToLowerInvariant()) {
        '' { return '' }
        'company' { return 'Corporate' }
        'personal' { return 'Personal' }
        'unknown' { return 'Unknown' }
        default { return $Value.Trim() }
    }
}

function Get-SmartWorkplaceCMDBDefaultConfidenceScore {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Configuration)
    $value = 0.5
    if ($Configuration.Contains('DataQuality') -and
        $Configuration['DataQuality'] -is [System.Collections.IDictionary] -and
        $Configuration['DataQuality'].Contains('DefaultConfidenceScore')) {
        $candidate = 0.0
        if ([double]::TryParse(
                [string]$Configuration['DataQuality']['DefaultConfidenceScore'],
                [Globalization.NumberStyles]::Float,
                [Globalization.CultureInfo]::InvariantCulture,
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

function Get-SmartWorkplaceCMDBDateSortValue {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Value,
        [Parameter(Mandatory)][string]$FieldName,
        [Parameter(Mandatory)][string]$SourceId
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
        throw "$FieldName '$Value' is invalid for source device '$SourceId'."
    }
    return $parsed.ToUniversalTime()
}

function Select-SmartWorkplaceCMDBNewerDateText {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$First,
        [AllowEmptyString()][string]$Second,
        [Parameter(Mandatory)][string]$SourceId
    )
    $firstValue = Get-SmartWorkplaceCMDBDateSortValue $First 'SourceCollectedDateTime' $SourceId
    $secondValue = Get-SmartWorkplaceCMDBDateSortValue $Second 'SourceCollectedDateTime' $SourceId
    if ($secondValue -gt $firstValue) {
        return $Second
    }
    return $First
}

function Test-SmartWorkplaceCMDBUsableAzureAdDeviceId {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }
    return $Value.Trim() -notmatch '^0{8}-0{4}-0{4}-0{4}-0{12}$'
}

function Get-SmartWorkplaceCMDBPreferredText {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Preferred,
        [AllowEmptyString()][string]$Fallback
    )
    if (-not [string]::IsNullOrWhiteSpace($Preferred)) {
        return $Preferred
    }
    return $Fallback
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
$entraTable = @($rawContract.tables | Where-Object name -eq 'Entra_Devices.csv')
$intuneTable = @($rawContract.tables | Where-Object name -eq 'Intune_ManagedDevices.csv')
$cmdbTable = @($curatedContract.tables | Where-Object name -eq 'CMDB_Devices.csv')
$dimTable = @($curatedContract.tables | Where-Object name -eq 'DimDevice.csv')
$complianceTable = @($curatedContract.tables | Where-Object name -eq 'FactDeviceCompliance.csv')
if ($entraTable.Count -ne 1 -or $intuneTable.Count -ne 1 -or
    $cmdbTable.Count -ne 1 -or $dimTable.Count -ne 1 -or
    $complianceTable.Count -ne 1) {
    throw 'The contracts must contain one Entra device, Intune managed device, CMDB device, DimDevice, and FactDeviceCompliance definition.'
}
$entraTable = $entraTable[0]; $intuneTable = $intuneTable[0]
$cmdbTable = $cmdbTable[0]; $dimTable = $dimTable[0]
$complianceTable = $complianceTable[0]

if ([string]::IsNullOrWhiteSpace($EntraRawInputPath)) {
    $EntraRawInputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$entraTable.area) ([string]$entraTable.name)
    )
}
if ([string]::IsNullOrWhiteSpace($IntuneRawInputPath)) {
    $IntuneRawInputPath = Join-Path $paths.LatestOutputRootPath (
        Join-Path ([string]$intuneTable.area) ([string]$intuneTable.name)
    )
}
$EntraRawInputPath = [IO.Path]::GetFullPath($EntraRawInputPath)
$IntuneRawInputPath = [IO.Path]::GetFullPath($IntuneRawInputPath)

$entraHeader = Test-SmartWorkplaceCMDBExactCsvHeader $EntraRawInputPath @($entraTable.columns | ForEach-Object { [string]$_ })
$intuneHeader = Test-SmartWorkplaceCMDBExactCsvHeader $IntuneRawInputPath @($intuneTable.columns | ForEach-Object { [string]$_ })
foreach ($rawInput in @(
        [pscustomobject]@{ Name = $entraTable.name; Header = $entraHeader; Path = $EntraRawInputPath },
        [pscustomobject]@{ Name = $intuneTable.name; Header = $intuneHeader; Path = $IntuneRawInputPath }
    )) {
    if ($rawInput.Header.Status -eq 'Incompatible') {
        throw "Raw table '$($rawInput.Name)' is incompatible."
    }
    if ($rawInput.Header.Status -eq 'Missing' -and -not $ValidateOnly) {
        throw "Raw table '$($rawInput.Name)' was not found: $($rawInput.Path)"
    }
}

$cmdbOutputPath = Join-Path $paths.CmdbLatestPath ([string]$cmdbTable.name)
$dimOutputPath = Join-Path $paths.PowerBILatestPath ([string]$dimTable.name)
$complianceOutputPath = Join-Path $paths.PowerBILatestPath ([string]$complianceTable.name)
$cmdbHeader = Test-SmartWorkplaceCMDBExactCsvHeader $cmdbOutputPath @($cmdbTable.columns | ForEach-Object { [string]$_ })
$dimHeader = Test-SmartWorkplaceCMDBExactCsvHeader $dimOutputPath @($dimTable.columns | ForEach-Object { [string]$_ })
$complianceHeader = Test-SmartWorkplaceCMDBExactCsvHeader $complianceOutputPath @($complianceTable.columns | ForEach-Object { [string]$_ })
foreach ($target in @(
        [pscustomobject]@{ Name = $cmdbTable.name; Header = $cmdbHeader },
        [pscustomobject]@{ Name = $dimTable.name; Header = $dimHeader },
        [pscustomobject]@{ Name = $complianceTable.name; Header = $complianceHeader }
    )) {
    if ($target.Header.Status -eq 'Incompatible') {
        throw "Existing curated table '$($target.Name)' is incompatible."
    }
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status = 'Valid'; ScriptVersion = $ScriptVersion
        RawContractVersion = [string]$rawContract.contractVersion
        CuratedContractVersion = [string]$curatedContract.contractVersion
        EntraRawInputStatus = $entraHeader.Status; EntraRawInputPath = $EntraRawInputPath
        IntuneRawInputStatus = $intuneHeader.Status; IntuneRawInputPath = $IntuneRawInputPath
        CmdbDeviceTargetStatus = $cmdbHeader.Status; DimDeviceTargetStatus = $dimHeader.Status
        FactDeviceComplianceTargetStatus = $complianceHeader.Status
        CmdbDeviceOutputPath = $cmdbOutputPath; DimDeviceOutputPath = $dimOutputPath
        FactDeviceComplianceOutputPath = $complianceOutputPath
    } | Format-List
    return
}

$entraRows = @(Import-Csv -LiteralPath $EntraRawInputPath -ErrorAction Stop)
$intuneRows = @(Import-Csv -LiteralPath $IntuneRawInputPath -ErrorAction Stop)
$identityFields = [ordered]@{
    TenantKey = $paths.TenantKey; OrganizationKey = $paths.OrganizationKey
    EnvironmentKey = $paths.EnvironmentKey; TenantId = $paths.TenantId
}
foreach ($source in @(
        [pscustomobject]@{ Name = 'Entra'; Rows = $entraRows; Key = 'SourceDeviceId' },
        [pscustomobject]@{ Name = 'Intune'; Rows = $intuneRows; Key = 'ManagedDeviceId' }
    )) {
    foreach ($row in @($source.Rows)) {
        foreach ($identityName in $identityFields.Keys) {
            if ([string]$row.$identityName -ne [string]$identityFields[$identityName]) {
                throw "Raw $($source.Name) devices identity mismatch for '$identityName'."
            }
        }
        $sourceKey = [string]$row.($source.Key)
        if ([string]::IsNullOrWhiteSpace($sourceKey)) {
            throw "Raw $($source.Name) devices data contains an empty $($source.Key)."
        }
    }
}
$duplicateEntraIds = @($entraRows | Group-Object SourceDeviceId | Where-Object Count -gt 1)
if ($duplicateEntraIds.Count -gt 0) {
    throw "Raw Entra devices data contains duplicate SourceDeviceId values: $($duplicateEntraIds.Name -join ', ')"
}
$duplicateIntuneIds = @($intuneRows | Group-Object ManagedDeviceId | Where-Object Count -gt 1)
if ($duplicateIntuneIds.Count -gt 0) {
    throw "Raw Intune devices data contains duplicate ManagedDeviceId values: $($duplicateIntuneIds.Name -join ', ')"
}

$usableIntuneRows = @($intuneRows | Where-Object {
    Test-SmartWorkplaceCMDBUsableAzureAdDeviceId ([string]$_.AzureAdDeviceId)
})
$duplicateCorrelationCount = @(
    $usableIntuneRows |
        Group-Object { ([string]$_.AzureAdDeviceId).Trim().ToLowerInvariant() } |
        Where-Object Count -gt 1
).Count
$intuneByAzureId = @{}
foreach ($group in @($usableIntuneRows | Group-Object {
            ([string]$_.AzureAdDeviceId).Trim().ToLowerInvariant()
        })) {
    $selected = @($group.Group | Sort-Object `
        @{ Expression = { Get-SmartWorkplaceCMDBDateSortValue ([string]$_.LastSyncDateTime) 'LastSyncDateTime' ([string]$_.ManagedDeviceId) }; Descending = $true },
        @{ Expression = { Get-SmartWorkplaceCMDBDateSortValue ([string]$_.EnrolledDateTime) 'EnrolledDateTime' ([string]$_.ManagedDeviceId) }; Descending = $true },
        @{ Expression = { [string]$_.ManagedDeviceId }; Descending = $false })[0]
    $intuneByAzureId[[string]$group.Name] = $selected
}

$confidence = (Get-SmartWorkplaceCMDBDefaultConfidenceScore $context.Configuration).ToString(
    '0.################',
    [Globalization.CultureInfo]::InvariantCulture
)
$consumedIntuneIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$outputRows = New-Object System.Collections.Generic.List[object]
$enrichedCount = 0
foreach ($entra in @($entraRows | Sort-Object DeviceName, SourceDeviceId)) {
    $sourceDeviceId = ([string]$entra.SourceDeviceId).Trim()
    $correlationKey = $sourceDeviceId.ToLowerInvariant()
    $intune = if ($intuneByAzureId.ContainsKey($correlationKey)) { $intuneByAzureId[$correlationKey] } else { $null }
    if ($null -ne $intune) {
        [void]$consumedIntuneIds.Add([string]$intune.ManagedDeviceId)
        $enrichedCount++
    }
    $managed = ConvertTo-SmartWorkplaceCMDBNormalizedBoolean ([string]$entra.IsManaged) 'IsManaged' $sourceDeviceId
    $compliant = ConvertTo-SmartWorkplaceCMDBNormalizedBoolean ([string]$entra.IsCompliant) 'IsCompliant' $sourceDeviceId
    $baselineCompliance = if ($compliant -eq $true) { 'Compliant' } elseif ($compliant -eq $false -and $compliant -is [bool]) { 'NonCompliant' } else { '' }
    $baselineManagement = if ($managed -eq $true) { 'Managed' } elseif ($managed -eq $false -and $managed -is [bool]) { 'Unmanaged' } else { '' }
    $outputRows.Add([pscustomobject][ordered]@{
        CmdbDeviceId = ('{0}|device|{1}' -f $paths.TenantKey, $correlationKey)
        SourceSystem = if ($null -ne $intune) { 'MicrosoftEntraID+MicrosoftIntune' } else { [string]$entra.SourceSystem }
        SourceDeviceId = $sourceDeviceId
        DeviceName = if ($null -ne $intune) { Get-SmartWorkplaceCMDBPreferredText ([string]$intune.DeviceName) ([string]$entra.DeviceName) } else { [string]$entra.DeviceName }
        OperatingSystem = if ($null -ne $intune) { Get-SmartWorkplaceCMDBPreferredText ([string]$intune.OperatingSystem) ([string]$entra.OperatingSystem) } else { [string]$entra.OperatingSystem }
        OperatingSystemVersion = if ($null -ne $intune) { Get-SmartWorkplaceCMDBPreferredText ([string]$intune.OperatingSystemVersion) ([string]$entra.OperatingSystemVersion) } else { [string]$entra.OperatingSystemVersion }
        Ownership = if ($null -ne $intune) { ConvertTo-SmartWorkplaceCMDBOwnership ([string]$intune.ManagedDeviceOwnerType) } else { '' }
        ComplianceState = if ($null -ne $intune) { ConvertTo-SmartWorkplaceCMDBComplianceState ([string]$intune.ComplianceState) } else { $baselineCompliance }
        ManagementState = if ($null -ne $intune) { 'Managed' } else { $baselineManagement }
        PrimaryUserId = if ($null -ne $intune) { [string]$intune.UserId } else { '' }
        LastSyncDateTime = if ($null -ne $intune) { [string]$intune.LastSyncDateTime } else { '' }
        ConfidenceScore = $confidence
        SourceCollectedDateTime = if ($null -ne $intune) {
            Select-SmartWorkplaceCMDBNewerDateText ([string]$entra.SourceCollectedDateTime) ([string]$intune.SourceCollectedDateTime) $sourceDeviceId
        } else { [string]$entra.SourceCollectedDateTime }
    })
}

$intuneOnlyCount = 0
$missingAzureAdDeviceIdCount = 0
foreach ($intune in @($intuneRows | Sort-Object DeviceName, ManagedDeviceId)) {
    if ($consumedIntuneIds.Contains([string]$intune.ManagedDeviceId)) {
        continue
    }
    $hasAzureId = Test-SmartWorkplaceCMDBUsableAzureAdDeviceId ([string]$intune.AzureAdDeviceId)
    if ($hasAzureId) {
        $correlationKey = ([string]$intune.AzureAdDeviceId).Trim().ToLowerInvariant()
        if ($intuneByAzureId[$correlationKey].ManagedDeviceId -ne $intune.ManagedDeviceId) {
            continue
        }
        $sourceDeviceId = ([string]$intune.AzureAdDeviceId).Trim()
    }
    else {
        $sourceDeviceId = 'intune:{0}' -f ([string]$intune.ManagedDeviceId).Trim().ToLowerInvariant()
        $missingAzureAdDeviceIdCount++
    }
    $intuneOnlyCount++
    $outputRows.Add([pscustomobject][ordered]@{
        CmdbDeviceId = ('{0}|device|{1}' -f $paths.TenantKey, $sourceDeviceId.ToLowerInvariant())
        SourceSystem = 'MicrosoftIntune'
        SourceDeviceId = $sourceDeviceId
        DeviceName = [string]$intune.DeviceName
        OperatingSystem = [string]$intune.OperatingSystem
        OperatingSystemVersion = [string]$intune.OperatingSystemVersion
        Ownership = ConvertTo-SmartWorkplaceCMDBOwnership ([string]$intune.ManagedDeviceOwnerType)
        ComplianceState = ConvertTo-SmartWorkplaceCMDBComplianceState ([string]$intune.ComplianceState)
        ManagementState = 'Managed'
        PrimaryUserId = [string]$intune.UserId
        LastSyncDateTime = [string]$intune.LastSyncDateTime
        ConfidenceScore = $confidence
        SourceCollectedDateTime = [string]$intune.SourceCollectedDateTime
    })
}

$cmdbRows = @($outputRows.ToArray() | Sort-Object DeviceName, SourceDeviceId)
$duplicateOutputIds = @($cmdbRows | Group-Object CmdbDeviceId | Where-Object Count -gt 1)
if ($duplicateOutputIds.Count -gt 0) {
    throw "Device enrichment produced duplicate CmdbDeviceId values: $($duplicateOutputIds.Name -join ', ')"
}
$dimRows = @($cmdbRows | ForEach-Object {
    [pscustomobject][ordered]@{
        TenantDeviceKey = $_.CmdbDeviceId; CmdbDeviceId = $_.CmdbDeviceId
        DeviceName = $_.DeviceName; OperatingSystem = $_.OperatingSystem
        OperatingSystemVersion = $_.OperatingSystemVersion; Ownership = $_.Ownership
        ComplianceState = $_.ComplianceState; ManagementState = $_.ManagementState
        ConfidenceScore = $_.ConfidenceScore
    }
})
$complianceRows = @($cmdbRows | ForEach-Object {
    [pscustomobject][ordered]@{
        TenantDeviceKey = $_.CmdbDeviceId
        CmdbDeviceId = $_.CmdbDeviceId
        ComplianceState = $_.ComplianceState
        LastSyncDateTime = $_.LastSyncDateTime
        SourceSystem = $_.SourceSystem
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
Export-SmartWorkplaceCMDBCsv -InputObject $complianceRows -Path $complianceOutputPath `
    -Columns @($complianceTable.columns | ForEach-Object { [string]$_ }) @identityExport

$cmdbValidation = Test-SmartWorkplaceCMDBExactCsvHeader $cmdbOutputPath @($cmdbTable.columns | ForEach-Object { [string]$_ })
$dimValidation = Test-SmartWorkplaceCMDBExactCsvHeader $dimOutputPath @($dimTable.columns | ForEach-Object { [string]$_ })
$complianceValidation = Test-SmartWorkplaceCMDBExactCsvHeader $complianceOutputPath @($complianceTable.columns | ForEach-Object { [string]$_ })
if ($cmdbValidation.Status -ne 'Valid' -or $dimValidation.Status -ne 'Valid' -or
    $complianceValidation.Status -ne 'Valid') {
    throw 'Enriched device outputs did not satisfy the curated contracts.'
}
Write-Information (
    "SmartWorkplaceCMDB Intune device enrichment completed. Entra={0}; Intune={1}; Enriched={2}; IntuneOnly={3}; CMDB={4}." -f
    $entraRows.Count, $intuneRows.Count, $enrichedCount, $intuneOnlyCount, $cmdbRows.Count
) -InformationAction Continue
[pscustomobject]@{
    Status = 'Completed'; ScriptVersion = $ScriptVersion
    EntraDeviceCount = $entraRows.Count; IntuneManagedDeviceCount = $intuneRows.Count
    EnrichedDeviceCount = $enrichedCount; IntuneOnlyDeviceCount = $intuneOnlyCount
    MissingAzureAdDeviceIdCount = $missingAzureAdDeviceIdCount
    DuplicateCorrelationCount = $duplicateCorrelationCount
    DeviceCount = $cmdbRows.Count; EntraRawInputPath = $EntraRawInputPath
    IntuneRawInputPath = $IntuneRawInputPath; CmdbDeviceOutputPath = $cmdbOutputPath
    DimDeviceOutputPath = $dimOutputPath
    FactDeviceComplianceOutputPath = $complianceOutputPath
    FactDeviceComplianceCount = $complianceRows.Count
    RawContractVersion = [string]$rawContract.contractVersion
    CuratedContractVersion = [string]$curatedContract.contractVersion
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAoAWn3k7XJuFq+
# I9qnDBNRkjPE4X5Lph/InTj/uVQdCqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIE21XPD/6vDs4pFtwLnubR1neRXdWdNMsm+6UAbbyF7PMA0GCSqG
# SIb3DQEBAQUABIIBgG/IMWdLbRfVwPjd/np8VfMMjwv9ABosmVgMLJm455bv3Msh
# Qy0S3cgHD0kW18BcQkLv9oXELFupoas3F3tfHlRpVLW9I7vAvudfLNuLpNBwTmuN
# 2FLU/TUPTxIWFndj49Ent0zQ35/wQnkQOIdDvuUpWq4mz/lAAsdc4TrjF/JVMqs3
# 9NAgNYFF5UJWbKymhFAI9a9fz4vJh4GKr/VKQKYC9eV/7/vhxk0Xl2/tLb2h06hi
# NZ6ylnvW5GAKl2wmhaKgK+TKsgRkxjMShBjTqOWljprvn2CDYR1IODhYwIPSDn4T
# mww77O8sKAoJe2DJpodQUHuoHeUOxPmNgCpg7NZ2cF/VleXaD0mXPiB35UcfM7z/
# lMZdIgL9H3BgFiyjfhE02re379ZW4FYTqZiy9b1z2aFfiwISKqxqWzts7aRTPzZg
# xUlnuttKrAKg9O1SmpZIdQaKLo6OsOKQO6zfoqtonroLjpX2oULJhreNFOzo9HJg
# N0dJgbviY1yEAXR3J6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxODEx
# MTJaMC8GCSqGSIb3DQEJBDEiBCAUt1u+6SeLmUtSNhWzaPAwwSrMNttyjmZmxmBq
# Lm5dvTANBgkqhkiG9w0BAQEFAASCAgBbfE5V3k8SjUvaeJIxkI5g1IY5Wl0usBcT
# iDGt3MK7/7CRubBOYZppnB3pNQ3AUGfDd994FZdG8ttYmCa5mp3qIMR3FpRYIJk7
# Zqk2SjRpGC9IBUR/2u9YkiMvHu7sLcIPmEfVguKz7c5fcf+SjjXOpkEE9NoQUhnm
# sCoiipxyY8Xj9hvy+wP4DBdH+dpQByV3vIpd32Eh07jnx1BEXooUZivTBL3G9lae
# qgphV3RuLAvcjOz9b0i2Wa8ear93Dk17YQgE/9r+wgne/jThVpy14vjgBXWHezPF
# zTlCTyakVExR5q8EkBluTrLevqg7P8gK6F3CSASAy23GAsIOOgjyopBsTAQ2XNAB
# M4Jm/ZI9HVXVXXhgfUOmmVmIl+zd0r6gG8MPOEX2tkU281liAeETv+1IOaJnZgJp
# 9i5kVmkYroQCA01b+32TGDp2XdU9kra4sAhT/0JjdQY/6LSOboxjD4m6T2jU8c+2
# z/BXVvdh6DwEQoqvh+NSmE5ay3KsrGuNMvO4D8TZ1gSYPb2HIYqxIEFgrJZXMqBc
# 3wmMvEEkrefPJRVquf1WOPD/cqEyNplbhgimKgUpvHzO0n/czQNbwQ2eI5jk6O7R
# r9BTKVdn9LuMz4DwyYey61LHemyYp6Pti+VDcON5m7DKKZ7rkQhHiOI0GCakSyvG
# 4Z5FEDbnbg==
# SIG # End signature block
