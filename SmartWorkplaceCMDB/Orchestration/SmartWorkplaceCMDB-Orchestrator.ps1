<#
.SYNOPSIS
Runs autonomous SmartWorkplaceCMDB collection and curation pipelines.

.DESCRIPTION
Coordinates source collection, normalization, relationship consolidation,
data-quality publication, tenant/date dimensions, contract build, and the local
HTML report. The default mode is read-only validation. Live collection requires
the explicit -Collect switch. Offline fixture runs never connect to a tenant.

.VERSION
0.2.0
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
    [ValidateSet(
        'Full',
        'EntraUsers',
        'EntraGroups',
        'EntraDevices',
        'IntuneDevices',
        'M365SubscribedSkus',
        'M365UserLicenses',
        'ExchangeOnlineMailboxes',
        'ActiveDirectory',
        'CuratedOnly'
    )]
    [string]$Pipeline = 'Full',
    [switch]$Collect,
    [switch]$ValidateOnly,
    [switch]$ValidateExistingOutputs,
    [string]$FixtureRootPath,
    [ValidateRange(0, 2147483647)]
    [int]$MaxItems = 0,
    [switch]$NoConfigWrite,
    [switch]$DisableSharePointUpload
)

$ScriptVersion = '0.2.0'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Add-SmartWorkplaceCMDBOrchestratorStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$List,
        [Parameter(Mandatory)][hashtable]$Catalog,
        [Parameter(Mandatory)][string[]]$Names
    )
    foreach ($name in $Names) {
        if (-not $Catalog.ContainsKey($name)) {
            throw "Unknown orchestrator step '$name'."
        }
        $List.Add($Catalog[$name])
    }
}

function Get-SmartWorkplaceCMDBStepParameter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Step,
        [Parameter(Mandatory)][hashtable]$CommonParameters,
        [Parameter(Mandatory)][string]$Mode,
        [string]$FixtureRootPath,
        [int]$MaxItems
    )
    $parameters = @{} + $CommonParameters
    if ($Mode -eq 'Validate') {
        $parameters['ValidateOnly'] = $true
    }
    if ($Step.Kind -eq 'Collect') {
        if ($Mode -eq 'Fixture' -or
            ($Mode -eq 'Validate' -and
                -not [string]::IsNullOrWhiteSpace($FixtureRootPath))) {
            $parameters['InputJsonPath'] = Join-Path `
                $FixtureRootPath `
                $Step.FixtureName
        }
        if ($Mode -ne 'Validate' -and $MaxItems -gt 0) {
            $parameters['MaxItems'] = $MaxItems
        }
    }
    return $parameters
}

function Get-SmartWorkplaceCMDBOrchestratorSetting {
    [CmdletBinding()]
    param(
        [AllowNull()]$Configuration,
        [Parameter(Mandatory)][string]$Name,
        $DefaultValue
    )

    if ($null -eq $Configuration) {
        return $DefaultValue
    }
    if ($Configuration -is [System.Collections.IDictionary] -and
        $Configuration.Contains($Name)) {
        return $Configuration[$Name]
    }
    $property = $Configuration.PSObject.Properties[$Name]
    if ($null -ne $property) {
        return $property.Value
    }
    return $DefaultValue
}

function Get-SmartWorkplaceCMDBCsvSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$RootPath)

    $snapshot = @{}
    foreach ($root in $RootPath) {
        if ([string]::IsNullOrWhiteSpace($root) -or
            -not (Test-Path -LiteralPath $root -PathType Container)) {
            continue
        }
        foreach ($file in @(Get-ChildItem -LiteralPath $root -Filter '*.csv' -File -Recurse)) {
            $snapshot[$file.FullName] = '{0}|{1}' -f
                $file.Length,
                $file.LastWriteTimeUtc.Ticks
        }
    }
    return $snapshot
}

function Write-SmartWorkplaceCMDBOrchestratorLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Results,
        [Parameter(Mandatory)][string]$Path
    )
    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    @($Results.ToArray()) |
        Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

if ($Collect -and $ValidateOnly) {
    throw '-Collect and -ValidateOnly cannot be used together.'
}
if ($Collect -and -not [string]::IsNullOrWhiteSpace($FixtureRootPath)) {
    throw '-Collect cannot be combined with -FixtureRootPath.'
}
if ($MaxItems -gt 0 -and $Pipeline -in @('Full', 'CuratedOnly')) {
    throw '-MaxItems requires an individual source pipeline.'
}

$mode = if (-not [string]::IsNullOrWhiteSpace($FixtureRootPath) -and
    -not $ValidateOnly) {
    'Fixture'
}
elseif ($Collect) {
    'Collect'
}
else {
    'Validate'
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$coreModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
Import-Module $coreModulePath -Force
$sharePointModulePath = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.SharePoint\SmartWorkplaceCMDB.SharePoint.psd1'
Import-Module $sharePointModulePath -Force

if (-not [string]::IsNullOrWhiteSpace($FixtureRootPath)) {
    $FixtureRootPath = [IO.Path]::GetFullPath($FixtureRootPath)
    if (-not (Test-Path -LiteralPath $FixtureRootPath -PathType Container)) {
        throw "FixtureRootPath does not exist: '$FixtureRootPath'."
    }
}

if ($mode -eq 'Collect' -and
    $MaxItems -gt 0 -and
    [string]::IsNullOrWhiteSpace($DataRootPath)) {
    $boundedName = '{0}_MAXITEMS-{1}_{2}' -f
        $Tenant,
        $MaxItems,
        ([datetime]::UtcNow.ToString('yyyyMMdd-HHmmssfff'))
    $DataRootPath = Join-Path $projectRoot (
        Join-Path 'Data\TestRuns' $boundedName
    )
    Write-Information (
        "Bounded live output is isolated from canonical data: '{0}'." -f
        $DataRootPath
    ) -InformationAction Continue
}

$boundParameterCopy = @{}
foreach ($key in @(
        'Tenant',
        'OrganizationKey',
        'EnvironmentKey',
        'TenantKey',
        'TenantId',
        'DataRootPath',
        'DataAllRootPath',
        'LatestOutputRootPath',
        'LogRootPath'
    )) {
    $value = Get-Variable -Name $key -ValueOnly
    if ($key -eq 'Tenant' -or
        -not [string]::IsNullOrWhiteSpace([string]$value)) {
        $boundParameterCopy[$key] = $value
    }
}
$context = Resolve-SmartWorkplaceCMDBContext `
    -BoundParameters $boundParameterCopy `
    -GlobalConfigPath $GlobalConfigPath `
    -TenantConfigPath $TenantConfigPath `
    -NoConfigWrite:($mode -ne 'Collect' -or $NoConfigWrite)
$paths = $context.Paths
$sharePointConfiguration = Get-SmartWorkplaceCMDBOrchestratorSetting `
    -Configuration $context.Configuration `
    -Name 'SharePoint' `
    -DefaultValue $null
$graphConfiguration = Get-SmartWorkplaceCMDBOrchestratorSetting `
    -Configuration $context.Configuration `
    -Name 'MicrosoftGraph' `
    -DefaultValue $null
$sharePointEnabled = [bool](Get-SmartWorkplaceCMDBOrchestratorSetting `
        -Configuration $sharePointConfiguration `
        -Name 'Enabled' `
        -DefaultValue $false)
$sharePointEligible = $mode -eq 'Collect' -and
    $MaxItems -eq 0 -and
    -not $DisableSharePointUpload -and
    $sharePointEnabled
$sharePointBeforeSnapshot = if ($sharePointEligible) {
    Get-SmartWorkplaceCMDBCsvSnapshot -RootPath @(
        $paths.DataAllRootPath,
        $paths.LatestOutputRootPath,
        $paths.LogRootPath
    )
}
else {
    @{}
}

$commonParameters = @{
    Tenant = $Tenant
}
foreach ($entry in @(
        @{ Name = 'OrganizationKey'; Value = $OrganizationKey },
        @{ Name = 'EnvironmentKey'; Value = $EnvironmentKey },
        @{ Name = 'TenantKey'; Value = $TenantKey },
        @{ Name = 'TenantId'; Value = $TenantId },
        @{ Name = 'DataRootPath'; Value = $DataRootPath },
        @{ Name = 'DataAllRootPath'; Value = $DataAllRootPath },
        @{ Name = 'LatestOutputRootPath'; Value = $LatestOutputRootPath },
        @{ Name = 'LogRootPath'; Value = $LogRootPath },
        @{ Name = 'GlobalConfigPath'; Value = $GlobalConfigPath },
        @{ Name = 'TenantConfigPath'; Value = $TenantConfigPath }
    )) {
    if (-not [string]::IsNullOrWhiteSpace([string]$entry.Value)) {
        $commonParameters[$entry.Name] = $entry.Value
    }
}
if ($mode -ne 'Collect' -or $NoConfigWrite) {
    $commonParameters['NoConfigWrite'] = $true
}

$catalog = @{}
$catalog['EntraUsersCollect'] = [pscustomobject]@{
    Name = 'Entra users collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Collect.ps1'
    FixtureName = 'EntraUsers.sample.json'
}
$catalog['EntraUsersNormalize'] = [pscustomobject]@{
    Name = 'Entra users normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Normalize.ps1'
    FixtureName = ''
}
$catalog['EntraGroupsCollect'] = [pscustomobject]@{
    Name = 'Entra groups collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraGroups-Collect.ps1'
    FixtureName = 'EntraGroups.sample.json'
}
$catalog['EntraGroupsNormalize'] = [pscustomobject]@{
    Name = 'Entra groups normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraGroups-Normalize.ps1'
    FixtureName = ''
}
$catalog['EntraDevicesCollect'] = [pscustomobject]@{
    Name = 'Entra devices collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraDevices-Collect.ps1'
    FixtureName = 'EntraDevices.sample.json'
}
$catalog['EntraDevicesNormalize'] = [pscustomobject]@{
    Name = 'Entra devices normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Entra\SmartWorkplaceCMDB-EntraDevices-Normalize.ps1'
    FixtureName = ''
}
$catalog['IntuneDevicesCollect'] = [pscustomobject]@{
    Name = 'Intune managed devices collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneManagedDevices-Collect.ps1'
    FixtureName = 'IntuneManagedDevices.sample.json'
}
$catalog['IntuneDevicesNormalize'] = [pscustomobject]@{
    Name = 'Intune device enrichment'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneDevices-Normalize.ps1'
    FixtureName = ''
}
$catalog['UserDeviceRelationships'] = [pscustomobject]@{
    Name = 'Primary user-device relationships'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneUserDeviceRelationships-Normalize.ps1'
    FixtureName = ''
}
$catalog['M365SkusCollect'] = [pscustomobject]@{
    Name = 'Microsoft 365 subscribed SKUs collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Collect.ps1'
    FixtureName = 'M365SubscribedSkus.sample.json'
}
$catalog['M365SkusNormalize'] = [pscustomobject]@{
    Name = 'Microsoft 365 subscribed SKUs normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Normalize.ps1'
    FixtureName = ''
}
$catalog['M365LicensesCollect'] = [pscustomobject]@{
    Name = 'Microsoft 365 user licenses collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Collect.ps1'
    FixtureName = 'M365UserLicenseAssignments.sample.json'
}
$catalog['M365LicensesNormalize'] = [pscustomobject]@{
    Name = 'Microsoft 365 user licenses normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Normalize.ps1'
    FixtureName = ''
}
$catalog['ExchangeMailboxesCollect'] = [pscustomobject]@{
    Name = 'Exchange Online mailboxes collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\ExchangeOnline\SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Collect.ps1'
    FixtureName = 'ExchangeOnlineMailboxes.sample.json'
}
$catalog['ExchangeMailboxesNormalize'] = [pscustomobject]@{
    Name = 'Exchange Online mailboxes normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\ExchangeOnline\SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Normalize.ps1'
    FixtureName = ''
}
$catalog['ActiveDirectoryCollect'] = [pscustomobject]@{
    Name = 'Active Directory collection'
    Kind = 'Collect'
    ScriptPath = Join-Path $projectRoot 'Collectors\ActiveDirectory\SmartWorkplaceCMDB-ActiveDirectory-Collect.ps1'
    FixtureName = 'ActiveDirectory.sample.json'
}
$catalog['ActiveDirectoryNormalize'] = [pscustomobject]@{
    Name = 'Active Directory normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\ActiveDirectory\SmartWorkplaceCMDB-ActiveDirectory-Normalize.ps1'
    FixtureName = ''
}
$catalog['Relationships'] = [pscustomobject]@{
    Name = 'General relationship consolidation'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\SmartWorkplaceCMDB-Relationships-Normalize.ps1'
    FixtureName = ''
}
$catalog['DataQuality'] = [pscustomobject]@{
    Name = 'Data-quality normalization'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\SmartWorkplaceCMDB-DataQuality-Normalize.ps1'
    FixtureName = ''
}
$catalog['Dimensions'] = [pscustomobject]@{
    Name = 'Tenant and date dimensions'
    Kind = 'Normalize'
    ScriptPath = Join-Path $projectRoot 'Collectors\SmartWorkplaceCMDB-Dimensions-Normalize.ps1'
    FixtureName = ''
}
$catalog['Build'] = [pscustomobject]@{
    Name = 'Contract build and manifest'
    Kind = 'Build'
    ScriptPath = Join-Path $projectRoot 'Build\SmartWorkplaceCMDB-Build.ps1'
    FixtureName = ''
}
$catalog['Report'] = [pscustomobject]@{
    Name = 'Local HTML overview report'
    Kind = 'Report'
    ScriptPath = Join-Path $projectRoot 'Reports\SmartWorkplaceCMDB-Report.ps1'
    FixtureName = ''
}

$selectedSteps = New-Object System.Collections.Generic.List[object]
switch ($Pipeline) {
    'Full' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'EntraUsersCollect', 'EntraUsersNormalize',
            'EntraGroupsCollect', 'EntraGroupsNormalize',
            'EntraDevicesCollect', 'EntraDevicesNormalize',
            'ActiveDirectoryCollect', 'ActiveDirectoryNormalize',
            'IntuneDevicesCollect', 'IntuneDevicesNormalize',
            'UserDeviceRelationships',
            'M365SkusCollect', 'M365SkusNormalize',
            'M365LicensesCollect', 'M365LicensesNormalize',
            'ExchangeMailboxesCollect', 'ExchangeMailboxesNormalize',
            'Relationships', 'DataQuality', 'Dimensions', 'Build', 'Report'
        )
    }
    'EntraUsers' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'EntraUsersCollect', 'EntraUsersNormalize'
        )
    }
    'EntraGroups' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'EntraGroupsCollect', 'EntraGroupsNormalize'
        )
    }
    'EntraDevices' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'EntraDevicesCollect', 'EntraDevicesNormalize'
        )
    }
    'IntuneDevices' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'IntuneDevicesCollect', 'IntuneDevicesNormalize',
            'UserDeviceRelationships'
        )
    }
    'M365SubscribedSkus' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'M365SkusCollect', 'M365SkusNormalize'
        )
    }
    'M365UserLicenses' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'M365LicensesCollect', 'M365LicensesNormalize'
        )
    }
    'ExchangeOnlineMailboxes' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'ExchangeMailboxesCollect', 'ExchangeMailboxesNormalize'
        )
    }
    'ActiveDirectory' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'ActiveDirectoryCollect', 'ActiveDirectoryNormalize'
        )
    }
    'CuratedOnly' {
        Add-SmartWorkplaceCMDBOrchestratorStep $selectedSteps $catalog @(
            'Relationships', 'DataQuality', 'Dimensions', 'Build', 'Report'
        )
    }
}

foreach ($step in @($selectedSteps.ToArray())) {
    if (-not (Test-Path -LiteralPath $step.ScriptPath -PathType Leaf)) {
        throw "Orchestrator step script is missing: '$($step.ScriptPath)'."
    }
    if (($mode -eq 'Fixture' -or
            ($mode -eq 'Validate' -and
                -not [string]::IsNullOrWhiteSpace($FixtureRootPath))) -and
        $step.Kind -eq 'Collect') {
        $fixturePath = Join-Path $FixtureRootPath $step.FixtureName
        if (-not (Test-Path -LiteralPath $fixturePath -PathType Leaf)) {
            throw "Required fixture is missing: '$fixturePath'."
        }
    }
}

$executionSteps = @($selectedSteps.ToArray())
if ($mode -eq 'Validate' -and
    -not $ValidateExistingOutputs -and
    $Pipeline -ne 'CuratedOnly') {
    $executionSteps = @($executionSteps |
        Where-Object Kind -in @('Collect', 'Build', 'Report'))
}

$runId = [guid]::NewGuid().ToString('N')
$runStarted = [datetimeoffset]::UtcNow
$results = New-Object System.Collections.Generic.List[object]
$logPath = if ($mode -eq 'Validate') {
    ''
}
else {
    Join-Path $paths.LogRootPath (
        Join-Path 'Orchestration' (
            'SmartWorkplaceCMDB-Orchestrator_{0}.csv' -f
            $runStarted.ToString('yyyyMMdd-HHmmssfff')
        )
    )
}

$failedMessage = ''
try {
    for ($index = 0; $index -lt $executionSteps.Count; $index++) {
        $step = $executionSteps[$index]
        $sequence = $index + 1
        $elapsed = [datetimeoffset]::UtcNow - $runStarted
        $etaText = 'estimating'
        if ($index -gt 0) {
            $averageSeconds = $elapsed.TotalSeconds / $index
            $remainingSeconds = $averageSeconds * (
                $executionSteps.Count - $index
            )
            $etaText = [timespan]::FromSeconds(
                [math]::Max(0, $remainingSeconds)
            ).ToString('hh\:mm\:ss')
        }
        Write-Information (
            "[{0}/{1}] {2} (elapsed {3}; ETA {4})" -f
            $sequence,
            $executionSteps.Count,
            $step.Name,
            $elapsed.ToString('hh\:mm\:ss'),
            $etaText
        ) -InformationAction Continue

        $stepStarted = [datetimeoffset]::UtcNow
        $status = 'Completed'
        $errorText = ''
        try {
            $parameters = Get-SmartWorkplaceCMDBStepParameter `
                -Step $step `
                -CommonParameters $commonParameters `
                -Mode $mode `
                -FixtureRootPath $FixtureRootPath `
                -MaxItems $MaxItems
            @(& $step.ScriptPath @parameters) | Out-Null
            if ($mode -eq 'Validate') {
                $status = 'Validated'
            }
        }
        catch {
            $status = 'Failed'
            $errorText = $_.Exception.Message
            throw
        }
        finally {
            $stepEnded = [datetimeoffset]::UtcNow
            $results.Add([pscustomobject][ordered]@{
                RunId = $runId
                Sequence = $sequence
                Pipeline = $Pipeline
                Mode = $mode
                Step = $step.Name
                Status = $status
                StartedDateTime = $stepStarted.ToString('o')
                EndedDateTime = $stepEnded.ToString('o')
                DurationSeconds = [math]::Round(
                    ($stepEnded - $stepStarted).TotalSeconds,
                    3
                )
                Error = $errorText
            })
        }
    }
}
catch {
    $failedMessage = $_.Exception.Message
}
finally {
    if ($mode -ne 'Validate') {
        Write-SmartWorkplaceCMDBOrchestratorLog `
            -Results $results `
            -Path $logPath
    }
}

if (-not [string]::IsNullOrWhiteSpace($failedMessage)) {
    throw "SmartWorkplaceCMDB orchestration failed: $failedMessage"
}

$sharePointRecords = @()
$sharePointError = ''
if ($sharePointEligible) {
    try {
        $afterSnapshot = Get-SmartWorkplaceCMDBCsvSnapshot -RootPath @(
            $paths.DataAllRootPath,
            $paths.LatestOutputRootPath,
            $paths.LogRootPath
        )
        $changedFiles = @($afterSnapshot.Keys | Where-Object {
                -not $sharePointBeforeSnapshot.ContainsKey($_) -or
                $sharePointBeforeSnapshot[$_] -ne $afterSnapshot[$_]
            } | Sort-Object)
        if ($changedFiles.Count -gt 0) {
            $sharePointRecords = @(Publish-SmartWorkplaceCMDBSharePointFile `
                    -LocalFilePath $changedFiles `
                    -DataAllRootPath $paths.DataAllRootPath `
                    -LatestOutputRootPath $paths.LatestOutputRootPath `
                    -LogRootPath $paths.LogRootPath `
                    -TenantId ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $graphConfiguration 'TenantId' '')) `
                    -ClientId ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $graphConfiguration 'ClientId' '')) `
                    -CertificateThumbprint ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $graphConfiguration 'CertificateThumbprint' '')) `
                    -SiteHostname ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $sharePointConfiguration 'SiteHostname' '')) `
                    -SitePath ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $sharePointConfiguration 'SitePath' '')) `
                    -LibraryDisplayName ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $sharePointConfiguration 'LibraryDisplayName' 'Documents')) `
                    -TargetFolderPath ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                            $sharePointConfiguration 'TargetFolderPath' 'SMART-CMDB/DATA')))
        }
    }
    catch {
        $sharePointError = $_.Exception.Message
        Write-Warning "SmartWorkplaceCMDB SharePoint publication failed but collection outputs are preserved: $sharePointError"
    }
}
elseif ($mode -eq 'Collect' -and $MaxItems -gt 0 -and $sharePointEnabled) {
    Write-Information 'SharePoint publication skipped for the bounded MaxItems run.' `
        -InformationAction Continue
}

$sharePointFailureCount = @($sharePointRecords |
    Where-Object Status -eq 'Failed').Count
$sharePointUploadCount = @($sharePointRecords |
    Where-Object Status -eq 'Uploaded').Count
if ($sharePointEligible) {
    Write-Information (
        'SmartWorkplaceCMDB SharePoint publication completed. Uploaded={0}; Failed={1}; Target={2}.' -f
        $sharePointUploadCount,
        ($sharePointFailureCount + [int](-not [string]::IsNullOrWhiteSpace($sharePointError))),
        ([string](Get-SmartWorkplaceCMDBOrchestratorSetting `
                $sharePointConfiguration 'TargetFolderPath' 'SMART-CMDB/DATA'))
    ) -InformationAction Continue
}

$runEnded = [datetimeoffset]::UtcNow
$runStatus = if ($mode -eq 'Validate') {
    'Validated'
}
elseif ($sharePointFailureCount -gt 0 -or
    -not [string]::IsNullOrWhiteSpace($sharePointError)) {
    'CompletedWithWarnings'
}
else {
    'Completed'
}
Write-Information (
    "SmartWorkplaceCMDB orchestration {0}. Pipeline={1}; Mode={2}; Steps={3}; Duration={4}." -f
    $runStatus.ToLowerInvariant(),
    $Pipeline,
    $mode,
    $results.Count,
    ($runEnded - $runStarted).ToString('hh\:mm\:ss')
) -InformationAction Continue

[pscustomobject]@{
    Status = $runStatus
    ScriptVersion = $ScriptVersion
    RunId = $runId
    Pipeline = $Pipeline
    Mode = $mode
    StepCount = $results.Count
    FailedStepCount = @($results |
        Where-Object Status -eq 'Failed').Count
    StartedDateTime = $runStarted.ToString('o')
    EndedDateTime = $runEnded.ToString('o')
    DurationSeconds = [math]::Round(
        ($runEnded - $runStarted).TotalSeconds,
        3
    )
    DataRootPath = $paths.DataRootPath
    LatestOutputRootPath = $paths.LatestOutputRootPath
    LogPath = $logPath
    SharePointEnabled = $sharePointEnabled
    SharePointEligible = $sharePointEligible
    SharePointUploadCount = $sharePointUploadCount
    SharePointFailureCount = $sharePointFailureCount +
        [int](-not [string]::IsNullOrWhiteSpace($sharePointError))
    SharePointTargetFolderPath = [string](
        Get-SmartWorkplaceCMDBOrchestratorSetting `
            $sharePointConfiguration `
            'TargetFolderPath' `
            'SMART-CMDB/DATA'
    )
    SharePointError = $sharePointError
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCyLEBL/WV2E6fn
# 5ZNqg/M7u4O7x8av7lsZt3bJVWfw9qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINh5H7fT2OdBOM3SrTdGVJl/5gkoKQlEje0BTp2kzdbrMA0GCSqG
# SIb3DQEBAQUABIIBgGi08V6aBooAqA4l5BwoJF3/fASRCfGxhaLr7B25+w/Wqy0e
# RLsZrDZlWO8fo+ZPutVJ+Qr94ne0Rf9YSYUgwE2VQ9zTBmJqFRzmCdpiu7BPhsY1
# 1yIUlvFANRRGeGF4CH7VHn9NsCP2ZN6hxM+I1TjCdY+Mb/U1euSkMv9ZGRbqER7s
# dX3WJUeWBxfb+TeCn/Fxo6deF94Tg3SVmhR7ynEXHybYpF642G8SydbGWcQk/UZd
# GG0sCYVoPnh76Yy0VJzx6n+A//aFlXVM8s/cnFUBeroObNDsNvET5guY3nAiL4hF
# GHrttC1C1wgC4qKw119yNK99MtCaPcn5OHfU3J6D/zWw2fXNAmREC6QNFBKpn+ba
# Rz1TtPpSbFGb1Nb+r70QJSil8gP+ZBNQzhjbbhSj4iaZfi8rXiyq3EtvSlNEjH89
# Hl1Voq5Uahd/kvAN3BdhmNIUPcxl3IccIWNknbn6BC7MbLSTzQ1CZ31dQF04GJ82
# PYNa4Q+49wm4J+PsDqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMjAy
# MDNaMC8GCSqGSIb3DQEJBDEiBCBzZGfSvgk9yXSwfM5rCo9fWLyxW4cJMEN6Ei8f
# B4G/4DANBgkqhkiG9w0BAQEFAASCAgBnEp6oCcrYiAa/aYAy06apO4kWmkWtOg90
# PAk1hr/eFbh5XqrB15FlwIZyZZZpiHMASf3nUhcQlvKo+pwPXwttAXwVpFnsHYJm
# lwvgCTX5959Pmaivn/KEdM0yhmYQlKVdkFCdFcHoRsq6x1QVEUNkj6VBg5vD9p/u
# k29MoDBYkfQ6oADvPpxhGWYWVdJ7N0xJTWo63wYvmo+vR16+t9NkgmoFqet75NB0
# +ZO6+SOrxWlE9HGUMhRlDWVITXToKQlN/dxpgQM4ClTbkqYCfFo4/hP5UiK1072J
# uCsk0pQKss1vXXKs52Z3XkpWp8gWZIKPhfpTNPfrj/7e0vaK4DyuRdReo8iDCqco
# QGGQqxV3E1TLY9gqPkGWYbiyyprASaG/S30H4L3FB9Cyh3NqMHjJqHjEt9gRiL5t
# DR5J1XyosOj4MoCuoeBU/wAq/DdhZSxdZZhQdBcScJZarjA32g1fj5u0jDB/WYVy
# Ik48/9j381y+/SS6ziqp7K6ZkHxh87YkxpljhvTJT2KsBKGPXb+k7fDnGo5cKh1D
# /8RX7tJ/ldz565GBZ1/1Od9+VYytaXFvu2AHfV3mq7u+SgTE6N9sy0NOMd53FM/s
# O+51B9OHmo+Cq+7rg+pQMotwbb5RJOfzxeU+CYwfgLl9xNeyPquHK514hQWv9S0Z
# Nhg6S/nt5w==
# SIG # End signature block
