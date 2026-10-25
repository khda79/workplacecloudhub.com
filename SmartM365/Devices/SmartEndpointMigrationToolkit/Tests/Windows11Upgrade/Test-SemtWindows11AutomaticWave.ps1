<#
.SYNOPSIS
Offline synthetic tests of the Windows11Upgrade automatic Wave engine.

.DESCRIPTION
Copies the package (without Runs, Wave contents, Inventory and DemoAssets) into a fresh
temporary folder, isolates the technician runtime with SEMT_RUNTIME_ROOT_OVERRIDE and runs
the Windows11Upgrade automatic Wave engine against synthetic AD and Intune CSV files.
The repository folder is never written. No network, AD or Microsoft Graph access is used.
Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M"; exit code 1 on
any failure. Compatible with Windows PowerShell 5.1 and PowerShell 7.
Provenance: assertions ported from SmartM365/Devices/Windows11UpgradeToolkit/Tests/Test-SmartM365Windows11AutomaticLot.ps1 1.4.0 (engine rules only).

.VERSION
1.0.0

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Passed = 0
$script:Failed = 0

function Write-TestLine {
    param([string]$Text)
    Write-Host ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Text)
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if ($Condition) { $script:Passed++; Write-TestLine "PASS: $Message" }
    else { $script:Failed++; Write-TestLine "FAIL: $Message" }
}

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    Assert-True -Condition ([string]$Actual -eq [string]$Expected) -Message ('{0} (expected={1}; actual={2})' -f $Message, $Expected, $Actual)
}

function Assert-Throws {
    param([scriptblock]$Script, [string]$Pattern, [string]$Message)
    $thrown = $false
    $detail = ''
    try { & $Script | Out-Null }
    catch { $detail = $_.Exception.Message; $thrown = ($detail -match $Pattern) }
    Assert-True -Condition $thrown -Message ('{0} (error={1})' -f $Message, $detail)
}

function Copy-SemtPackageForTest {
    param([string]$Source, [string]$Destination)
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    foreach ($item in @(Get-ChildItem -LiteralPath $Source -Force)) {
        if ($item.PSIsContainer -and @('Runs', 'Waves', 'Inventory', 'DemoAssets') -contains $item.Name) { continue }
        Copy-Item -LiteralPath $item.FullName -Destination $Destination -Recurse -Force
    }
    foreach ($folder in @('Waves\HybridJoin', 'Waves\Windows11Upgrade', 'Runs')) {
        New-Item -ItemType Directory -Path (Join-Path $Destination $folder) -Force | Out-Null
    }
}

function Get-RepositoryRuntimeSnapshot {
    param([string]$Root)
    $entries = foreach ($folder in @('Runs', 'Waves', 'Inventory')) {
        $path = Join-Path $Root $folder
        if (Test-Path -LiteralPath $path) {
            Get-ChildItem -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue |
                ForEach-Object { '{0}|{1}' -f $_.FullName, $(if ($_.PSIsContainer) { 'D' } else { $_.LastWriteTimeUtc.Ticks }) }
        }
    }
    return (@($entries | Sort-Object) -join "`n")
}

$repositoryPackageRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SEMT-W11-AutomaticWave-{0}' -f [guid]::NewGuid().ToString('N'))
$packageRoot = Join-Path $testRoot 'Package'
$dataRoot = Join-Path $testRoot 'Data'
$evidenceRoot = Join-Path $testRoot 'Evidence'
$previousRuntimeOverride = [Environment]::GetEnvironmentVariable('SEMT_RUNTIME_ROOT_OVERRIDE', 'Process')
$repositorySnapshotBefore = Get-RepositoryRuntimeSnapshot -Root $repositoryPackageRoot

try {
    [Environment]::SetEnvironmentVariable('SEMT_RUNTIME_ROOT_OVERRIDE', (Join-Path $testRoot 'Runtime'), 'Process')
    Copy-SemtPackageForTest -Source $repositoryPackageRoot -Destination $packageRoot
    New-Item -ItemType Directory -Path $dataRoot -Force | Out-Null
    $engine = Join-Path $packageRoot 'Campaigns\Windows11Upgrade\Scripts\SmartM365-EndpointMigration-Windows11Upgrade-New-AutomaticWave.ps1'
    $waveRoot = Join-Path $packageRoot 'Waves\Windows11Upgrade'
    Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force

    $adCsv = Join-Path $dataRoot 'DevicesAD.csv'
    $intuneCsv = Join-Path $dataRoot 'DevicesIntune.csv'
    $filterAdCsv = Join-Path $dataRoot 'FilterDevicesAD.csv'
    $filterIntuneCsv = Join-Path $dataRoot 'FilterDevicesIntune.csv'

    @(
        [pscustomobject]@{ ComputerName = 'FR-PC1'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $true; DNSHostName = 'fr-pc1.one.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'FR-PC2'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $false; DNSHostName = 'fr-pc2.one.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC3'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $true; DNSHostName = 'pc3.one.example'; OperatingSystem = 'Windows 11 Enterprise'; OperatingSystemVersion = '10.0 (22631)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC4'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $true; DNSHostName = 'pc4.one.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'DUP'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $true; DNSHostName = 'dup.one.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'DUP'; ADInventoryPresent = $true; ADDomain = 'two.example'; Enabled = $true; DNSHostName = 'dup.two.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'STALEAD'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $true; DNSHostName = 'stalead.one.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-120).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC8'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $true; DNSHostName = 'pc8.one.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC9'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $false; DNSHostName = 'pc9.one.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'SERVER1'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $true; DNSHostName = 'server1.one.example'; OperatingSystem = 'Windows Server 2022'; OperatingSystemVersion = '10.0 (20348)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
    ) | Export-Csv -LiteralPath $adCsv -NoTypeInformation -Encoding UTF8

    @(
        [pscustomobject]@{ ComputerName = 'PC4'; DeviceName = 'PC4'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = '4'; OperatingSystem = 'Windows'; OSVersion = '10.0.22631.1'; ManagementState = 'managed'; LastSyncDateTime = (Get-Date).AddHours(-1).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC5'; DeviceName = 'PC5'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = '5-new'; OperatingSystem = 'Windows'; OSVersion = '10.0.19045.1'; ManagementState = 'managed'; LastSyncDateTime = (Get-Date).AddHours(-1).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC5'; DeviceName = 'PC5'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = '5-old'; OperatingSystem = 'Windows'; OSVersion = '10.0.19044.1'; ManagementState = 'managed'; LastSyncDateTime = (Get-Date).AddDays(-40).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC6'; DeviceName = 'PC6'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = '6'; OperatingSystem = 'Windows'; OSVersion = '10.0.19045.1'; ManagementState = 'retirePending'; LastSyncDateTime = (Get-Date).AddHours(-1).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC7'; DeviceName = 'PC7'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = '7'; OperatingSystem = 'Linux'; OSVersion = '6.1'; ManagementState = 'managed'; LastSyncDateTime = (Get-Date).AddHours(-1).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC8'; DeviceName = 'PC8'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = '8'; OperatingSystem = 'Windows'; OSVersion = '10.0.19045.1'; ManagementState = 'retirePending'; LastSyncDateTime = (Get-Date).AddHours(-1).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC9'; DeviceName = 'PC9'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = '9'; OperatingSystem = 'Windows'; OSVersion = '10.0.19045.1'; ManagementState = 'managed'; LastSyncDateTime = (Get-Date).AddHours(-1).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC10'; DeviceName = 'PC10'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = '10-new'; OperatingSystem = 'Windows'; OSVersion = '10.0.19045.1'; ManagementState = 'managed'; LastSyncDateTime = (Get-Date).AddHours(-1).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC10'; DeviceName = 'PC10'; IntuneInventoryPresent = $true; IntuneManagedDeviceId = '10-old-w11'; OperatingSystem = 'Windows'; OSVersion = '10.0.22631.1'; ManagementState = 'managed'; LastSyncDateTime = (Get-Date).AddDays(-10).ToString('o') }
    ) | Export-Csv -LiteralPath $intuneCsv -NoTypeInformation -Encoding UTF8

    @(
        [pscustomobject]@{ ComputerName = 'FR-A-RECENT'; ADInventoryPresent = $true; Enabled = $true; DNSHostName = 'fr-a-recent.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'FR-A-OLD'; ADInventoryPresent = $true; Enabled = $true; DNSHostName = 'fr-a-old.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-60).ToString('o') }
        [pscustomobject]@{ ComputerName = 'FR-A-UNKNOWN'; ADInventoryPresent = $true; Enabled = $true; DNSHostName = 'fr-a-unknown.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = '' }
        [pscustomobject]@{ ComputerName = 'BE-A-RECENT'; ADInventoryPresent = $true; Enabled = $true; DNSHostName = 'be-a-recent.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'FR-B-RECENT'; ADInventoryPresent = $true; Enabled = $true; DNSHostName = 'fr-b-recent.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'FR-A-INTUNE'; ADInventoryPresent = $true; Enabled = $true; DNSHostName = 'fr-a-intune.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'FR-A-MULTI'; ADInventoryPresent = $true; Enabled = $true; DNSHostName = 'fr-a-multi.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-60).ToString('o') }
        [pscustomobject]@{ ComputerName = 'FR-A-MULTI'; ADInventoryPresent = $true; Enabled = $true; DNSHostName = 'fr-a-multi.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
    ) | Export-Csv -LiteralPath $filterAdCsv -NoTypeInformation -Encoding UTF8
    @(
        [pscustomobject]@{ DeviceName = 'FR-A-INTUNE'; IntuneInventoryPresent = $true; OperatingSystem = 'Windows'; OSVersion = '10.0.19045.1'; ManagementState = 'managed'; LastSyncDateTime = (Get-Date).AddHours(-1).ToString('o') }
        [pscustomobject]@{ DeviceName = 'OTHER'; IntuneInventoryPresent = $true; OperatingSystem = 'Windows'; OSVersion = '10.0.19045.1'; ManagementState = 'managed'; LastSyncDateTime = (Get-Date).AddHours(-1).ToString('o') }
    ) | Export-Csv -LiteralPath $filterIntuneCsv -NoTypeInformation -Encoding UTF8

    # Default preview.
    $preview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-AUTO-TEST' -EvidenceRoot $evidenceRoot
    Assert-Equal $preview.Summary.SelectedDevices 3 'selected device count'
    Assert-True (-not [bool]$preview.Summary.NameFilterEnabled) 'blank name filter preserves the default selection scope'
    Assert-Equal $preview.Summary.UniqueInventoryDevices 13 'unique inventory device count'
    Assert-Equal $preview.Summary.NameFilterMatchedDevices 13 'blank name filter matches every unique device'
    Assert-Equal $preview.Summary.NameFilterExcludedDevices 0 'blank name filter excludes no device'
    Assert-Equal $preview.Summary.Windows11Excluded 3 'Windows 11 exclusion count, including duplicate historical Intune evidence'
    Assert-Equal $preview.Summary.ADDisabledExcluded 2 'disabled AD exclusion count, including a conflicting Intune Windows 10 record'
    Assert-Equal $preview.Summary.IntuneStateExcluded 2 'Intune management-state exclusion count, including a conflicting AD Windows 10 record'
    Assert-Equal $preview.Summary.ADNameCollisions 1 'AD collision count'
    Assert-Equal $preview.Summary.IntuneDuplicateRowsIgnored 2 'Intune duplicate count'
    Assert-Equal $preview.Summary.UnknownOSExcluded 2 'non-Windows and Windows Server exclusion count'
    Assert-Equal $preview.Summary.ADStaleWarnings 1 'AD stale warning count'
    Assert-True (-not [bool]$preview.Summary.WaveCreated) 'preview does not create a Wave'
    Assert-Equal @(Get-ChildItem -LiteralPath $waveRoot -Directory).Count 0 'preview leaves Waves\Windows11Upgrade empty'
    Assert-Equal @($preview.Summary.PSObject.Properties.Name | Where-Object { $_ -match '(?i)lot|batch' }).Count 0 'summary uses Wave terminology only'
    foreach ($name in @('DevicesAD.csv', 'DevicesIntune.csv', 'AutomaticWaveSelection.csv', 'AutomaticWaveExclusions.csv', 'AutomaticWaveFilterExclusions.csv', 'AutomaticWaveSummary.json')) {
        Assert-True (Test-Path -LiteralPath (Join-Path $preview.Summary.EvidencePath $name) -PathType Leaf) "preview evidence file exists: $name"
    }

    $selectedNames = @($preview.SelectedDevices | Select-Object -ExpandProperty ComputerName)
    Assert-True ($selectedNames -contains 'fr-pc1.one.example') 'AD FQDN is preferred for FR-PC1'
    Assert-True ($selectedNames -contains 'PC5') 'Intune-only PC5 is selected'
    Assert-True ($selectedNames -contains 'stalead.one.example') 'stale AD Windows 10 remains selected'

    # Prefix filter.
    $filteredPreview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-AUTO-FILTER-PREVIEW' -EvidenceRoot $evidenceRoot -ComputerNamePrefix 'fr-'
    Assert-True ([bool]$filteredPreview.Summary.NameFilterEnabled) 'prefix filter is enabled'
    Assert-Equal $filteredPreview.Summary.ComputerNamePrefixes 'FR-' 'prefix filter is normalized case-insensitively'
    Assert-Equal $filteredPreview.Summary.UniqueInventoryDevices 13 'filtered preview keeps the total unique inventory count'
    Assert-Equal $filteredPreview.Summary.NameFilterMatchedDevices 2 'FR prefix matched device count'
    Assert-Equal $filteredPreview.Summary.NameFilterExcludedDevices 11 'FR prefix filtered-out device count'
    Assert-Equal $filteredPreview.Summary.SelectedDevices 1 'FR prefix selected Windows 10 count'
    Assert-Equal $filteredPreview.Summary.ExcludedDevices 1 'FR prefix safety exclusion count'
    Assert-Equal $filteredPreview.Summary.ADDisabledExcluded 1 'FR prefix keeps safety exclusions active'
    Assert-True (@($filteredPreview.SelectedDevices | Select-Object -ExpandProperty ComputerName) -contains 'fr-pc1.one.example') 'FR prefix matches the short name of an AD FQDN'
    Assert-Equal @($filteredPreview.FilterExcludedDevices).Count 11 'filter-excluded rows are returned separately'
    $filterEvidencePath = Join-Path $filteredPreview.Summary.EvidencePath 'AutomaticWaveFilterExclusions.csv'
    Assert-True (Test-Path -LiteralPath $filterEvidencePath -PathType Leaf) 'filter exclusion evidence was created'
    $filterEvidence = @(Import-Csv -LiteralPath $filterEvidencePath)
    Assert-Equal $filterEvidence.Count 11 'filter exclusion evidence row count'
    Assert-True (@($filterEvidence | Where-Object { $_.FilterReason -ne 'COMPUTER_NAME_PREFIX_NOT_MATCHED' }).Count -eq 0) 'filter exclusion evidence uses the expected reason'

    $multiPrefixPreview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-AUTO-MULTI-PREFIX' -ComputerNamePrefix @('fr-', 'stale') -NoEvidence
    Assert-Equal $multiPrefixPreview.Summary.ComputerNamePrefixes 'FR-;STALE' 'multiple prefixes are normalized and reported'
    Assert-Equal $multiPrefixPreview.Summary.NameFilterMatchedDevices 3 'multiple prefixes matched device count'
    Assert-Equal $multiPrefixPreview.Summary.NameFilterExcludedDevices 10 'multiple prefixes filtered-out device count'
    Assert-Equal $multiPrefixPreview.Summary.SelectedDevices 2 'multiple prefixes selected Windows 10 count'

    # Advanced filters.
    $advanced = & $engine -Source Both -AdInventoryCsv $filterAdCsv -IntuneInventoryCsv $filterIntuneCsv -WaveName 'WAVE-AUTO-ADVANCED-FILTERS' -EvidenceRoot $evidenceRoot -ComputerNamePrefix 'FR-' -ComputerNameContains '-A-' -ExcludeIntunePresent -ExcludeStaleAd -AdLastLogonMaxAgeDays 45
    Assert-Equal $advanced.Summary.ComputerNameContains '-A-' 'contains filter is normalized and reported'
    Assert-True ([bool]$advanced.Summary.ExcludeIntunePresent) 'Intune presence filter is enabled'
    Assert-True ([bool]$advanced.Summary.ExcludeStaleAd) 'AD LastLogon filter is enabled'
    Assert-Equal $advanced.Summary.ADLastLogonMaxAgeDays 45 'AD LastLogon threshold is reported'
    Assert-Equal $advanced.Summary.UniqueInventoryDevices 8 'advanced filter unique inventory count'
    Assert-Equal $advanced.Summary.NameFilterMatchedDevices 2 'advanced filters combine with AND semantics'
    Assert-Equal $advanced.Summary.NameFilterExcludedDevices 6 'advanced filters keep one exclusion row per device'
    Assert-Equal $advanced.Summary.SelectedDevices 2 'advanced filters retain recent AD-only devices'
    Assert-Equal $advanced.Summary.PrefixFilterExcluded 2 'prefix filter exclusion count'
    Assert-Equal $advanced.Summary.ContainsFilterExcluded 2 'contains filter exclusion count'
    Assert-Equal $advanced.Summary.IntunePresentFilterExcluded 2 'Intune presence filter exclusion count'
    Assert-Equal $advanced.Summary.ADLastLogonFilterExcluded 3 'AD LastLogon filter exclusion count'
    Assert-Equal $advanced.Summary.ADLastLogonUnknownExcluded 2 'unknown AD LastLogon exclusion count'
    Assert-True (@($advanced.SelectedDevices | Where-Object { $_.ComputerKey -eq 'FR-A-MULTI' }).Count -eq 1) 'newest valid AD LastLogon wins for duplicate AD rows'
    $advancedEvidence = @(Import-Csv -LiteralPath (Join-Path $advanced.Summary.EvidencePath 'AutomaticWaveFilterExclusions.csv'))
    Assert-Equal $advancedEvidence.Count 6 'advanced filter evidence has one row per excluded device'
    $multiReasonEvidence = @($advancedEvidence | Where-Object { $_.ComputerKey -eq 'OTHER' })
    Assert-Equal $multiReasonEvidence.Count 1 'multi-filter exclusion is consolidated into one evidence row'
    Assert-True ($multiReasonEvidence.Count -eq 1 -and $multiReasonEvidence[0].FilterReason -match 'COMPUTER_NAME_PREFIX_NOT_MATCHED.+COMPUTER_NAME_CONTAINS_NOT_MATCHED.+INTUNE_DEVICE_PRESENT.+AD_LAST_LOGON_UNKNOWN') 'multi-filter evidence preserves every exclusion reason'

    # Names.
    $generated = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ComputerNamePrefix @('fr-', 'be-') -NoEvidence
    Assert-True ([string]$generated.Summary.WaveName -match '^WAVE-AUTO-W10-FR-BE-\d{8}-\d{6}$') ('default Wave name is WAVE-AUTO-W10-<prefix>-<timestamp> (actual={0})' -f $generated.Summary.WaveName)
    $sanitized = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'AUTO: W10/FR*?' -NoEvidence
    Assert-Equal $sanitized.Summary.WaveName 'WAVE-AUTO-W10-FR' 'Windows-forbidden directory characters are sanitized from a manual Wave name'
    $reserved = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'CON' -NoEvidence
    Assert-Equal $reserved.Summary.WaveName 'WAVE-CON' 'mandatory WAVE- prefix neutralizes Windows reserved device names'

    # Guards.
    Assert-Throws { & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-AUTO-WILDCARD' -ComputerNamePrefix 'FR-*' -NoEvidence } 'wildcards and regular expressions are not supported' 'wildcard prefix is rejected explicitly'
    Assert-Throws { & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-AUTO-CONTAINS-WILDCARD' -ComputerNameContains '*-A-*' -NoEvidence } 'wildcards and regular expressions are not supported' 'wildcard contains filter is rejected explicitly'
    Assert-Throws { & $engine -Source AD -AdInventoryCsv $adCsv -WaveName 'WAVE-AUTO-INTUNE-FILTER-SOURCE' -ExcludeIntunePresent -NoEvidence } 'requires an available Intune inventory source' 'Intune presence filter requires Intune inventory evidence'
    Assert-Throws { & $engine -Source Intune -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-AUTO-AD-FILTER-SOURCE' -ExcludeStaleAd -NoEvidence } 'requires an available AD inventory source' 'AD LastLogon filter requires AD inventory evidence'
    Assert-Throws { & $engine -Source Both -AdInventoryCsv $adCsv -WaveName 'WAVE-AUTO-MISSING-SOURCE' -NoEvidence } 'source\(s\) missing' 'missing requested source is blocked by default (shared Intune cache absent)'

    $partial = & $engine -Source Both -AdInventoryCsv $adCsv -WaveName 'WAVE-AUTO-PARTIAL' -AllowPartialSource -NoEvidence
    Assert-True ([bool]$partial.Summary.PartialSource) 'explicit partial-source preview is marked partial'
    Assert-Equal $partial.Summary.AvailableSources 'AD' 'partial-source preview reports AD as the only source'
    Assert-Equal $partial.Summary.SelectedDevices 4 'AD-only partial-source selected count'

    # Create through New-SemtWave.
    $progressStages = New-Object System.Collections.Generic.List[string]
    $progressCallback = { param([string]$Stage, [string]$Detail) [void]$progressStages.Add($Stage) }.GetNewClosure()
    $createdOutput = @(& $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-AUTO-TEST' -Create -ProgressCallback $progressCallback)
    Assert-Equal $createdOutput.Count 1 'create returns a single result object'
    $created = $createdOutput[0]
    Assert-Equal $created.Summary.WaveName 'WAVE-AUTO-TEST' 'create keeps the requested Wave name'
    Assert-Equal $created.Summary.WavePath (Join-Path $waveRoot 'WAVE-AUTO-TEST') 'Wave is created under Waves\Windows11Upgrade'
    Assert-True (Test-Path -LiteralPath $created.Summary.ComputersPath -PathType Leaf) 'Computers.txt was created'
    Assert-Equal @(Get-Content -LiteralPath $created.Summary.ComputersPath).Count 3 'Computers.txt row count'
    foreach ($wrapper in @('Start-Wave-Loop.cmd', 'Start-Wave-Once.cmd', 'Start-Wave-Loop-IgnoreRunGuard.cmd', 'Start-Wave-Once-IgnoreRunGuard.cmd')) {
        Assert-True (Test-Path -LiteralPath (Join-Path $created.Summary.WavePath $wrapper) -PathType Leaf) "Wave wrapper exists: $wrapper"
    }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $created.Summary.WavePath 'Windows11UpgradeToolkit.config'))) 'no legacy toolkit configuration file is created'
    $waveConfig = Get-Content -LiteralPath (Join-Path $created.Summary.WavePath 'Wave.local.json') -Raw | ConvertFrom-Json
    Assert-Equal $waveConfig.Source 'Automatic' 'Wave.local.json records Source=Automatic'
    Assert-Equal $waveConfig.Campaign 'Windows11Upgrade' 'Wave.local.json records the Windows11Upgrade Campaign'
    Assert-Equal $waveConfig.EvidencePath $created.Summary.EvidencePath 'Wave.local.json records the evidence path'
    Assert-True ([string]$waveConfig.Description -match 'Source=Both' -and [string]$waveConfig.Description -match 'Selected=3') 'Wave description summarizes the filters'
    $expectedEvidenceRoot = Join-Path $packageRoot 'Runs\Windows11Upgrade\_AutomaticWave'
    Assert-True (([string]$created.Summary.EvidencePath).StartsWith($expectedEvidenceRoot, [System.StringComparison]::OrdinalIgnoreCase)) 'create evidence is written under Runs\Windows11Upgrade\_AutomaticWave by default'
    foreach ($name in @('AutomaticWaveSelection.csv', 'AutomaticWaveExclusions.csv', 'AutomaticWaveFilterExclusions.csv', 'AutomaticWaveSummary.json')) {
        Assert-True (Test-Path -LiteralPath (Join-Path $created.Summary.EvidencePath $name) -PathType Leaf) "create evidence file exists: $name"
    }
    Assert-True ($progressStages -contains 'Creating the automatic Wave...') 'Wave creation stage is reported through the progress callback'
    Assert-True ($progressStages -contains 'Automatic Wave created.') 'automatic Wave completion is reported through the progress callback'

    $filteredCreated = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-AUTO-FILTERED-CREATE' -EvidenceRoot $evidenceRoot -ComputerNamePrefix 'FR-' -Create
    $filteredComputers = @(Get-Content -LiteralPath $filteredCreated.Summary.ComputersPath)
    Assert-Equal $filteredComputers.Count 1 'filtered Computers.txt row count'
    Assert-Equal $filteredComputers[0] 'fr-pc1.one.example' 'filtered Computers.txt contains only the eligible FR device'
    Assert-True (Test-Path -LiteralPath (Join-Path $filteredCreated.Summary.EvidencePath 'AutomaticWaveFilterExclusions.csv') -PathType Leaf) 'filtered create evidence was created'

    Assert-Throws { & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-AUTO-TEST' -Create -NoEvidence } 'SEMT_WAVE_EXISTS' 'creating an existing Wave is refused'
    Assert-Throws { & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -WaveName 'WAVE-EMPTY' -ComputerNamePrefix 'ZZZ' -Create -NoEvidence } 'No eligible Windows 10 device' 'creating a Wave without selected device is refused'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $waveRoot 'WAVE-EMPTY'))) 'refused empty Wave leaves no folder'
    Assert-Equal @(Get-SemtWaves -PackageRoot $packageRoot -Campaign Windows11Upgrade | Where-Object { $_.Launchable }).Count 2 'created Waves are listed as launchable by the engine'

    # Default inputs are the shared caches of the package.
    New-Item -ItemType Directory -Path (Join-Path $packageRoot 'Inventory') -Force | Out-Null
    Copy-Item -LiteralPath $adCsv -Destination (Get-SemtSharedInventoryPath -PackageRoot $packageRoot -Kind AD) -Force
    Copy-Item -LiteralPath $intuneCsv -Destination (Get-SemtSharedInventoryPath -PackageRoot $packageRoot -Kind Intune) -Force
    $defaults = & $engine -NoEvidence
    Assert-Equal $defaults.Summary.SelectedDevices 3 'default inventory inputs are the shared caches'
    Assert-Equal $defaults.Summary.AvailableSources 'AD+Intune' 'shared AD and Intune caches are both used by default'
}
catch {
    $script:Failed++
    Write-TestLine ("FAIL: unexpected error: {0}" -f $_.Exception.Message)
}
finally {
    [Environment]::SetEnvironmentVariable('SEMT_RUNTIME_ROOT_OVERRIDE', $previousRuntimeOverride, 'Process')
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Assert-True ((Get-RepositoryRuntimeSnapshot -Root $repositoryPackageRoot) -eq $repositorySnapshotBefore) 'repository Runs, Waves and Inventory folders are not written'
Write-TestLine ('Summary: Passed={0}; Failed={1}' -f $script:Passed, $script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfWapc7LzvP9sS
# V/hWYzAyTUJisT5V2tu0+Dkw+7n8e6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIH+oG0/D4uNUtY0N2YGSiPCCPVse9b6OTucNP4D2tqtxMA0GCSqG
# SIb3DQEBAQUABIIBgGRtA7DmCxnGvwhzZRu6FEks8hDHN2Tjz51n+O/KWP6i8/Rk
# KkSCWjhztIq0wY5UQYi1HpgFylKsIJJnuV3SHx49eGrpA5EEtasl9TMAzHtQtt19
# CXGJKzW30l1VemSVvuXagE9vXDINHYlZELRVDz4uH0Ut3pwhdD7KKCTXMkke6aUX
# lhJ8PB2b1WWjNP1+vmZENUo3ukhLHj7DJNXqUfSWRUtHb25BkhEryQSmfJA3Ln8v
# lZ3/2zBaM6kdppdbBVbqLsHfmydj9IVdePB5FmJtzMZdQhim6+fJURAva8G7TK10
# +EuBJTiF/TASMSjYJmWjq9AvDBQ7nBiQlPmm89s4AsEk9VkanROj+eiLvOqsYd8j
# FCPEUaO2yOafctiIEggEvgkI3mexjx8dpIojgsEAyWuSUICVn/QxHyYlHRrsYaGj
# LiBJKc+8d6HB76DPZtxd4rt6S9RuCedW/sRbu7cX4ekUXvPvnJGOB6ZtjhTgsuyq
# 0cGQvGTtbiX8Ub1H7qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTlaMC8GCSqGSIb3DQEJBDEiBCAj0Goh7cB8QxAth4AubYeKV9P2Vk/rr1uwbOWH
# LqnTVzANBgkqhkiG9w0BAQEFAASCAgBOjRFKwavkpT/oyOXutE9W/YEHVUzJotmU
# Nwmb5KPPMIqdgYqgqtn4SG8ooiLVnLDYqOGqn+2BIn1XtVBb+r+PvmOOocV8Z8fp
# qiZNb1Pw7qgCFUQBYEnryiFQGLxjlhgLfMaounP9yt8oAO8qJ//bnPz5AclWVB8P
# PKXmOGZndhCCR3G+kn3AKj+YCLkRT5DZjK7h4ONbiph6wYJ6l4pbdgIzGrjetVST
# OGhPrr9mZhJC8/9+aS07QIUG6Nd6fPHBIMrnHFD8nT7M+upeLYPhxDV/44JzdUdl
# ahoXPN4+MYSHTAksIXQYGLhtIT32eovzJG8lF3EG+1G7NfiSG4MXl7zlJn8yxjNa
# ZQ1IrdluxEjoXjQVarQEkUHv6CqM+ODeLEhVf4CbMkiqXdA2U6xMsfJmQPdEkKT6
# w1B2Eo6mhuYSQbU9kuOLFPHK4DZ96HxMrCcUCTSjgEkSnF69bLTl0oo8brR8Vfmj
# b291cC9WsPQfyOwVSGMeviYtIb50upTZOO5FjVKyBZ1A7qXyP7JiJufByw/p0Vq4
# kOy/V+OQh9fSKPQwoMujGenSs7y/t30snBO6IQ3Da0AeKbzyAipDvuXiMgSouRYD
# 66I/Jxh14nXqIwYqW7Sq+amwQgT3ViT2xOUaF1k1zky3o0JybYNi74WNlNtnTVxg
# 2tmLnYcRSw==
# SIG # End signature block
