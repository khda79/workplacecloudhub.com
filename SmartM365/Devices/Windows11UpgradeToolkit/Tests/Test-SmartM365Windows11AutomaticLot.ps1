<#
.SYNOPSIS
Runs synthetic tests for automatic Windows 10 LOT selection.

.VERSION
1.4.0
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-Equal {
    param(
        [Parameter(Mandatory = $true)]$Actual,
        [Parameter(Mandatory = $true)]$Expected,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if ([string]$Actual -ne [string]$Expected) {
        throw ("ASSERT FAILED: {0}. Expected={1}; Actual={2}" -f $Message, $Expected, $Actual)
    }
}

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

$toolkitRoot = Split-Path -Parent $PSScriptRoot
$engine = Join-Path $toolkitRoot 'Scripts\SmartM365-Windows11Upgrade-New-AutomaticLot.ps1'
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SmartM365-W11UT-AutomaticLot-{0}' -f [guid]::NewGuid().ToString('N'))
$testToolkitRoot = Join-Path $testRoot 'Toolkit'
$evidenceRoot = Join-Path $testRoot 'Evidence'
$adCsv = Join-Path $testRoot 'DevicesAD.csv'
$intuneCsv = Join-Path $testRoot 'DevicesIntune.csv'
$filterAdCsv = Join-Path $testRoot 'FilterDevicesAD.csv'
$filterIntuneCsv = Join-Path $testRoot 'FilterDevicesIntune.csv'

try {
    New-Item -ItemType Directory -Path $testToolkitRoot -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $testToolkitRoot 'Windows11UpgradeToolkit.config.template') -Encoding ASCII -Value @(
        '# Synthetic automatic LOT test configuration.'
        'W11UT_SETUP_SOURCE='
    )

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

    $preview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-TEST' -EvidenceRoot $evidenceRoot
    Assert-Equal -Actual $preview.Summary.SelectedDevices -Expected 3 -Message 'selected device count'
    Assert-True -Condition (-not [bool]$preview.Summary.NameFilterEnabled) -Message 'blank name filter preserves the default selection scope'
    Assert-Equal -Actual $preview.Summary.UniqueInventoryDevices -Expected 13 -Message 'unique inventory device count'
    Assert-Equal -Actual $preview.Summary.NameFilterMatchedDevices -Expected 13 -Message 'blank name filter matches every unique device'
    Assert-Equal -Actual $preview.Summary.NameFilterExcludedDevices -Expected 0 -Message 'blank name filter excludes no device'
    Assert-Equal -Actual $preview.Summary.Windows11Excluded -Expected 3 -Message 'Windows 11 exclusion count, including duplicate historical Intune evidence'
    Assert-Equal -Actual $preview.Summary.ADDisabledExcluded -Expected 2 -Message 'disabled AD exclusion count, including a conflicting Intune Windows 10 record'
    Assert-Equal -Actual $preview.Summary.IntuneStateExcluded -Expected 2 -Message 'Intune management-state exclusion count, including a conflicting AD Windows 10 record'
    Assert-Equal -Actual $preview.Summary.ADNameCollisions -Expected 1 -Message 'AD collision count'
    Assert-Equal -Actual $preview.Summary.IntuneDuplicateRowsIgnored -Expected 2 -Message 'Intune duplicate count'
    Assert-Equal -Actual $preview.Summary.UnknownOSExcluded -Expected 2 -Message 'non-Windows and Windows Server exclusion count'
    Assert-Equal -Actual $preview.Summary.ADStaleWarnings -Expected 1 -Message 'AD stale warning count'

    $selectedNames = @($preview.SelectedDevices | Select-Object -ExpandProperty ComputerName)
    Assert-True -Condition ($selectedNames -contains 'fr-pc1.one.example') -Message 'AD FQDN is preferred for FR-PC1'
    Assert-True -Condition ($selectedNames -contains 'PC5') -Message 'Intune-only PC5 is selected'
    Assert-True -Condition ($selectedNames -contains 'stalead.one.example') -Message 'stale AD Windows 10 remains selected'

    $filteredPreview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-FILTER-PREVIEW' -EvidenceRoot $evidenceRoot -ComputerNamePrefix 'fr-'
    Assert-True -Condition ([bool]$filteredPreview.Summary.NameFilterEnabled) -Message 'prefix filter is enabled'
    Assert-Equal -Actual $filteredPreview.Summary.ComputerNamePrefixes -Expected 'FR-' -Message 'prefix filter is normalized case-insensitively'
    Assert-Equal -Actual $filteredPreview.Summary.UniqueInventoryDevices -Expected 13 -Message 'filtered preview keeps the total unique inventory count'
    Assert-Equal -Actual $filteredPreview.Summary.NameFilterMatchedDevices -Expected 2 -Message 'FR prefix matched device count'
    Assert-Equal -Actual $filteredPreview.Summary.NameFilterExcludedDevices -Expected 11 -Message 'FR prefix filtered-out device count'
    Assert-Equal -Actual $filteredPreview.Summary.SelectedDevices -Expected 1 -Message 'FR prefix selected Windows 10 count'
    Assert-Equal -Actual $filteredPreview.Summary.ExcludedDevices -Expected 1 -Message 'FR prefix safety exclusion count'
    Assert-Equal -Actual $filteredPreview.Summary.ADDisabledExcluded -Expected 1 -Message 'FR prefix keeps safety exclusions active'
    Assert-True -Condition (@($filteredPreview.SelectedDevices | Select-Object -ExpandProperty ComputerName) -contains 'fr-pc1.one.example') -Message 'FR prefix matches the short name of an AD FQDN'
    Assert-Equal -Actual $filteredPreview.FilterExcludedDevices.Count -Expected 11 -Message 'filter-excluded rows are returned separately'
    $filterEvidencePath = Join-Path $filteredPreview.Summary.EvidencePath 'AutomaticLotFilterExclusions.csv'
    Assert-True -Condition (Test-Path -LiteralPath $filterEvidencePath -PathType Leaf) -Message 'filter exclusion evidence was created'
    $filterEvidence = @(Import-Csv -LiteralPath $filterEvidencePath)
    Assert-Equal -Actual $filterEvidence.Count -Expected 11 -Message 'filter exclusion evidence row count'
    Assert-True -Condition (@($filterEvidence | Where-Object { $_.FilterReason -ne 'COMPUTER_NAME_PREFIX_NOT_MATCHED' }).Count -eq 0) -Message 'filter exclusion evidence uses the expected reason'

    $multiPrefixPreview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-MULTI-PREFIX' -ComputerNamePrefix @('fr-','stale') -NoEvidence
    Assert-Equal -Actual $multiPrefixPreview.Summary.ComputerNamePrefixes -Expected 'FR-;STALE' -Message 'multiple prefixes are normalized and reported'
    Assert-Equal -Actual $multiPrefixPreview.Summary.NameFilterMatchedDevices -Expected 3 -Message 'multiple prefixes matched device count'
    Assert-Equal -Actual $multiPrefixPreview.Summary.NameFilterExcludedDevices -Expected 10 -Message 'multiple prefixes filtered-out device count'
    Assert-Equal -Actual $multiPrefixPreview.Summary.SelectedDevices -Expected 2 -Message 'multiple prefixes selected Windows 10 count'

    $advancedFilterPreview = & $engine -Source Both -AdInventoryCsv $filterAdCsv -IntuneInventoryCsv $filterIntuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-ADVANCED-FILTERS' -EvidenceRoot $evidenceRoot -ComputerNamePrefix 'FR-' -ComputerNameContains '-A-' -ExcludeIntunePresent -ExcludeStaleAd -AdLastLogonMaxAgeDays 45
    Assert-Equal -Actual $advancedFilterPreview.Summary.ComputerNameContains -Expected '-A-' -Message 'contains filter is normalized and reported'
    Assert-True -Condition ([bool]$advancedFilterPreview.Summary.ExcludeIntunePresent) -Message 'Intune presence filter is enabled'
    Assert-True -Condition ([bool]$advancedFilterPreview.Summary.ExcludeStaleAd) -Message 'AD LastLogon filter is enabled'
    Assert-Equal -Actual $advancedFilterPreview.Summary.ADLastLogonMaxAgeDays -Expected 45 -Message 'AD LastLogon default threshold is 45 days'
    Assert-Equal -Actual $advancedFilterPreview.Summary.UniqueInventoryDevices -Expected 8 -Message 'advanced filter unique inventory count'
    Assert-Equal -Actual $advancedFilterPreview.Summary.NameFilterMatchedDevices -Expected 2 -Message 'advanced filters combine with AND semantics'
    Assert-Equal -Actual $advancedFilterPreview.Summary.NameFilterExcludedDevices -Expected 6 -Message 'advanced filters keep one exclusion row per device'
    Assert-Equal -Actual $advancedFilterPreview.Summary.SelectedDevices -Expected 2 -Message 'advanced filters retain recent AD-only devices'
    Assert-Equal -Actual $advancedFilterPreview.Summary.PrefixFilterExcluded -Expected 2 -Message 'prefix filter exclusion count'
    Assert-Equal -Actual $advancedFilterPreview.Summary.ContainsFilterExcluded -Expected 2 -Message 'contains filter exclusion count'
    Assert-Equal -Actual $advancedFilterPreview.Summary.IntunePresentFilterExcluded -Expected 2 -Message 'Intune presence filter exclusion count'
    Assert-Equal -Actual $advancedFilterPreview.Summary.ADLastLogonFilterExcluded -Expected 3 -Message 'AD LastLogon filter exclusion count'
    Assert-Equal -Actual $advancedFilterPreview.Summary.ADLastLogonUnknownExcluded -Expected 2 -Message 'unknown AD LastLogon exclusion count'
    Assert-True -Condition (@($advancedFilterPreview.SelectedDevices | Where-Object { $_.ComputerKey -eq 'FR-A-MULTI' }).Count -eq 1) -Message 'newest valid AD LastLogon wins for duplicate AD rows'
    $advancedFilterEvidence = @(Import-Csv -LiteralPath (Join-Path $advancedFilterPreview.Summary.EvidencePath 'AutomaticLotFilterExclusions.csv'))
    Assert-Equal -Actual $advancedFilterEvidence.Count -Expected 6 -Message 'advanced filter evidence has one row per excluded device'
    $multiReasonEvidence = @($advancedFilterEvidence | Where-Object { $_.ComputerKey -eq 'OTHER' })
    Assert-Equal -Actual $multiReasonEvidence.Count -Expected 1 -Message 'multi-filter exclusion is consolidated into one evidence row'
    Assert-True -Condition ($multiReasonEvidence[0].FilterReason -match 'COMPUTER_NAME_PREFIX_NOT_MATCHED.+COMPUTER_NAME_CONTAINS_NOT_MATCHED.+INTUNE_DEVICE_PRESENT.+AD_LAST_LOGON_UNKNOWN') -Message 'multi-filter evidence preserves every exclusion reason'

    $generatedPrefixNamePreview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -ComputerNamePrefix @('fr-','be-') -NoEvidence
    Assert-True -Condition ($generatedPrefixNamePreview.Summary.LotName -match '^LOT-AUTO-W10-FR-BE-\d{8}-\d{6}$') -Message 'default LOT name contains normalized prefix segments in operator order'

    $sanitizedNamePreview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'AUTO: W10/FR*?' -NoEvidence
    Assert-Equal -Actual $sanitizedNamePreview.Summary.LotName -Expected 'LOT-AUTO-W10-FR' -Message 'Windows-forbidden directory characters are sanitized from a manual LOT name'

    $reservedNamePreview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'CON' -NoEvidence
    Assert-Equal -Actual $reservedNamePreview.Summary.LotName -Expected 'LOT-CON' -Message 'mandatory LOT prefix neutralizes Windows reserved device names'

    $wildcardBlocked = $false
    try {
        & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-WILDCARD' -ComputerNamePrefix 'FR-*' -NoEvidence | Out-Null
    }
    catch {
        $wildcardBlocked = $_.Exception.Message -match 'wildcards and regular expressions are not supported'
    }
    Assert-True -Condition $wildcardBlocked -Message 'wildcard prefix is rejected explicitly'

    $containsWildcardBlocked = $false
    try {
        & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-CONTAINS-WILDCARD' -ComputerNameContains '*-A-*' -NoEvidence | Out-Null
    }
    catch {
        $containsWildcardBlocked = $_.Exception.Message -match 'wildcards and regular expressions are not supported'
    }
    Assert-True -Condition $containsWildcardBlocked -Message 'wildcard contains filter is rejected explicitly'

    $intuneFilterSourceBlocked = $false
    try {
        & $engine -Source AD -AdInventoryCsv $adCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-INTUNE-FILTER-SOURCE' -ExcludeIntunePresent -NoEvidence | Out-Null
    }
    catch {
        $intuneFilterSourceBlocked = $_.Exception.Message -match 'requires an available Intune inventory source'
    }
    Assert-True -Condition $intuneFilterSourceBlocked -Message 'Intune presence filter requires Intune inventory evidence'

    $adFilterSourceBlocked = $false
    try {
        & $engine -Source Intune -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-AD-FILTER-SOURCE' -ExcludeStaleAd -NoEvidence | Out-Null
    }
    catch {
        $adFilterSourceBlocked = $_.Exception.Message -match 'requires an available AD inventory source'
    }
    Assert-True -Condition $adFilterSourceBlocked -Message 'AD LastLogon filter requires AD inventory evidence'

    $missingSourceBlocked = $false
    try {
        & $engine -Source Both -AdInventoryCsv $adCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-MISSING-SOURCE' -NoEvidence | Out-Null
    }
    catch {
        $missingSourceBlocked = $_.Exception.Message -match 'source\(s\) missing'
    }
    Assert-True -Condition $missingSourceBlocked -Message 'missing requested source is blocked by default'

    $partial = & $engine -Source Both -AdInventoryCsv $adCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-PARTIAL' -AllowPartialSource -NoEvidence
    Assert-True -Condition ([bool]$partial.Summary.PartialSource) -Message 'explicit partial-source preview is marked partial'
    Assert-Equal -Actual $partial.Summary.AvailableSources -Expected 'AD' -Message 'partial-source preview reports AD as the only source'
    Assert-Equal -Actual $partial.Summary.SelectedDevices -Expected 4 -Message 'AD-only partial-source selected count'

    $created = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-TEST' -EvidenceRoot $evidenceRoot -Create -SkipWrapperRefresh
    Assert-True -Condition (Test-Path -LiteralPath $created.Summary.ComputersPath -PathType Leaf) -Message 'Computers.txt was created'
    Assert-Equal -Actual @(Get-Content -LiteralPath $created.Summary.ComputersPath).Count -Expected 3 -Message 'Computers.txt row count'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $created.Summary.EvidencePath 'AutomaticLotSelection.csv') -PathType Leaf) -Message 'selection evidence was created'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $created.Summary.EvidencePath 'AutomaticLotExclusions.csv') -PathType Leaf) -Message 'exclusion evidence was created'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $created.Summary.EvidencePath 'AutomaticLotFilterExclusions.csv') -PathType Leaf) -Message 'empty filter evidence was created for the default scope'

    $filteredCreated = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-FILTERED-CREATE' -EvidenceRoot $evidenceRoot -ComputerNamePrefix 'FR-' -Create -SkipWrapperRefresh
    $filteredComputers = @(Get-Content -LiteralPath $filteredCreated.Summary.ComputersPath)
    Assert-Equal -Actual $filteredComputers.Count -Expected 1 -Message 'filtered Computers.txt row count'
    Assert-Equal -Actual $filteredComputers[0] -Expected 'fr-pc1.one.example' -Message 'filtered Computers.txt contains only the eligible FR device'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $filteredCreated.Summary.EvidencePath 'AutomaticLotFilterExclusions.csv') -PathType Leaf) -Message 'filtered create evidence was created'

    $syntheticWrapperPath = Join-Path $testToolkitRoot 'Scripts\SmartM365-Windows11Upgrade-Update-LotCmdWrappers.ps1'
    New-Item -ItemType Directory -Path (Split-Path -Parent $syntheticWrapperPath) -Force | Out-Null
    Set-Content -LiteralPath $syntheticWrapperPath -Encoding UTF8 -Value @(
        'param([string]$ToolkitRoot)'
        'Write-Output "SYNTHETIC_WRAPPER_OUTPUT"'
        'Write-Host "SYNTHETIC_WRAPPER_HOST"'
        'exit 0'
    )
    $progressStages = New-Object System.Collections.Generic.List[string]
    $progressCallback = {
        param([string]$Stage, [string]$Detail)
        [void]$progressStages.Add($Stage)
    }.GetNewClosure()
    $wrapperCreateOutput = @(
        & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-WRAPPER-OUTPUT' -NoEvidence -Create -ProgressCallback $progressCallback
    )
    Assert-Equal -Actual $wrapperCreateOutput.Count -Expected 1 -Message 'wrapper console output does not contaminate the automatic LOT engine result'
    Assert-True -Condition ($null -ne $wrapperCreateOutput[0].PSObject.Properties['Summary']) -Message 'wrapper create returns the Summary result object'
    Assert-Equal -Actual $wrapperCreateOutput[0].Summary.LotName -Expected 'LOT-AUTO-WRAPPER-OUTPUT' -Message 'wrapper create result preserves the requested LOT name'
    Assert-True -Condition ($progressStages -contains 'Refreshing LOT command wrappers...') -Message 'wrapper refresh stage is reported through the progress callback'
    Assert-True -Condition ($progressStages -contains 'Automatic LOT created.') -Message 'automatic LOT completion is reported through the progress callback'

    Write-Output 'SmartM365 Windows 11 automatic LOT synthetic tests passed.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBg+bDSkyRFitaT
# dpt1ZNxJjJeI+YW5l3B+s7K4GnZcr6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICLgdkaGa2hFij0ZbJqXvzlKHRY0R6SlFC2Q1O+ZQJyeMA0GCSqG
# SIb3DQEBAQUABIIBgDXwjaz0FFychg0CyFXLLyxo8MMtDaQcNgUg4wwF1qWitwdd
# zpTatiry5yqgOmsb+c4aWVKvDxUbH1eycyOuyPDaZTYK9AR6i7Ihu/Nil83UuyZu
# CJPOikuB5EHnE65s79KZozbR/Fp64jzZUr4wCs44EC/3ZGyUmJshMVmazSE7T5LG
# zBYeHG8JQ41trbB0xE7qXc5gx54lWtjCaRQuRrP3azBlh9BfmwRjo9WHM2a3GYEM
# Il2/6AGseN1cTJOT87Fz3rZ4XDoM0S3+y7MudpEJLek5y8Kz+Uu2D5PXHGauyCez
# eysP4Dm2xfig8ScDKqE3413Xsf7U3/3y2/W1esU4Gl+WS3OnVQIXek3I3YINziBA
# XbKAQVmOPTE4Mcm3+7w68QiWEk8g4o4kLRZ7HDAWnoyYUkWIkJewCWQOFpMJF3Mn
# DylCkWV+Q/9ktoTJ+Mv1cSwe30K9htSy7V91Kv7QiJvxKu2lLP+PFCSdkSgJEQtj
# 1452uyISu2cC/2Jmw6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjcxMzU5
# MzVaMC8GCSqGSIb3DQEJBDEiBCCOP1CgHS0iC8+m01i8WEKvZG4WZCS6Yn++a6/+
# xR2LFTANBgkqhkiG9w0BAQEFAASCAgCHRe0h39lSj9VVpzQk0m1pwyBdFvRqEQ5k
# k1hPom3KKcRjWrpfAKoBccYl0O+NX9Gg4UO3AHjUMTGyWfXw6LBQKEgQIsqWgl5V
# LwVnSYCWgjKOtfcKW0nyZsXqVHGK+diZ0W3UCVCqog/4xSd8cnRZu52ABoF0wAZi
# KQEpyP8oOd3GSXceEBEzzML6XDva+DWfA8zVGlmQe91o3JqVIfKVuvXM3tPJ1duV
# OzyKf3jrJwmGT2CkNurvaM6FvwTwvPqi2r7LCpV72nG1O2cRUd2cXoD0hcH35ykk
# 08rFtmJ4ZZlUmeMtcSnK5ywmxHh8RmwDB9mP8qMQiCPUJRF17SV+5laoTnWj4oyU
# W9b7yVSnG8Uu9zShmjvV9KiJnx7K8QNhyYY++eKXf/CWXGUZ1b+txJOdnr2AgvUp
# tBpkAEhLRjdWDRUKaAhn3uQ0PXMGsGWQ0RWA8amPWYY7kOsw7PYjAUkEWI0h/nqU
# Fj2SMxC/DUzrW+zrM+24+J3UnEmmaKzORt+FttdWRQJezHaAuftJHsu3CWSCRYzp
# VEFrWcMcE/cCghO4HnsS86Gnwkq5uNooCwYsrrQUKk5ka0YBaKlQsQbKXrwmMOcU
# DTpkIG1qTDnsIxnlLwcrDB5e+lOsy3IINhxGC2mOklE39yBQ/ZK3KqO2U8jaNJoc
# iuWlCrwzJg==
# SIG # End signature block
