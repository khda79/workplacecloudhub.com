<#
.SYNOPSIS
Runs synthetic tests for automatic Windows 10 LOT selection.

.VERSION
1.0.0
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

try {
    New-Item -ItemType Directory -Path $testToolkitRoot -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $testToolkitRoot 'Windows11UpgradeToolkit.config.template') -Encoding ASCII -Value @(
        '# Synthetic automatic LOT test configuration.'
        'W11UT_SETUP_SOURCE='
    )

    @(
        [pscustomobject]@{ ComputerName = 'PC1'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $true; DNSHostName = 'pc1.one.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
        [pscustomobject]@{ ComputerName = 'PC2'; ADInventoryPresent = $true; ADDomain = 'one.example'; Enabled = $false; DNSHostName = 'pc2.one.example'; OperatingSystem = 'Windows 10 Enterprise'; OperatingSystemVersion = '10.0 (19045)'; LastLogonTimestampUtc = (Get-Date).AddDays(-2).ToString('o') }
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

    $preview = & $engine -Source Both -AdInventoryCsv $adCsv -IntuneInventoryCsv $intuneCsv -ToolkitRoot $testToolkitRoot -LotName 'LOT-AUTO-TEST' -EvidenceRoot $evidenceRoot
    Assert-Equal -Actual $preview.Summary.SelectedDevices -Expected 3 -Message 'selected device count'
    Assert-Equal -Actual $preview.Summary.Windows11Excluded -Expected 3 -Message 'Windows 11 exclusion count, including duplicate historical Intune evidence'
    Assert-Equal -Actual $preview.Summary.ADDisabledExcluded -Expected 2 -Message 'disabled AD exclusion count, including a conflicting Intune Windows 10 record'
    Assert-Equal -Actual $preview.Summary.IntuneStateExcluded -Expected 2 -Message 'Intune management-state exclusion count, including a conflicting AD Windows 10 record'
    Assert-Equal -Actual $preview.Summary.ADNameCollisions -Expected 1 -Message 'AD collision count'
    Assert-Equal -Actual $preview.Summary.IntuneDuplicateRowsIgnored -Expected 2 -Message 'Intune duplicate count'
    Assert-Equal -Actual $preview.Summary.UnknownOSExcluded -Expected 2 -Message 'non-Windows and Windows Server exclusion count'
    Assert-Equal -Actual $preview.Summary.ADStaleWarnings -Expected 1 -Message 'AD stale warning count'

    $selectedNames = @($preview.SelectedDevices | Select-Object -ExpandProperty ComputerName)
    Assert-True -Condition ($selectedNames -contains 'pc1.one.example') -Message 'AD FQDN is preferred for PC1'
    Assert-True -Condition ($selectedNames -contains 'PC5') -Message 'Intune-only PC5 is selected'
    Assert-True -Condition ($selectedNames -contains 'stalead.one.example') -Message 'stale AD Windows 10 remains selected'
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

    Write-Output 'SmartM365 Windows 11 automatic LOT synthetic tests passed.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
