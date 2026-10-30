<#
.SYNOPSIS
Offline synthetic tests of the HybridJoin automatic Wave engine.

.DESCRIPTION
Copies the package (without Runs, Wave contents, Inventory and DemoAssets) into a fresh
temporary folder, isolates the technician runtime with SEMT_RUNTIME_ROOT_OVERRIDE and runs
the HybridJoin automatic Wave engine against synthetic AD, Intune and Entra CSV files.
The repository folder is never written. No network, AD or Microsoft Graph access is used.
Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M"; exit code 1 on
any failure. Compatible with Windows PowerShell 5.1 and PowerShell 7.
Provenance: assertions ported from SmartM365/Devices/IntuneHybridJoinToolkit/Tests/Test-SmartM365IntuneHybridJoinAutomaticLot.ps1 (engine rules only).

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
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SEMT-HJ-AutomaticWave-{0}' -f [guid]::NewGuid().ToString('N'))
$packageRoot = Join-Path $testRoot 'Package'
$dataRoot = Join-Path $testRoot 'Data'
$previousRuntimeOverride = [Environment]::GetEnvironmentVariable('SEMT_RUNTIME_ROOT_OVERRIDE', 'Process')
$repositorySnapshotBefore = Get-RepositoryRuntimeSnapshot -Root $repositoryPackageRoot
$powerShellExe = (Get-Process -Id $PID).Path

try {
    [Environment]::SetEnvironmentVariable('SEMT_RUNTIME_ROOT_OVERRIDE', (Join-Path $testRoot 'Runtime'), 'Process')
    Copy-SemtPackageForTest -Source $repositoryPackageRoot -Destination $packageRoot
    New-Item -ItemType Directory -Path $dataRoot -Force | Out-Null
    $engine = Join-Path $packageRoot 'Campaigns\HybridJoin\Scripts\SmartM365-EndpointMigration-HybridJoin-New-AutomaticWave.ps1'
    $wrapperScript = Join-Path $packageRoot 'Scripts\SmartM365-EndpointMigration-Update-WaveCmdWrappers.ps1'
    $waveRoot = Join-Path $packageRoot 'Waves\HybridJoin'
    Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force

    $now = [datetime]::UtcNow
    $adRows = @(
        [pscustomobject]@{ ComputerName='FR-PC01'; Name='FR-PC01'; ADInventoryPresent=$true; ADDomain='fr.test'; Enabled=$true; DNSHostName='fr-pc01.fr.test'; DistinguishedName='CN=FR-PC01,DC=fr,DC=test'; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)'; LastLogonTimestampUtc=$now.ToString('o') }
        [pscustomobject]@{ ComputerName='FR-PC02'; Name='FR-PC02'; ADInventoryPresent=$true; ADDomain='fr.test'; Enabled=$true; DNSHostName='fr-pc02.fr.test'; DistinguishedName='CN=FR-PC02,DC=fr,DC=test'; OperatingSystem='Windows 11 Enterprise'; OperatingSystemVersion='10.0 (22631)'; LastLogonTimestampUtc=$now.ToString('o') }
        [pscustomobject]@{ ComputerName='FR-PC03'; Name='FR-PC03'; ADInventoryPresent=$true; ADDomain='fr.test'; Enabled=$true; DNSHostName='fr-pc03.fr.test'; DistinguishedName='CN=FR-PC03,DC=fr,DC=test'; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)'; LastLogonTimestampUtc=$now.ToString('o') }
        [pscustomobject]@{ ComputerName='FR-PC04'; Name='FR-PC04'; ADInventoryPresent=$true; ADDomain='fr.test'; Enabled=$true; DNSHostName='fr-pc04.fr.test'; DistinguishedName='CN=FR-PC04,DC=fr,DC=test'; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)'; LastLogonTimestampUtc=$now.ToString('o') }
        [pscustomobject]@{ ComputerName='FR-SRV01'; Name='FR-SRV01'; ADInventoryPresent=$true; ADDomain='fr.test'; Enabled=$true; DNSHostName='fr-srv01.fr.test'; DistinguishedName='CN=FR-SRV01,DC=fr,DC=test'; OperatingSystem='Windows Server 2022'; OperatingSystemVersion='10.0 (20348)'; LastLogonTimestampUtc=$now.ToString('o') }
        [pscustomobject]@{ ComputerName='FR-DIS01'; Name='FR-DIS01'; ADInventoryPresent=$true; ADDomain='fr.test'; Enabled=$false; DNSHostName='fr-dis01.fr.test'; DistinguishedName='CN=FR-DIS01,DC=fr,DC=test'; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)'; LastLogonTimestampUtc=$now.ToString('o') }
        [pscustomobject]@{ ComputerName='BE-OLD-A-01'; Name='BE-OLD-A-01'; ADInventoryPresent=$true; ADDomain='be.test'; Enabled=$true; DNSHostName='be-old-a-01.be.test'; DistinguishedName='CN=BE-OLD-A-01,DC=be,DC=test'; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)'; LastLogonTimestampUtc=$now.AddDays(-90).ToString('o') }
        [pscustomobject]@{ ComputerName='FR-COLLIDE'; Name='FR-COLLIDE'; ADInventoryPresent=$true; ADDomain='fr.test'; Enabled=$true; DNSHostName='fr-collide.fr.test'; DistinguishedName='CN=FR-COLLIDE,DC=fr,DC=test'; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)'; LastLogonTimestampUtc=$now.ToString('o') }
        [pscustomobject]@{ ComputerName='FR-COLLIDE'; Name='FR-COLLIDE'; ADInventoryPresent=$true; ADDomain='de.test'; Enabled=$true; DNSHostName='fr-collide.de.test'; DistinguishedName='CN=FR-COLLIDE,DC=de,DC=test'; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)'; LastLogonTimestampUtc=$now.ToString('o') }
        [pscustomobject]@{ ComputerName='FR-ENTRADUP'; Name='FR-ENTRADUP'; ADInventoryPresent=$true; ADDomain='fr.test'; Enabled=$true; DNSHostName='fr-entradup.fr.test'; DistinguishedName='CN=FR-ENTRADUP,DC=fr,DC=test'; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)'; LastLogonTimestampUtc=$now.ToString('o') }
    )
    $intuneRows = @(
        [pscustomobject]@{ ComputerName='FR-PC04'; DeviceName='FR-PC04'; IntuneInventoryPresent=$true; IntuneManagedDeviceId='intune-4'; LastSyncDateTime=$now.ToString('o') }
    )
    $entraRows = @(
        [pscustomobject]@{ ComputerName='FR-PC02'; DisplayName='FR-PC02'; EntraInventoryPresent=$true; EntraRegisteredState='Pending'; TrustType='ServerAd'; AlternativeSecurityIdCount=0; AccountEnabled=$true; DeviceId='device-2'; EntraObjectId='object-2' }
        [pscustomobject]@{ ComputerName='FR-PC03'; DisplayName='FR-PC03'; EntraInventoryPresent=$true; EntraRegisteredState='Registered'; TrustType='ServerAd'; AlternativeSecurityIdCount=1; AccountEnabled=$true; DeviceId='device-3'; EntraObjectId='object-3' }
        [pscustomobject]@{ ComputerName='FR-ENTRADUP'; DisplayName='FR-ENTRADUP'; EntraInventoryPresent=$true; EntraRegisteredState='Pending'; TrustType='ServerAd'; AlternativeSecurityIdCount=0; AccountEnabled=$true; DeviceId='dup-1'; EntraObjectId='dup-object-1' }
        [pscustomobject]@{ ComputerName='FR-ENTRADUP'; DisplayName='FR-ENTRADUP'; EntraInventoryPresent=$true; EntraRegisteredState='Registered'; TrustType='ServerAd'; AlternativeSecurityIdCount=1; AccountEnabled=$true; DeviceId='dup-2'; EntraObjectId='dup-object-2' }
    )
    $adPath = Join-Path $dataRoot 'DevicesAD.csv'
    $intunePath = Join-Path $dataRoot 'DevicesIntune.csv'
    $entraPath = Join-Path $dataRoot 'DevicesEntra.csv'
    $adRows | Export-Csv -LiteralPath $adPath -NoTypeInformation -Encoding UTF8
    $intuneRows | Export-Csv -LiteralPath $intunePath -NoTypeInformation -Encoding UTF8
    $entraRows | Export-Csv -LiteralPath $entraPath -NoTypeInformation -Encoding UTF8

    # Default preview: selection and safety rules.
    $progressStages = New-Object System.Collections.Generic.List[string]
    $progressCallback = { param($stage, $detail) [void]$progressStages.Add([string]$stage) }.GetNewClosure()
    $preview = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-TEST' -NoEvidence -ProgressCallback $progressCallback
    Assert-Equal $preview.Summary.SelectedDevices 5 'name-only Intune match remains a candidate'
    Assert-Equal $preview.Summary.NeedsHybridJoin 3 'missing ServerAd evidence is classified as needing Hybrid Join'
    Assert-Equal $preview.Summary.HybridJoinPending 1 'pending ServerAd evidence is classified'
    Assert-Equal $preview.Summary.NeedsIntuneEnrollment 1 'registered ServerAd without Intune is classified'
    Assert-Equal $preview.Summary.IntunePresentExcluded 0 'Intune name matches are not excluded'
    Assert-Equal $preview.Summary.ADDisabledExcluded 1 'disabled AD object is excluded'
    Assert-Equal $preview.Summary.WindowsServerExcluded 1 'Windows Server is excluded'
    Assert-Equal $preview.Summary.ADNameCollisions 1 'AD short-name collision is excluded'
    Assert-Equal $preview.Summary.EntraAmbiguousExcluded 1 'multiple ServerAd objects are excluded'
    Assert-True ($progressStages -contains 'Automatic Wave preview ready.') 'preview completion is reported'
    Assert-Equal $preview.Summary.WaveName 'WAVE-AUTO-HJ-TEST' 'explicit Wave name is kept'
    Assert-True (-not [bool]$preview.Summary.WaveCreated) 'preview does not create a Wave'
    Assert-Equal @(Get-ChildItem -LiteralPath $waveRoot -Directory).Count 0 'preview leaves Waves\HybridJoin empty'
    Assert-Equal @($preview.Summary.PSObject.Properties.Name | Where-Object { $_ -match '(?i)lot|batch' }).Count 0 'summary uses Wave terminology only'

    # Strict client OS rule.
    $strictAdPath = Join-Path $dataRoot 'DevicesAD-StrictClientOS.csv'
    @(
        [pscustomobject]@{ ComputerName='STRICT-W10'; Name='STRICT-W10'; ADInventoryPresent=$true; Enabled=$true; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)' }
        [pscustomobject]@{ ComputerName='STRICT-W11'; Name='STRICT-W11'; ADInventoryPresent=$true; Enabled=$true; OperatingSystem='Windows 11 Enterprise'; OperatingSystemVersion='10.0 (22631)' }
        [pscustomobject]@{ ComputerName='STRICT-GENERIC-SERVER'; Name='STRICT-GENERIC-SERVER'; ADInventoryPresent=$true; Enabled=$true; OperatingSystem='Windows'; OperatingSystemVersion='10.0 (20348)' }
        [pscustomobject]@{ ComputerName='STRICT-BLANK-SERVER'; Name='STRICT-BLANK-SERVER'; ADInventoryPresent=$true; Enabled=$true; OperatingSystem=''; OperatingSystemVersion='10.0 (26100)' }
    ) | Export-Csv -LiteralPath $strictAdPath -NoTypeInformation -Encoding UTF8
    $strictPreview = & $engine -AdInventoryCsv $strictAdPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-STRICT-OS' -NoEvidence
    Assert-Equal $strictPreview.Summary.SelectedDevices 2 'only explicitly identified Windows 10 and Windows 11 clients are selected'
    Assert-Equal $strictPreview.Summary.UnknownOSExcluded 2 'generic or blank AD operating systems remain excluded'
    $strictSelectedNames = @($strictPreview.SelectedDevices | Select-Object -ExpandProperty ComputerName)
    Assert-True ($strictSelectedNames -contains 'STRICT-W10') 'explicit Windows 10 client remains eligible'
    Assert-True ($strictSelectedNames -contains 'STRICT-W11') 'explicit Windows 11 client remains eligible'
    Assert-True (-not ($strictSelectedNames -contains 'STRICT-GENERIC-SERVER' -or $strictSelectedNames -contains 'STRICT-BLANK-SERVER')) 'version-only server-like rows are never selected'

    # Name and stale filters.
    $filtered = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-FILTERED' -ComputerNamePrefix 'FR-' -ComputerNameContains '-PC' -NoEvidence
    Assert-Equal $filtered.Summary.SelectedDevices 4 'prefix and contains filters are combined with AND'
    Assert-Equal $filtered.Summary.FilterExcludedDevices 5 'nonmatching AD keys are recorded once as filter exclusions'
    $stale = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-STALE' -ExcludeStaleAd -AdLastLogonMaxAgeDays 45 -NoEvidence
    Assert-Equal $stale.Summary.ADLastLogonFilterExcluded 1 'stale AD device is excluded by the optional 45-day filter'

    # AD domain filter: the inventory covers the forest, the Wave keeps one domain.
    $frOnly = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-FR' -AdDomainFilter 'FR.test.' -NoEvidence
    Assert-Equal $frOnly.Summary.SelectedDevices 4 'AD domain filter keeps the safe clients of that domain only'
    Assert-Equal $frOnly.Summary.ADDomainFilter 'fr.test' 'AD domain filter is normalized (case, trailing dot) and reported'
    Assert-Equal @($frOnly.FilterExcludedDevices | Where-Object { $_.FilterReason -eq 'AD_DOMAIN_NOT_MATCHED' -and $_.FilterType -eq 'ADDomain' } | ForEach-Object { $_.ComputerKey }) -join ',' 'BE-OLD-A-01' 'computer of another domain is a visible filter exclusion'
    Assert-Equal $frOnly.Summary.ADDomainFilterExcluded 1 'summary counts AD domain filter exclusions'
    Assert-True (@($frOnly.ExcludedDevices | Where-Object { $_.ComputerKey -eq 'FR-COLLIDE' -and $_.ExclusionReason -match 'AD_NAME_COLLISION' }).Count -eq 1) 'a name present in the selected and another domain stays a collision exclusion'
    $beOnly = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-BE' -AdDomainFilter 'be.test' -NoEvidence
    Assert-Equal (@($beOnly.SelectedDevices | ForEach-Object { $_.ComputerKey }) -join ',') 'BE-OLD-A-01' 'AD domain filter selects the other domain on request'
    $noDomainFilter = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-ALL' -AdDomainFilter '' -NoEvidence
    Assert-Equal $noDomainFilter.Summary.SelectedDevices 5 'empty AD domain filter keeps every domain'
    foreach ($case in @(@('xx.test','SEMT_AD_DOMAIN_NOT_IN_INVENTORY'),@('fr test','Invalid AD domain'))) {
        $failure = ''
        try { [void](& $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-BAD' -AdDomainFilter $case[0] -NoEvidence) } catch { $failure = $_.Exception.Message }
        Assert-True ($failure -match [regex]::Escape($case[1])) ("AD domain filter '{0}' is refused ({1})" -f $case[0],$case[1])
    }

    # Entra is optional (the shared Entra cache of the test package does not exist).
    $withoutEntra = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -WaveName 'WAVE-AUTO-HJ-NO-ENTRA' -NoEvidence
    Assert-Equal $withoutEntra.Summary.EntraInventoryAvailable $false 'Entra is optional'
    Assert-Equal $withoutEntra.Summary.EntraUnavailableSelected 6 'safe selected devices are marked when Entra is unavailable'

    # Generated and sanitized names.
    $generated = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -ComputerNamePrefix @('fr-', 'be-') -NoEvidence
    Assert-True ([string]$generated.Summary.WaveName -match '^WAVE-AUTO-HJ-FR-BE-\d{8}-\d{6}$') ('default Wave name is WAVE-AUTO-HJ-<prefix>-<timestamp> (actual={0})' -f $generated.Summary.WaveName)
    $longName = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName ('W' * 120) -NoEvidence
    Assert-True (([string]$longName.Summary.WaveName).Length -le 80 -and (Test-SemtWaveName -Name $longName.Summary.WaveName)) 'long Wave names are truncated to a valid 80-character name'

    # Preview with evidence writes only evidence under Runs\HybridJoin\_AutomaticWave.
    $evidencePreview = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-EVIDENCE'
    $expectedEvidenceRoot = Join-Path $packageRoot 'Runs\HybridJoin\_AutomaticWave'
    Assert-True (([string]$evidencePreview.Summary.EvidencePath).StartsWith($expectedEvidenceRoot, [System.StringComparison]::OrdinalIgnoreCase)) 'preview evidence is written under Runs\HybridJoin\_AutomaticWave'
    Assert-True (Test-Path -LiteralPath (Join-Path $evidencePreview.Summary.EvidencePath 'AutomaticWaveSummary.json') -PathType Leaf) 'preview writes AutomaticWaveSummary.json'
    Assert-Equal @(Get-ChildItem -LiteralPath $waveRoot -Directory).Count 0 'preview with evidence still creates no Wave'

    # Create through New-SemtWave.
    $createStages = New-Object System.Collections.Generic.List[string]
    $createCallback = { param($stage, $detail) [void]$createStages.Add([string]$stage) }.GetNewClosure()
    $createdOutput = @(& $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'pilot:fr/ihj' -Create -ProgressCallback $createCallback)
    Assert-Equal $createdOutput.Count 1 'create returns a single result object'
    $created = $createdOutput[0]
    Assert-Equal $created.Summary.WaveName 'WAVE-pilot-fr-ihj' 'invalid folder characters are sanitized and the WAVE- prefix is added'
    Assert-Equal $created.Summary.WavePath (Join-Path $waveRoot 'WAVE-pilot-fr-ihj') 'Wave is created under Waves\HybridJoin'
    Assert-True (Test-Path -LiteralPath $created.Summary.ComputersPath -PathType Leaf) 'Computers.txt is created'
    Assert-Equal @(Get-Content -LiteralPath $created.Summary.ComputersPath).Count 5 'created Computers.txt contains selected devices only'
    foreach ($wrapper in @('Start-Wave-Loop.cmd', 'Start-Wave-Once.cmd', 'Start-Wave-Loop-IgnoreRunGuard.cmd', 'Start-Wave-Once-IgnoreRunGuard.cmd')) {
        Assert-True (Test-Path -LiteralPath (Join-Path $created.Summary.WavePath $wrapper) -PathType Leaf) "Wave wrapper exists: $wrapper"
    }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $created.Summary.WavePath 'AdDomain.txt'))) 'no legacy AdDomain.txt is created'
    $waveConfig = Get-Content -LiteralPath (Join-Path $created.Summary.WavePath 'Wave.local.json') -Raw | ConvertFrom-Json
    Assert-Equal $waveConfig.Source 'Automatic' 'Wave.local.json records Source=Automatic'
    Assert-Equal $waveConfig.Campaign 'HybridJoin' 'Wave.local.json records the HybridJoin Campaign'
    Assert-Equal $waveConfig.EvidencePath $created.Summary.EvidencePath 'Wave.local.json records the evidence path'
    Assert-True ([string]$waveConfig.Description -match 'ComputerNamePrefix=' -and [string]$waveConfig.Description -match 'Selected=5') 'Wave description summarizes the filters'
    Assert-True ($createStages -contains 'Automatic Wave created.') 'creation completion is reported'
    Assert-True (([string]$created.Summary.EvidencePath).StartsWith($expectedEvidenceRoot, [System.StringComparison]::OrdinalIgnoreCase)) 'create evidence is written under Runs\HybridJoin\_AutomaticWave'
    foreach ($name in @('DevicesAD.csv', 'DevicesIntune.csv', 'DevicesEntra.csv', 'AutomaticWaveSelection.csv', 'AutomaticWaveExclusions.csv', 'AutomaticWaveFilterExclusions.csv', 'AutomaticWaveSummary.json')) {
        Assert-True (Test-Path -LiteralPath (Join-Path $created.Summary.EvidencePath $name) -PathType Leaf) "evidence file exists: $name"
    }
    $summaryJson = Get-Content -LiteralPath (Join-Path $created.Summary.EvidencePath 'AutomaticWaveSummary.json') -Raw | ConvertFrom-Json
    Assert-Equal $summaryJson.WavePath $created.Summary.WavePath 'evidence summary records the created Wave path'
    Assert-Equal @(Get-SemtWaves -PackageRoot $packageRoot -Campaign HybridJoin | Where-Object { $_.Launchable }).Count 1 'created Wave is listed as launchable by the engine'

    # Create guards.
    Assert-Throws { & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'pilot:fr/ihj' -Create } 'SEMT_WAVE_EXISTS' 'creating an existing Wave is refused'
    Assert-Throws { & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-EMPTY' -ComputerNamePrefix 'ZZZ' -Create } 'No eligible Hybrid Join device' 'creating a Wave without selected device is refused'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $waveRoot 'WAVE-EMPTY'))) 'refused empty Wave leaves no folder'
    Assert-Throws { & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -WaveName 'WAVE-INVALID' -ComputerNamePrefix 'FR-*' -NoEvidence } 'Invalid computer name prefix' 'wildcard prefix is rejected'
    Assert-Throws { & $engine -AdInventoryCsv (Join-Path $dataRoot 'Missing.csv') -IntuneInventoryCsv $intunePath -NoEvidence } 'Required automatic Wave inventory not found' 'missing AD inventory is refused'

    $reuseInventory = $adPath
    # Create reuses a matching Preview instead of analyzing the inventories again; other filters or a
    # changed inventory are analyzed again. The Waves of this block are removed afterwards.
    $reusePreview = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-REUSE'
    Assert-True (-not [string]::IsNullOrWhiteSpace([string]$reusePreview.Summary.SelectionSignature) -and -not [string]::IsNullOrWhiteSpace([string]$reusePreview.Summary.InventoryFingerprint)) 'preview records the selection signature and the inventory fingerprint'
    $reuseStages = [System.Collections.Generic.List[string]]::new()
    $reuseCallback = { param($stage, $detail) [void]$reuseStages.Add(('{0} {1}' -f $stage, $detail)) }.GetNewClosure()
    $reused = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-REUSE' -Create -PreviewEvidencePath $reusePreview.Summary.EvidencePath -ProgressCallback $reuseCallback
    Assert-True (@($reuseStages | Where-Object { $_ -like 'Reusing the previewed selection*' }).Count -eq 1 -and @($reuseStages | Where-Object { $_ -like 'Loading inventory records*' }).Count -eq 0) 'Create reuses a matching preview without loading the inventories'
    Assert-Equal $reused.Summary.PreviewEvidencePath $reusePreview.Summary.EvidencePath 'the Create summary records the reused preview'
    Assert-Equal @(Get-Content -LiteralPath $reused.Summary.ComputersPath).Count ([int]$reusePreview.Summary.SelectedDevices) 'the Wave contains the previewed selection'
    Assert-True (Test-Path -LiteralPath (Join-Path $reused.Summary.EvidencePath 'AutomaticWaveSelection.csv') -PathType Leaf) 'the preview reports are copied to the Create evidence'
    Assert-True ([string](Get-Content -LiteralPath (Join-Path $reused.Summary.WavePath 'Wave.local.json') -Raw | ConvertFrom-Json).Description -match ('Selected={0}$' -f $reusePreview.Summary.SelectedDevices)) 'the reused Create keeps the Wave description format'
    $reuseStages.Clear()
    try { [void](& $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-REUSE-FILTER' -AdLastLogonMaxAgeDays 400 -Create -PreviewEvidencePath $reusePreview.Summary.EvidencePath -ProgressCallback $reuseCallback) } catch { }
    Assert-True (@($reuseStages | Where-Object { $_ -like 'Computing the selection again*filters changed*' }).Count -eq 1 -and @($reuseStages | Where-Object { $_ -like 'Loading inventory records*' }).Count -eq 1) 'Create with other filters computes the selection again'
    $reuseStages.Clear()
    $inventoryItem = Get-Item -LiteralPath $reuseInventory
    $inventoryItem.LastWriteTimeUtc = $inventoryItem.LastWriteTimeUtc.AddMinutes(1)
    try { [void](& $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -WaveName 'WAVE-AUTO-HJ-REUSE-INVENTORY' -Create -PreviewEvidencePath $reusePreview.Summary.EvidencePath -ProgressCallback $reuseCallback) } catch { }
    Assert-True (@($reuseStages | Where-Object { $_ -like 'Computing the selection again*inventory changed*' }).Count -eq 1) 'Create after an inventory change computes the selection again'
    foreach ($reuseWave in @('WAVE-AUTO-HJ-REUSE', 'WAVE-AUTO-HJ-REUSE-FILTER', 'WAVE-AUTO-HJ-REUSE-INVENTORY')) {
        $reuseWavePath = Join-Path $waveRoot $reuseWave
        if (Test-Path -LiteralPath $reuseWavePath) { Remove-Item -LiteralPath $reuseWavePath -Recurse -Force }
    }

    # Update-WaveCmdWrappers restores wrappers and removes obsolete ones.
    Remove-Item -LiteralPath (Join-Path $created.Summary.WavePath 'Start-Wave-Once.cmd') -Force
    Set-Content -LiteralPath (Join-Path $created.Summary.WavePath 'Start-Wave-Obsolete.cmd') -Value '@echo off' -Encoding ASCII
    $wrapperOutput = @(& $powerShellExe -NoProfile -ExecutionPolicy Bypass -File $wrapperScript -Campaign HybridJoin 2>&1)
    $wrapperExit = $LASTEXITCODE
    Assert-Equal $wrapperExit 0 'Update-WaveCmdWrappers exits with code 0'
    Assert-True (Test-Path -LiteralPath (Join-Path $created.Summary.WavePath 'Start-Wave-Once.cmd') -PathType Leaf) 'missing wrapper is regenerated'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $created.Summary.WavePath 'Start-Wave-Obsolete.cmd'))) 'obsolete wrapper is removed'
    Assert-True (@($wrapperOutput | Where-Object { [string]$_ -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} Summary: Waves=1; Updated=1; Failed=0' }).Count -eq 1) 'Update-WaveCmdWrappers prints a timestamped summary'
    $singleOutput = @(& $powerShellExe -NoProfile -ExecutionPolicy Bypass -File $wrapperScript -WavePath $created.Summary.WavePath 2>&1)
    Assert-Equal $LASTEXITCODE 0 'Update-WaveCmdWrappers -WavePath exits with code 0'
    $badOutput = @(& $powerShellExe -NoProfile -ExecutionPolicy Bypass -File $wrapperScript -WavePath (Join-Path $testRoot 'NotAWave') 2>&1)
    Assert-Equal $LASTEXITCODE 1 'Update-WaveCmdWrappers fails with code 1 for an invalid Wave path'

    # Shared inventory provenance guard used for cache reuse.
    $cacheTenant = '00000000-0000-0000-0000-000000000001'
    $cacheIntunePath = Join-Path $dataRoot 'CacheIntune.csv'
    @([pscustomobject]@{ ComputerName='CACHE-PC'; DeviceName='CACHE-PC'; IntuneInventoryPresent=$true; InventoryTenantId=$cacheTenant; InventoryAuthenticationMode='DelegatedInteractive'; InventoryScope='AllManagedDevices' }) |
        Export-Csv -LiteralPath $cacheIntunePath -NoTypeInformation -Encoding UTF8
    Assert-True ((Test-SemtSharedInventoryUsable -Path $cacheIntunePath -Kind Intune -MaxAgeHours 2 -TenantId $cacheTenant).Usable) 'fresh delegated full Intune cache is reusable'
    Assert-True (-not (Test-SemtSharedInventoryUsable -Path $cacheIntunePath -Kind Intune -MaxAgeHours 2 -TenantId 'different-tenant').Usable) 'tenant-mismatched Intune cache is rejected'
    $badScopePath = Join-Path $dataRoot 'CacheIntuneScoped.csv'
    @([pscustomobject]@{ ComputerName='CACHE-PC'; DeviceName='CACHE-PC'; IntuneInventoryPresent=$true; InventoryTenantId=$cacheTenant; InventoryAuthenticationMode='DelegatedInteractive'; InventoryScope='ComputerList' }) |
        Export-Csv -LiteralPath $badScopePath -NoTypeInformation -Encoding UTF8
    Assert-True (-not (Test-SemtSharedInventoryUsable -Path $badScopePath -Kind Intune -MaxAgeHours 2 -TenantId $cacheTenant).Usable) 'targeted Intune cache is rejected for full-inventory reuse'

    # Default inputs are the shared caches of the package.
    $inventoryRoot = Join-Path $packageRoot 'Inventory'
    New-Item -ItemType Directory -Path $inventoryRoot -Force | Out-Null
    Copy-Item -LiteralPath $adPath -Destination (Get-SemtSharedInventoryPath -PackageRoot $packageRoot -Kind AD) -Force
    Copy-Item -LiteralPath $intunePath -Destination (Get-SemtSharedInventoryPath -PackageRoot $packageRoot -Kind Intune) -Force
    Copy-Item -LiteralPath $entraPath -Destination (Get-SemtSharedInventoryPath -PackageRoot $packageRoot -Kind Entra) -Force
    $defaults = & $engine -NoEvidence
    Assert-Equal $defaults.Summary.SelectedDevices 5 'default inventory inputs are the shared caches'
    Assert-Equal $defaults.Summary.EntraInventoryAvailable $true 'shared Entra cache is used by default when present'
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZM48mfecnp1on
# bOa8FJV1+u3+n/VOcwjqOQMOr4orNKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDspeidMXhR48W0IyEDt7Tw
# ruBkRfBMrqaR4EiUJyMZwzANBgkqhkiG9w0BAQEFAASCAYBnnkl43ic+S0FQNBhs
# x0+cilcCInD0/0Xz7Pv22qbwHTLZkdIg3Y2KhfbS7Dahg0L19Mq+U//DBQN2qkKA
# qqfNjhLSxHbXnZ+zo9MFWDRSDfbbbO814kmCs+Y6j4XYrU771tDw1m20CwPVFP2X
# Dt3+27GbPqWxeMVYSTkE64VKRD39qhPBHkf4rOOk39O434MucyEAZuqCchlL6MEd
# mY9OaF1qpcy5YkINZJaxxBYkXxL0OgtWJqBm3uTL2ksc+AfRhHBWfda+0VPwM4ZA
# 4700lDLh+HPFuVLEhCmBUDUCXOgzf1jn7+HH2cCkHgPQr6CI9UpSQxkqrAHd/WnL
# Ga8tJQUw+YEmuGlD0FKjxiFwwNOyCtCROPLs/puIuwL3LY626slAEETGphNJ6MKp
# 3Qx8RlyJIuwTI9HtIyGWiWCaZeFd2U568Sru+NIBsi+8+FMbOjy5gVIKBi+OCkvS
# 3KtdwMDl29BM0MxKjYzN+2XlboryrdrXgyS5fQm0gCD31TI=
# SIG # End signature block
