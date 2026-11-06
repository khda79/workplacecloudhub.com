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

    # Intune enrollment proven by the Entra device ID (not by name) is excluded; anything weaker stays selected.
    $tenantA = '11111111-1111-1111-1111-111111111111'; $tenantB = '22222222-2222-2222-2222-222222222222'
    $idAdPath = Join-Path $dataRoot 'DevicesAD-IntuneId.csv'; $idIntunePath = Join-Path $dataRoot 'DevicesIntune-IntuneId.csv'; $idEntraPath = Join-Path $dataRoot 'DevicesEntra-IntuneId.csv'
    $idCases = [ordered]@{ 'ID-CONFIRMED'='a0000000-0000-0000-0000-000000000001'; 'ID-NAMEONLY'='a0000000-0000-0000-0000-000000000002'; 'ID-STALESYNC'='a0000000-0000-0000-0000-000000000003'; 'ID-PENDING'='a0000000-0000-0000-0000-000000000004'; 'ID-OTHERTENANT'='a0000000-0000-0000-0000-000000000005'; 'ID-NOTENANT'='a0000000-0000-0000-0000-000000000006'; 'ID-AMBIGUOUS'='a0000000-0000-0000-0000-000000000007' }
    @($idCases.Keys | ForEach-Object { [pscustomobject]@{ ComputerName=$_; Name=$_; ADInventoryPresent=$true; Enabled=$true; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)' } }) | Export-Csv -LiteralPath $idAdPath -NoTypeInformation -Encoding UTF8
    @($idCases.Keys | ForEach-Object { [pscustomobject]@{ ComputerName=$_; DisplayName=$_; EntraInventoryPresent=$true; EntraRegisteredState=$(if ($_ -eq 'ID-PENDING') { 'Pending' } else { 'Registered' }); TrustType='ServerAd'; AlternativeSecurityIdCount=$(if ($_ -eq 'ID-PENDING') { 0 } else { 1 }); AccountEnabled=$true; DeviceId=$idCases[$_]; EntraObjectId=('obj-' + $_); InventoryTenantId=$tenantA } }) | Export-Csv -LiteralPath $idEntraPath -NoTypeInformation -Encoding UTF8
    @($idCases.Keys | ForEach-Object {
        $aadId = if ($_ -eq 'ID-NAMEONLY') { 'b0000000-0000-0000-0000-000000000099' } else { $idCases[$_] }
        $sync = if ($_ -eq 'ID-STALESYNC') { $now.AddDays(-40) } else { $now.AddHours(-2) }
        $tenant = switch ($_) { 'ID-OTHERTENANT' { $tenantB } 'ID-NOTENANT' { '' } default { $tenantA } }
        [pscustomobject]@{ ComputerName=$_; DeviceName=$_; IntuneInventoryPresent=$true; IntuneManagedDeviceId=('managed-' + $_); AzureADDeviceId=$aadId; LastSyncDateTime=$sync.ToString('o'); InventoryTenantId=$tenant }
        # A second Intune record (another name) carries the same Entra device ID: ambiguous.
        if ($_ -eq 'ID-AMBIGUOUS') { [pscustomobject]@{ ComputerName='OTHER-NAME'; DeviceName='OTHER-NAME'; IntuneInventoryPresent=$true; IntuneManagedDeviceId='managed-other'; AzureADDeviceId=$aadId; LastSyncDateTime=$sync.ToString('o'); InventoryTenantId=$tenant } }
    }) | Export-Csv -LiteralPath $idIntunePath -NoTypeInformation -Encoding UTF8
    $idPreview = & $engine -AdInventoryCsv $idAdPath -IntuneInventoryCsv $idIntunePath -EntraInventoryCsv $idEntraPath -WaveName 'WAVE-AUTO-HJ-INTUNE-ID' -NoEvidence
    $idSelected = @($idPreview.SelectedDevices | Select-Object -ExpandProperty ComputerName)
    Assert-Equal $idPreview.Summary.IntuneConfirmedExcluded 2 'Intune enrollment proven by device ID is excluded (same tenant or no tenant column)'
    Assert-True (-not ($idSelected -contains 'ID-CONFIRMED' -or $idSelected -contains 'ID-NOTENANT')) 'confirmed Intune devices are not selected'
    foreach ($candidate in @('ID-NAMEONLY','ID-STALESYNC','ID-PENDING','ID-OTHERTENANT','ID-AMBIGUOUS')) { Assert-True ($idSelected -contains $candidate) ("$candidate remains a candidate for local verification") }
    $idChecks = @{}; foreach ($row in @($idPreview.SelectedDevices)) { $idChecks[[string]$row.ComputerName] = [string]$row.IntuneDeviceIdCheck }
    Assert-Equal $idChecks['ID-AMBIGUOUS'] 'AmbiguousIntuneDeviceId' 'a device ID shared by two Intune records is logged as ambiguous'
    Assert-Equal $idChecks['ID-NAMEONLY'] 'NameOnly' 'a name-only match is logged'
    Assert-Equal $idChecks['ID-STALESYNC'] 'StaleSync' 'an old Intune sync is logged'
    Assert-Equal $idChecks['ID-PENDING'] 'EntraNotRegistered' 'a pending Entra object is logged'
    Assert-Equal $idChecks['ID-OTHERTENANT'] 'OtherTenant' 'another tenant is logged'
    Assert-Equal $idPreview.Summary.IntuneDeviceIdAmbiguousSelected 1 'ambiguous Intune device IDs are counted'
    Assert-Equal ([int]$idPreview.Summary.ExclusionReasonCounts['INTUNE_ENROLLED_CONFIRMED_BY_DEVICE_ID']) 2 'count per exclusion reason includes the device-ID rule'
    Assert-Equal ([int]$preview.Summary.ExclusionReasonCounts['AD_DISABLED']) 1 'count per exclusion reason matches the safety exclusions'
    Assert-True ([string]$idPreview.Summary.SelectionSignature -match 'ExcludeIntuneConfirmedByDeviceId=True') 'selection signature records the device-ID rule (an older preview is not reused)'
    Assert-True ([string]$idPreview.Summary.SelectionSignature -match 'IntuneIdentityProof=AdObjectGuid,EntraDeviceId') 'selection signature records the objectGUID proof'

    # With an AD objectGUID (Entra device ID of a hybrid-joined computer), the proof is objectGUID = Intune Entra device ID.
    $guidAdPath = Join-Path $dataRoot 'DevicesAD-ObjectGuid.csv'; $guidIntunePath = Join-Path $dataRoot 'DevicesIntune-ObjectGuid.csv'; $guidEntraPath = Join-Path $dataRoot 'DevicesEntra-ObjectGuid.csv'
    $guidCases = [ordered]@{ 'GUID-CONFIRMED'='c0000000-0000-0000-0000-000000000001'; 'GUID-MISMATCH'='c0000000-0000-0000-0000-000000000002'; 'GUID-STALE'='c0000000-0000-0000-0000-000000000003'; 'GUID-NOENTRA'='c0000000-0000-0000-0000-000000000004' }
    $mismatchEntraId = 'd0000000-0000-0000-0000-000000000002'
    @($guidCases.Keys | ForEach-Object { [pscustomobject]@{ ComputerName=$_; Name=$_; ADInventoryPresent=$true; Enabled=$true; ObjectGUID=$guidCases[$_]; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)' } }) | Export-Csv -LiteralPath $guidAdPath -NoTypeInformation -Encoding UTF8
    @($guidCases.Keys | Where-Object { $_ -ne 'GUID-NOENTRA' } | ForEach-Object { [pscustomobject]@{ ComputerName=$_; DisplayName=$_; EntraInventoryPresent=$true; EntraRegisteredState='Registered'; TrustType='ServerAd'; AlternativeSecurityIdCount=1; AccountEnabled=$true; DeviceId=$(if ($_ -eq 'GUID-MISMATCH') { $mismatchEntraId } else { $guidCases[$_] }); EntraObjectId=('obj-' + $_); InventoryTenantId=$tenantA } }) | Export-Csv -LiteralPath $guidEntraPath -NoTypeInformation -Encoding UTF8
    @($guidCases.Keys | ForEach-Object {
        $aadId = if ($_ -eq 'GUID-MISMATCH') { $mismatchEntraId } else { $guidCases[$_] }
        $sync = if ($_ -eq 'GUID-STALE') { $now.AddDays(-40) } else { $now.AddHours(-2) }
        [pscustomobject]@{ ComputerName=$_; DeviceName=$_; IntuneInventoryPresent=$true; IntuneManagedDeviceId=('managed-' + $_); AzureADDeviceId=$aadId; LastSyncDateTime=$sync.ToString('o'); InventoryTenantId=$tenantA }
    }) | Export-Csv -LiteralPath $guidIntunePath -NoTypeInformation -Encoding UTF8
    $guidPreview = & $engine -AdInventoryCsv $guidAdPath -IntuneInventoryCsv $guidIntunePath -EntraInventoryCsv $guidEntraPath -WaveName 'WAVE-AUTO-HJ-OBJECTGUID' -NoEvidence
    $guidSelected = @{}; foreach ($row in @($guidPreview.SelectedDevices)) { $guidSelected[[string]$row.ComputerName] = [string]$row.IntuneDeviceIdCheck }
    $guidExcluded = @{}; foreach ($row in @($guidPreview.ExcludedDevices)) { $guidExcluded[[string]$row.ComputerName] = [string]$row.IntuneDeviceIdCheck }
    Assert-Equal $guidExcluded['GUID-CONFIRMED'] 'ConfirmedByAdObjectGuid' 'Intune device carrying the AD objectGUID is excluded'
    Assert-Equal $guidSelected['GUID-NOENTRA'] 'EntraNotRegistered' 'the objectGUID proof still needs the single registered Entra ServerAd object'
    Assert-Equal $guidSelected['GUID-MISMATCH'] 'AdObjectGuidMismatch' 'Entra and Intune registered under another identity than the AD object stay selected'
    Assert-Equal $guidSelected['GUID-STALE'] 'StaleSync' 'an old Intune sync stays selected with the objectGUID proof'
    Assert-Equal $guidPreview.Summary.IntuneConfirmedByAdObjectGuidExcluded 1 'objectGUID exclusions are counted'
    Assert-Equal $guidPreview.Summary.AdObjectGuidMismatchSelected 1 'objectGUID mismatches are counted'

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
    Assert-True (Test-Path -LiteralPath (Join-Path $evidencePreview.Summary.EvidencePath 'AutomaticWaveSummary.json.txt') -PathType Leaf) 'preview writes AutomaticWaveSummary.json.txt'
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
    $waveConfig = Get-Content -LiteralPath (Join-Path $created.Summary.WavePath 'Wave.local.json.txt') -Raw | ConvertFrom-Json
    Assert-Equal $waveConfig.Source 'Automatic' 'Wave.local.json.txt records Source=Automatic'
    Assert-Equal $waveConfig.Campaign 'HybridJoin' 'Wave.local.json.txt records the HybridJoin Campaign'
    Assert-Equal $waveConfig.EvidencePath $created.Summary.EvidencePath 'Wave.local.json.txt records the evidence path'
    Assert-True ([string]$waveConfig.Description -match 'ComputerNamePrefix=' -and [string]$waveConfig.Description -match 'Selected=5') 'Wave description summarizes the filters'
    Assert-True ($createStages -contains 'Automatic Wave created.') 'creation completion is reported'
    Assert-True (([string]$created.Summary.EvidencePath).StartsWith($expectedEvidenceRoot, [System.StringComparison]::OrdinalIgnoreCase)) 'create evidence is written under Runs\HybridJoin\_AutomaticWave'
    foreach ($name in @('DevicesAD.csv', 'DevicesIntune.csv', 'DevicesEntra.csv', 'AutomaticWaveSelection.csv', 'AutomaticWaveExclusions.csv', 'AutomaticWaveFilterExclusions.csv', 'AutomaticWaveSummary.json.txt')) {
        Assert-True (Test-Path -LiteralPath (Join-Path $created.Summary.EvidencePath $name) -PathType Leaf) "evidence file exists: $name"
    }
    $summaryJson = Get-Content -LiteralPath (Join-Path $created.Summary.EvidencePath 'AutomaticWaveSummary.json.txt') -Raw | ConvertFrom-Json
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
    Assert-True ([string](Get-Content -LiteralPath (Join-Path $reused.Summary.WavePath 'Wave.local.json.txt') -Raw | ConvertFrom-Json).Description -match ('Selected={0}$' -f $reusePreview.Summary.SelectedDevices)) 'the reused Create keeps the Wave description format'
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBMGtqxLzy2H3YE
# 0D30dPCF9RsmyzDfOMYV6UZuJipw6KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMvw5cVRngqkP498VAV6kjUWBvg7lv+PQ4uZ0asco1DNMA0GCSqG
# SIb3DQEBAQUABIIBgKwVAd63/N5wacB2V5h5EX/eHvcjhCk/Cy9bYqVabaacc2hZ
# 1lA2DdL3NIAusfpQMXwnvh5zpVAH/LTUhUE60LGwnXuZDUHiKmfwDLAtsuUvY0Mf
# uLhJ4fSC5/vCTjqCHxM4Ny4mszKdAYvnsNJiyuxGHICWbFyZBpFcYFidJH0f8tNw
# S/+lWyJp40VQv+ZxYrKP8IZEBR8b1rUOCSMhSDfxniAZzWxGeDdHgrlJrBADoe9v
# cHej3wUEi9uSA17nBX/wCFwKhkPCIGd1Y3P93RjUtn20h6XshhlLXXNDUhAoHumq
# hz/NdPTMLPhXscHeyZp6SgnjfnaKpr0A/wlHSXBUFqmmvcsLaBFzaq5lpU/apG+L
# Si3X11jgJD1EHfNsswcnQiSjzhWaSjbvC2xNcpPDjJ8GenI7180YL1drg+1vCyqh
# 5SpiAX5uk1yJV+Z/bDJdYex0ekEvXdX8zK1MnZanRfEgRICKkF2IEtn43h7T1E4X
# OqGdM/lzZLbi32jFGaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDYxMjMx
# MjhaMC8GCSqGSIb3DQEJBDEiBCAAk4oF40G+71P8Qd5s67+UF/bW1eyvG5nJpeIP
# p2I0ADANBgkqhkiG9w0BAQEFAASCAgChSNMKZ7ONfkr6ZxDqOnZrjr12XnvZ39BC
# vW8YWcxfV8TAXxh9HRwmXyeMkMoDZeRJx5KXCqRUL3N/xSFTsGQhVPKvkDp+O7fZ
# E8XE5u5++FW0cYnObYsLhx8xpMGNWwDrwSSPa4UBbRI08W1/F2FYFvFTvbpnjKMh
# HpjdoKNWygfOGSDm28M4Lcm5ZA24s8noU2YDWss7GPqGfcRIQy+STfPU9cv47cQz
# VmmvN6UU/Yn183U7hZKrfHS5G3XMwwHf/2ffa7Hd+KCb4AKi8sB0ggoA69xP2Kct
# WSPqNhqDWKDMv6InTBKs09/+JwkV/tMDj2+y8I+E9TxieVRFTLMMXUH5wKZ+jqzC
# gBA81c+vi+ZmUnVHIUl8f9d0GnD596SP7+NRAnfhowP2g84bw+jlYb/ZOqpeMLh0
# tx582MefF3PZq7KbSdyhFj/m7vFqvhIgBd9fiVBKR2GmxQaKv/AjXzv2lVDVqEdE
# zCeysLee14nrwXTyrfjGInYGwpTokriuM1mo3WxgXGroLYohtCnJjJ7o20q4FMZH
# 4sDfuaOYpBsP08hisAUvBsgtWm7GjSmJriCuURRLaF89Pm4TEjI2iXI1NFKq9Pu4
# AOoRRMn9Pb54z7OAPS8uS0KJq5ppb+a+y+cmty+VXsT0+Pu9x4K3IViqZPWpltWV
# P68W5JfRlQ==
# SIG # End signature block
