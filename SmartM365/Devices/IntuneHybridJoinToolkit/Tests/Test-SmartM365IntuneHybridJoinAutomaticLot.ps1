<#
.SYNOPSIS
Runs synthetic tests for the Intune Hybrid Join Automatic LOT workflow.
#>

#requires -Version 5.1

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Assert-Equal {
    param(
        $Actual,
        $Expected,
        [string]$Message
    )
    if ([string]$Actual -ne [string]$Expected) {
        throw "ASSERT FAILED: $Message. Expected='$Expected'; Actual='$Actual'."
    }
}

$toolkitRoot = Split-Path -Parent $PSScriptRoot
$engine = Join-Path $toolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-New-AutomaticLot.ps1'
$gui = Join-Path $toolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1'
$guiSupport = Join-Path $toolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-AutomaticLotGuiSupport.ps1'
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SmartM365-IHJ-AutomaticLot-{0}' -f [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $testRoot 'Runs') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $testRoot 'Lots') -Force | Out-Null

try {
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
    $adPath = Join-Path $testRoot 'DevicesAD.csv'
    $intunePath = Join-Path $testRoot 'DevicesIntune.csv'
    $entraPath = Join-Path $testRoot 'DevicesEntra.csv'
    $adRows | Export-Csv -LiteralPath $adPath -NoTypeInformation -Encoding UTF8
    $intuneRows | Export-Csv -LiteralPath $intunePath -NoTypeInformation -Encoding UTF8
    $entraRows | Export-Csv -LiteralPath $entraPath -NoTypeInformation -Encoding UTF8

    $progressStages = New-Object System.Collections.Generic.List[string]
    $preview = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath `
        -ToolkitRoot $testRoot -LotName 'LOT-AUTO-IHJ-TEST' -NoEvidence `
        -ProgressCallback { param($stage,$detail) $progressStages.Add([string]$stage) }

    Assert-Equal $preview.Summary.SelectedDevices 4 'default preview selects four safe AD clients'
    Assert-Equal $preview.Summary.NeedsHybridJoin 2 'missing ServerAd evidence is classified as needing Hybrid Join'
    Assert-Equal $preview.Summary.HybridJoinPending 1 'pending ServerAd evidence is classified'
    Assert-Equal $preview.Summary.NeedsIntuneEnrollment 1 'registered ServerAd without Intune is classified'
    Assert-Equal $preview.Summary.IntunePresentExcluded 1 'Intune presence is always excluded'
    Assert-Equal $preview.Summary.ADDisabledExcluded 1 'disabled AD object is excluded'
    Assert-Equal $preview.Summary.WindowsServerExcluded 1 'Windows Server is excluded'
    Assert-Equal $preview.Summary.ADNameCollisions 1 'AD short-name collision is excluded'
    Assert-Equal $preview.Summary.EntraAmbiguousExcluded 1 'multiple ServerAd objects are excluded'
    Assert-True ($progressStages -contains 'Automatic LOT preview ready.') 'preview completion is reported'

    $filtered = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath `
        -ToolkitRoot $testRoot -LotName 'LOT-AUTO-IHJ-FILTERED' -ComputerNamePrefix 'FR-' -ComputerNameContains '-PC' -NoEvidence
    Assert-Equal $filtered.Summary.SelectedDevices 3 'prefix and contains filters are combined with AND'
    Assert-Equal $filtered.Summary.FilterExcludedDevices 5 'nonmatching AD keys are recorded once as filter exclusions'

    $stale = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath `
        -ToolkitRoot $testRoot -LotName 'LOT-AUTO-IHJ-STALE' -ExcludeStaleAd -AdLastLogonMaxAgeDays 45 -NoEvidence
    Assert-Equal $stale.Summary.ADLastLogonFilterExcluded 1 'stale AD device is excluded by the optional 45-day filter'

    $withoutEntra = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath `
        -ToolkitRoot $testRoot -LotName 'LOT-AUTO-IHJ-NO-ENTRA' -NoEvidence
    Assert-Equal $withoutEntra.Summary.EntraInventoryAvailable $false 'Entra is optional'
    Assert-Equal $withoutEntra.Summary.EntraUnavailableSelected 5 'safe selected devices are marked when Entra is unavailable'

    $testScriptsRoot = Join-Path $testRoot 'Scripts'
    New-Item -ItemType Directory -Path $testScriptsRoot -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $toolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1') -Destination $testScriptsRoot -Force
    Copy-Item -LiteralPath (Join-Path $toolkitRoot 'Lots\LOT-TEMPLATE') -Destination (Join-Path $testRoot 'Lots') -Recurse -Force
    $createProgressStages = New-Object System.Collections.Generic.List[string]
    $created = & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath `
        -ToolkitRoot $testRoot -LotName 'pilot:fr/ihj' -EvidenceRoot (Join-Path $testRoot 'Runs\AutomaticLotInventory') `
        -Create -ProgressCallback { param($stage,$detail) $createProgressStages.Add([string]$stage) }
    Assert-Equal $created.Summary.LotName 'LOT-pilot-fr-ihj' 'invalid folder characters are sanitized'
    Assert-True (Test-Path -LiteralPath $created.Summary.ComputersPath -PathType Leaf) 'Computers.txt is created'
    Assert-True (Test-Path -LiteralPath (Join-Path $created.Summary.LotPath 'AdDomain.txt') -PathType Leaf) 'AdDomain.txt is created'
    Assert-True (Test-Path -LiteralPath (Join-Path $created.Summary.LotPath 'Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd') -PathType Leaf) 'standard LOT wrappers are refreshed'
    Assert-True ($createProgressStages -contains 'Refreshing LOT command wrappers...') 'wrapper refresh is reported to the progress popup'
    Assert-True ($createProgressStages -contains 'Automatic LOT created.') 'creation completion is reported to the progress popup'
    Assert-Equal @(Get-Content -LiteralPath $created.Summary.ComputersPath).Count 4 'created Computers.txt contains selected devices only'
    foreach ($name in @('DevicesAD.csv','DevicesIntune.csv','DevicesEntra.csv','AutomaticLotSelection.csv','AutomaticLotExclusions.csv','AutomaticLotFilterExclusions.csv','AutomaticLotSummary.json')) {
        Assert-True (Test-Path -LiteralPath (Join-Path $created.Summary.EvidencePath $name) -PathType Leaf) "evidence file exists: $name"
    }

    $invalidFilterThrown = $false
    try {
        & $engine -AdInventoryCsv $adPath -IntuneInventoryCsv $intunePath -ToolkitRoot $testRoot `
            -LotName 'LOT-INVALID' -ComputerNamePrefix 'FR-*' -NoEvidence | Out-Null
    }
    catch { $invalidFilterThrown = $true }
    Assert-True $invalidFilterThrown 'wildcard prefix is rejected'

    $supportWithoutInitialization = [regex]::Replace(
        (Get-Content -LiteralPath $guiSupport -Raw),
        '(?m)^Initialize-AutomaticLotGui\s*$',
        ''
    )
    $supportTestPath = Join-Path $testRoot 'AutomaticLotGuiSupport.functions.ps1'
    [System.IO.File]::WriteAllText($supportTestPath, $supportWithoutInitialization, [System.Text.UTF8Encoding]::new($false))
    . $supportTestPath
    $cacheTenant = '00000000-0000-0000-0000-000000000001'
    $cacheIntunePath = Join-Path $testRoot 'CacheIntune.csv'
    @([pscustomobject]@{
        ComputerName='CACHE-PC'; DeviceName='CACHE-PC'; IntuneInventoryPresent=$true
        InventoryTenantId=$cacheTenant; InventoryAuthenticationMode='DelegatedInteractive'; InventoryScope='AllManagedDevices'
    }) | Export-Csv -LiteralPath $cacheIntunePath -NoTypeInformation -Encoding UTF8
    $cacheIntuneInfo = Get-AutomaticInventoryFileInfo -Path $cacheIntunePath -FreshnessHours 2 -SourceName Intune -ExpectedTenantId $cacheTenant
    Assert-True $cacheIntuneInfo.Fresh 'fresh delegated full Intune cache is reusable'
    $wrongTenantInfo = Get-AutomaticInventoryFileInfo -Path $cacheIntunePath -FreshnessHours 2 -SourceName Intune -ExpectedTenantId 'different-tenant'
    Assert-True (-not $wrongTenantInfo.Fresh) 'tenant-mismatched Intune cache is rejected'
    $badScopePath = Join-Path $testRoot 'CacheIntuneScoped.csv'
    @([pscustomobject]@{
        ComputerName='CACHE-PC'; DeviceName='CACHE-PC'; IntuneInventoryPresent=$true
        InventoryTenantId=$cacheTenant; InventoryAuthenticationMode='DelegatedInteractive'; InventoryScope='RequestedComputers'
    }) | Export-Csv -LiteralPath $badScopePath -NoTypeInformation -Encoding UTF8
    $badScopeInfo = Get-AutomaticInventoryFileInfo -Path $badScopePath -FreshnessHours 2 -SourceName Intune -ExpectedTenantId $cacheTenant
    Assert-True (-not $badScopeInfo.Fresh) 'targeted Intune cache is rejected for automatic full-inventory selection'
    $guiText = Get-Content -LiteralPath $gui -Raw
    $supportText = Get-Content -LiteralPath $guiSupport -Raw
    Assert-True ($guiText -match '<TabItem Header="Automatic LOT">') 'GUI declares the Automatic LOT tab'
    Assert-True ($guiText -match 'AutomaticCreateButton.+Content="Create"') 'GUI Create button does not claim to launch'
    Assert-True ($supportText -match 'The LOT will be created but not launched') 'confirmation explicitly states create-only behavior'
    Assert-True ($supportText -notmatch 'Start-ToolkitLot') 'automatic GUI support never launches the created LOT'

    Write-Output 'PASS: Intune Hybrid Join Automatic LOT synthetic tests completed.'
}
finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}
