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
    $strictAdPath = Join-Path $testRoot 'DevicesAD-StrictClientOS.csv'
    $strictAdRows = @(
        [pscustomobject]@{ ComputerName='STRICT-W10'; Name='STRICT-W10'; ADInventoryPresent=$true; Enabled=$true; OperatingSystem='Windows 10 Enterprise'; OperatingSystemVersion='10.0 (19045)' }
        [pscustomobject]@{ ComputerName='STRICT-W11'; Name='STRICT-W11'; ADInventoryPresent=$true; Enabled=$true; OperatingSystem='Windows 11 Enterprise'; OperatingSystemVersion='10.0 (22631)' }
        [pscustomobject]@{ ComputerName='STRICT-GENERIC-SERVER'; Name='STRICT-GENERIC-SERVER'; ADInventoryPresent=$true; Enabled=$true; OperatingSystem='Windows'; OperatingSystemVersion='10.0 (20348)' }
        [pscustomobject]@{ ComputerName='STRICT-BLANK-SERVER'; Name='STRICT-BLANK-SERVER'; ADInventoryPresent=$true; Enabled=$true; OperatingSystem=''; OperatingSystemVersion='10.0 (26100)' }
    )
    $strictAdRows | Export-Csv -LiteralPath $strictAdPath -NoTypeInformation -Encoding UTF8
    $strictPreview = & $engine -AdInventoryCsv $strictAdPath -IntuneInventoryCsv $intunePath -EntraInventoryCsv $entraPath -ToolkitRoot $testRoot -LotName 'LOT-AUTO-IHJ-STRICT-OS' -NoEvidence
    Assert-Equal $strictPreview.Summary.SelectedDevices 2 'only explicitly identified Windows 10 and Windows 11 clients are selected'
    Assert-Equal $strictPreview.Summary.UnknownOSExcluded 2 'generic or blank AD operating systems remain excluded even when the build resembles Windows client'
    $strictSelectedNames = @($strictPreview.SelectedDevices | Select-Object -ExpandProperty ComputerName)
    Assert-True ($strictSelectedNames -contains 'STRICT-W10') 'explicit Windows 10 client remains eligible'
    Assert-True ($strictSelectedNames -contains 'STRICT-W11') 'explicit Windows 11 client remains eligible'
    Assert-True (-not ($strictSelectedNames -contains 'STRICT-GENERIC-SERVER' -or $strictSelectedNames -contains 'STRICT-BLANK-SERVER')) 'version-only server-like rows are never selected'


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
    Assert-True ($supportText -notmatch '\$controls\.AutomaticCreateButton\.IsEnabled\s*=\s*\$false') 'automatic LOT Create remains available before preview and after filter changes'
    $createHandler = [regex]::Match(
        $supportText,
        '(?s)\$controls\.AutomaticCreateButton\.Add_Click\(\{(?<Body>.*?)\r?\n    \}\)\r?\n    \$controls\.AutomaticOpenEvidenceButton'
    )
    Assert-True $createHandler.Success 'automatic LOT Create handler is present'
    Assert-True ($createHandler.Groups['Body'].Value -match '\$null -eq \$script:AutomaticPreviewResult') 'automatic LOT Create detects a missing preview'
    Assert-True ($createHandler.Groups['Body'].Value -match 'Update-AutomaticLotPreview\s+-ForceInventoryRefresh:\$force') 'automatic LOT Create refreshes a missing, stale, or forced preview'


    Write-Output 'PASS: Intune Hybrid Join Automatic LOT synthetic tests completed.'
}
finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC2iuxNxwTP2HbL
# UkRFEi+V5fUxYhA4VqBHInlSXsXnHKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBtT4DihkXHzJbdkd2AOZU+NPnh9By6ysjjwmreIi6ppMA0GCSqG
# SIb3DQEBAQUABIIBgFHGPPqC5SJqOyQVkTtsab3kQy6rU7QQF/XsxTgkyUeZm3sX
# YjG7RGitfMUe7EHC+Y3Xd2064x/lfO6nEgzhvIPhX0gZfd3Lrxaz83QZ6fcxm0Be
# Y15JLEAsUxjJwHLddgziPLzVHXpx2vrJB4Tj/itl25DOMXkYqcBg3wae8eonGDgq
# PrXHgbObgboMPdn4lK3o/Ss8gRly6FQhpWt01o0rgK289EZlUTBsn0amSBFHhp4y
# kR/laphIyCWccGkeYC4o2miFVDgjRxFVJ3rlZit5YezceUkzVPhaSebNMJxZc5z4
# uLHeNPr80lmeKe9CU1EFBcVjJGRAsT8UuYH7gfwBClzerzmpecG+4RHUojgCxowz
# B7i8Izrz+hmM2UT0TxwfgfDTThnBaqXHpu9xXm3u6EanYlZjh0y8KmCqbZDUj/TD
# xlhaPchMx2q3tjSXRi3Fj7u/OI2DfHFWiu4uFVA6d8JnZ8rq3xRUf5un68biYNT6
# +smIV4K9Y2lw3JigUaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjgyMDU2
# NDZaMC8GCSqGSIb3DQEJBDEiBCAV/3ngWZVJ/4jxstuHYQNSNQw7Bloh7+FBv/BD
# fSVwNjANBgkqhkiG9w0BAQEFAASCAgCPqmqr3BT6wKDS9ddSplPcFK1wGXFi2ftr
# F2UXDQqgUqV34KDknEUNDuwnvfLfp7L0bQ4yMpnOgV5iR8jsYHpgDtdlAwG2jiiX
# y6j6uM9DpZHcdA7B6tIK09yfF7S7OEiulJOmfIs5/WvqEznzn0TDUCRM+32Gdv9A
# gBz+JriE+7VUvgHV0h80WROyrJvvu6kzuO2ZwLTiVMEXjOyMvZr6/FfDwbFE3Wg7
# OxyFxCNNNlTvF60ZFh8k/UihH+8cyYN0MrS8gtJPrSDMv7MLeMZewZ9EJtRhMYwo
# ryags7j5MjaeWezweTUZjyakoZIzzGseEtQvXNrXP4/82bz5r2TC7HNvYloNPDsY
# yQ36fSm0RDiN496G6cwuatYac+FEJeOD9tQUAPmNE5l0lfyijZI4pEHZIU38ahx2
# HTiZ8A6ah4/WVFRzxXp/gw9gBCKi0IMnepZ9PoxPASrpkLMQ/OwdtPsfyr9PWN/d
# 8VhB2RzrXXyEXb3OAiYgDIbKeZMR1rbQE+pspWzqQrGpII77j7kn6leOaLLU70IU
# uYhaPqqrjJ04aWVzwiaJvxge0OPRUikaFvxsqIDWQGMTVLHywyQIHyecfoQdyIS6
# senHJY1hMG+ubNshuQz0UZDMcnXW4YgDL850TrbIL/qZ5BdoUw1U51LYW5Du06na
# Pa0mAOSxQQ==
# SIG # End signature block
