<#
.SYNOPSIS
Builds a guarded HybridJoin Wave from AD, Intune, and Entra inventories.

.DESCRIPTION
Selects enabled Windows 10 and Windows 11 client computers from Active Directory,
keeps name-only Intune matches for local verification, excludes ambiguous inventory states, classifies
the remaining devices with Entra evidence, writes selection evidence under
Runs\HybridJoin\_AutomaticWave, and can create a standard HybridJoin Wave through
New-SemtWave without launching it. Default inventory inputs are the shared caches
in <Package>\Inventory.
Provenance: copied from SmartM365/Devices/IntuneHybridJoinToolkit/Scripts/SmartM365-IntuneHybridJoinRepair-New-AutomaticLot.ps1 1.0.2 (see Docs\PROVENANCE.json.txt).

.PARAMETER AdInventoryCsv
AD inventory CSV. Defaults to the shared cache <Package>\Inventory\DevicesAD.csv.

.PARAMETER IntuneInventoryCsv
Intune inventory CSV. Defaults to the shared cache <Package>\Inventory\DevicesIntune.csv.

.PARAMETER EntraInventoryCsv
Optional Entra inventory CSV. Defaults to the shared cache <Package>\Inventory\DevicesEntra.csv when present.

.PARAMETER PackageRoot
Smart Endpoint Migration Toolkit package root. Defaults to the package that contains this script.

.PARAMETER WaveName
Wave name. Defaults to WAVE-AUTO-HJ-<prefix>-<yyyyMMdd-HHmmss>.

.PARAMETER AdDomainFilter
Optional AD domain DNS names (for example contoso.local). Only computers with an AD record in one of
these domains are kept; the others are listed in AutomaticWaveFilterExclusions.csv with
AD_DOMAIN_NOT_MATCHED, or AD_DOMAIN_UNKNOWN when they have no AD record or no domain. Each domain must
exist in the AD inventory. The shared AD inventory always covers the whole forest.

.PARAMETER Create
Creates the Wave (Computers.txt, Wave.local.json, wrappers). Without it, only evidence is written.

.PARAMETER PreviewEvidencePath
Evidence folder of a Preview (run without -Create). With -Create, the Wave is created from that
previewed selection without analyzing the inventories again, when the Preview used the same filters,
the inventory files have not changed since, and it is less than 24 hours old. Otherwise the selection
is computed again.

.VERSION
1.0.0

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$AdInventoryCsv,
    [string]$IntuneInventoryCsv,
    [string]$EntraInventoryCsv,
    [string]$PackageRoot,
    [string]$WaveName,
    [string[]]$ComputerNamePrefix,
    [string[]]$ComputerNameContains,
    [string[]]$AdDomainFilter,
    [switch]$ExcludeStaleAd,
    [ValidateRange(1, 3650)][int]$AdLastLogonMaxAgeDays = 45,
    [scriptblock]$ProgressCallback,
    [string]$EvidenceRoot,
    [switch]$Create,
    [string]$PreviewEvidencePath,
    [switch]$NoEvidence
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:AutomaticWaveVersion = '1.0.0'
$script:AutomaticWaveCampaign = 'HybridJoin'
$script:AutomaticWaveProgressCallback = $ProgressCallback

$script:AutomaticWaveScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$coreModulePath = Join-Path $script:AutomaticWaveScriptRoot '..\..\..\Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1'
Import-Module ([System.IO.Path]::GetFullPath($coreModulePath)) -MinimumVersion 1.0.0 -Force

function Publish-AutomaticWaveProgress {
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [string]$Detail = ''
    )

    if ($script:AutomaticWaveProgressCallback) {
        & $script:AutomaticWaveProgressCallback $Stage $Detail | Out-Null
    }
}

function Get-AutomaticWavePackageRoot {
    param([AllowNull()][string]$RequestedRoot)

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        return [System.IO.Path]::GetFullPath($RequestedRoot)
    }
    return (Get-SemtPackageRoot -StartPath $script:AutomaticWaveScriptRoot)
}

function Get-AutomaticWaveSafeName {
    param(
        [AllowNull()][string]$Name,
        [AllowNull()][string[]]$NamePrefixes
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $segments = [System.Collections.Generic.List[string]]::new()
        foreach ($prefix in @($NamePrefixes)) {
            $segment = [regex]::Replace(([string]$prefix).ToUpperInvariant(), '[^A-Z0-9_-]+', '-').Trim('-_')
            if (-not [string]::IsNullOrWhiteSpace($segment) -and -not $segments.Contains($segment)) {
                $segments.Add($segment)
            }
        }
        $prefixSegment = if ($segments.Count -gt 0) { "-$($segments -join '-')" } else { '' }
        $Name = 'WAVE-AUTO-HJ{0}-{1}' -f $prefixSegment, (Get-Date -Format 'yyyyMMdd-HHmmss')
    }

    $safeName = [regex]::Replace($Name.Trim(), '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        throw 'The automatic Wave name is empty after sanitization.'
    }
    if ($safeName -notmatch '^(?i)WAVE-') {
        $safeName = "WAVE-$safeName"
    }
    # Wave names are limited to 80 characters by the engine (Test-SemtWaveName).
    if ($safeName.Length -gt 80) {
        $safeName = $safeName.Substring(0, 80).TrimEnd('-._')
    }
    return $safeName
}

function Get-InventoryValue {
    param(
        [AllowNull()]$Row,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    if ($null -eq $Row) { return $null }
    foreach ($name in $Names) {
        if ($Row -is [System.Collections.IDictionary]) {
            if ($Row.Contains($name)) { return $Row[$name] }
            continue
        }
        $property = $Row.PSObject.Properties[$name]
        if ($property) { return $property.Value }
    }
    return $null
}

function ConvertTo-InventoryBoolean {
    param(
        [AllowNull()]$Value,
        [bool]$Default = $false
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $Default }
    if ($Value -is [bool]) { return [bool]$Value }
    $text = ([string]$Value).Trim()
    if ($text -match '^(?i:true|1|yes|enabled)$') { return $true }
    if ($text -match '^(?i:false|0|no|disabled)$') { return $false }
    return $Default
}

function ConvertTo-InventoryDate {
    param([AllowNull()]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return [datetime]::MinValue
    }

    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse(
        [string]$Value,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::AllowWhiteSpaces -bor [System.Globalization.DateTimeStyles]::AssumeUniversal,
        [ref]$parsed
    )) {
        return $parsed.ToUniversalTime()
    }
    if ([datetime]::TryParse([string]$Value, [ref]$parsed)) {
        return $parsed.ToUniversalTime()
    }
    return [datetime]::MinValue
}

function Get-InventoryShortName {
    param([AllowNull()][string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    return ($Name.Trim().Trim([char]34).Split('.')[0]).ToUpperInvariant()
}

function Get-AutomaticWaveNamePrefixes {
    param([AllowNull()][string[]]$Prefixes)

    $normalized = [System.Collections.Generic.List[string]]::new()
    foreach ($item in @($Prefixes)) {
        foreach ($candidate in @(([string]$item) -split ';')) {
            $value = $candidate.Trim().Trim([char]34)
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            $prefix = Get-InventoryShortName -Name $value
            if ($prefix -notmatch '^[A-Z0-9_-]+$') {
                throw "Invalid computer name prefix '$value'. Use only letters, digits, hyphens, or underscores."
            }
            if (-not $normalized.Contains($prefix)) { $normalized.Add($prefix) }
        }
    }
    return @($normalized.ToArray())
}

function Get-AutomaticWaveNameContainsValues {
    param([AllowNull()][string[]]$ContainsValues)

    $normalized = [System.Collections.Generic.List[string]]::new()
    foreach ($item in @($ContainsValues)) {
        foreach ($candidate in @(([string]$item) -split ';')) {
            $value = $candidate.Trim().Trim([char]34).ToUpperInvariant()
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            if ($value -notmatch '^[A-Z0-9_-]+$') {
                throw "Invalid computer name contains value '$candidate'. Use only letters, digits, hyphens, or underscores."
            }
            if (-not $normalized.Contains($value)) { $normalized.Add($value) }
        }
    }
    return @($normalized.ToArray())
}

function Test-AutomaticWaveNamePrefix {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerKey,
        [string[]]$Prefixes
    )

    if (@($Prefixes).Count -eq 0) { return $true }
    foreach ($prefix in @($Prefixes)) {
        if ($ComputerKey.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Test-AutomaticWaveNameContains {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerKey,
        [string[]]$ContainsValues
    )

    if (@($ContainsValues).Count -eq 0) { return $true }
    foreach ($value in @($ContainsValues)) {
        if ($ComputerKey.IndexOf($value, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
}

function Get-AutomaticWaveAdDomains {
    param([AllowNull()][string[]]$Domains)

    $normalized = [System.Collections.Generic.List[string]]::new()
    foreach ($item in @($Domains)) {
        foreach ($candidate in @(([string]$item) -split ';')) {
            $value = $candidate.Trim().Trim([char]34).Trim('.').ToLowerInvariant()
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            if ($value -notmatch '^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)*$') {
                throw "Invalid AD domain '$candidate'. Use the DNS name of the domain, for example contoso.local."
            }
            if (-not $normalized.Contains($value)) { $normalized.Add($value) }
        }
    }
    return @($normalized.ToArray())
}

function Get-AutomaticWaveAdDomainFilterReason {
    <#
    Empty when the computer has an AD record in one of the selected domains (or no domain filter).
    A computer without an AD record or without a domain value never matches a domain filter.
    #>
    param(
        [object[]]$AdRecords,
        [string[]]$Domains
    )

    if (@($Domains).Count -eq 0) { return '' }
    $hasDomain = $false
    foreach ($record in @($AdRecords)) {
        if ($null -eq $record) { continue }
        $domain = ([string]$record.Domain).Trim().Trim('.').ToLowerInvariant()
        if (-not $domain) { continue }
        $hasDomain = $true
        if ($Domains -contains $domain) { return '' }
    }
    if (-not $hasDomain) { return 'AD_DOMAIN_UNKNOWN' }
    return 'AD_DOMAIN_NOT_MATCHED'
}

function Assert-AutomaticWaveAdDomainsInInventory {
    param(
        [string[]]$Domains,
        [object[]]$AdRecords,
        [string]$InventoryPath
    )

    if (@($Domains).Count -eq 0) { return }
    $knownSet = @{}
    foreach ($record in @($AdRecords)) {
        $domain = ([string]$record.Domain).Trim().Trim('.').ToLowerInvariant()
        if ($domain) { $knownSet[$domain] = $true }
    }
    $known = @($knownSet.Keys | Sort-Object)
    $missing = @($Domains | Where-Object { $known -notcontains $_ })
    if ($missing.Count -gt 0) {
        throw ("SEMT_AD_DOMAIN_NOT_IN_INVENTORY: AD domain(s) {0} not found in the AD inventory {1}. Domains in the inventory: {2}." -f ($missing -join ', '),$InventoryPath,$(if ($known.Count -gt 0) { $known -join ', ' } else { 'none' }))
    }
}

function Get-WindowsClientClassification {
    param(
        [AllowNull()][string]$OperatingSystem
    )

    $osText = ([string]$OperatingSystem).Trim()
    if ($osText -match '(?i)Server') { return 'WindowsServer' }
    if (-not [string]::IsNullOrWhiteSpace($osText) -and $osText -notmatch '(?i)Windows') { return 'Unsupported' }
    if ($osText -match '(?i)Windows\s*11\b') { return 'Windows11' }
    if ($osText -match '(?i)Windows\s*10\b') { return 'Windows10' }
    return 'Unknown'
}

function Resolve-InventoryColumn {
    <#
    First of Names present in the inventory rows. CSV rows share one header, so the column is resolved
    once per file instead of once per row (tens of thousands of rows in a large forest).
    #>
    param(
        [object[]]$Rows,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    if (@($Rows).Count -eq 0 -or $null -eq $Rows[0]) { return $null }
    $first = $Rows[0]
    foreach ($name in $Names) {
        if ($first -is [System.Collections.IDictionary]) {
            if ($first.Contains($name)) { return $name }
            continue
        }
        if ($first.PSObject.Properties[$name]) { return $name }
    }
    return $null
}

function Get-AdAutomaticWaveRecord {
    param([object[]]$Rows)

    # Inline conversions (same rules as Get-InventoryShortName, ConvertTo-InventoryBoolean with a
    # default of $true, and ConvertTo-InventoryDate): one function call per row costs minutes here.
    $rows = @($Rows)
    $colName = Resolve-InventoryColumn -Rows $rows -Names @('ComputerName', 'Name')
    $colDns = Resolve-InventoryColumn -Rows $rows -Names @('DNSHostName')
    $colOs = Resolve-InventoryColumn -Rows $rows -Names @('OperatingSystem')
    $colOsVersion = Resolve-InventoryColumn -Rows $rows -Names @('OperatingSystemVersion', 'OSVersion')
    $colPresent = Resolve-InventoryColumn -Rows $rows -Names @('ADInventoryPresent')
    $colEnabled = Resolve-InventoryColumn -Rows $rows -Names @('Enabled')
    $colLastSeen = Resolve-InventoryColumn -Rows $rows -Names @('LastLogonTimestampUtc')
    $colDomain = Resolve-InventoryColumn -Rows $rows -Names @('ADDomain')
    $colDn = Resolve-InventoryColumn -Rows $rows -Names @('DistinguishedName')
    $classificationCache = @{}
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $dateStyles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces -bor [System.Globalization.DateTimeStyles]::AssumeUniversal

    foreach ($row in $rows) {
        if ($null -eq $row) { continue }
        $computerName = if ($colName) { [string]$row.$colName } else { '' }
        $dnsHostName = if ($colDns) { [string]$row.$colDns } else { '' }
        $identity = if (-not [string]::IsNullOrWhiteSpace($dnsHostName)) { $dnsHostName } else { $computerName }
        if ([string]::IsNullOrWhiteSpace($identity)) { continue }
        $key = $identity.Trim().Trim([char]34).Split('.')[0].ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($key)) { continue }

        $operatingSystem = if ($colOs) { [string]$row.$colOs } else { '' }
        $operatingSystemVersion = if ($colOsVersion) { [string]$row.$colOsVersion } else { '' }
        $present = $true
        if ($colPresent) { $text = [string]$row.$colPresent; if (-not [string]::IsNullOrWhiteSpace($text) -and $text.Trim() -match '^(?i:false|0|no|disabled)$') { $present = $false } }
        $enabled = $true
        if ($colEnabled) { $text = [string]$row.$colEnabled; if (-not [string]::IsNullOrWhiteSpace($text) -and $text.Trim() -match '^(?i:false|0|no|disabled)$') { $enabled = $false } }
        $lastSeen = [datetime]::MinValue
        if ($colLastSeen) {
            $text = [string]$row.$colLastSeen
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $parsed = [datetime]::MinValue
                if ([datetime]::TryParse($text, $invariant, $dateStyles, [ref]$parsed) -or [datetime]::TryParse($text, [ref]$parsed)) { $lastSeen = $parsed.ToUniversalTime() }
            }
        }
        if (-not $classificationCache.ContainsKey($operatingSystem)) { $classificationCache[$operatingSystem] = Get-WindowsClientClassification -OperatingSystem $operatingSystem }
        [pscustomobject]@{
            Key                = $key
            Name               = $computerName
            DNSHostName        = $dnsHostName
            PreferredName      = $identity
            Domain             = if ($colDomain) { [string]$row.$colDomain } else { '' }
            DistinguishedName  = if ($colDn) { [string]$row.$colDn } else { '' }
            Present            = $present
            Enabled            = $enabled
            Classification     = $classificationCache[$operatingSystem]
            OperatingSystem    = $operatingSystem
            OSVersion          = $operatingSystemVersion
            LastSeenUtc        = $lastSeen
        }
    }
}

function Get-IntuneAutomaticWaveRecord {
    param([object[]]$Rows)

    $rows = @($Rows)
    $colIdentity = Resolve-InventoryColumn -Rows $rows -Names @('DeviceName', 'ManagedDeviceName', 'ComputerName')
    $colPresent = Resolve-InventoryColumn -Rows $rows -Names @('IntuneInventoryPresent')
    $colId = Resolve-InventoryColumn -Rows $rows -Names @('IntuneManagedDeviceId', 'Id')
    $colLastSync = Resolve-InventoryColumn -Rows $rows -Names @('LastSyncDateTime')
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $dateStyles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces -bor [System.Globalization.DateTimeStyles]::AssumeUniversal

    foreach ($row in $rows) {
        if ($null -eq $row -or -not $colIdentity) { continue }
        $identity = [string]$row.$colIdentity
        if ([string]::IsNullOrWhiteSpace($identity)) { continue }
        $key = $identity.Trim().Trim([char]34).Split('.')[0].ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $present = $true
        if ($colPresent) { $text = [string]$row.$colPresent; if (-not [string]::IsNullOrWhiteSpace($text) -and $text.Trim() -match '^(?i:false|0|no|disabled)$') { $present = $false } }
        $lastSync = [datetime]::MinValue
        if ($colLastSync) {
            $text = [string]$row.$colLastSync
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $parsed = [datetime]::MinValue
                if ([datetime]::TryParse($text, $invariant, $dateStyles, [ref]$parsed) -or [datetime]::TryParse($text, [ref]$parsed)) { $lastSync = $parsed.ToUniversalTime() }
            }
        }
        [pscustomobject]@{
            Key             = $key
            DeviceName      = $identity
            Present         = $present
            ManagedDeviceId = if ($colId) { [string]$row.$colId } else { '' }
            LastSyncUtc     = $lastSync
        }
    }
}

function Get-EntraAutomaticWaveRecord {
    param([object[]]$Rows)

    $rows = @($Rows)
    $colIdentity = Resolve-InventoryColumn -Rows $rows -Names @('DisplayName', 'ComputerName')
    $colPresent = Resolve-InventoryColumn -Rows $rows -Names @('EntraInventoryPresent')
    $colTrust = Resolve-InventoryColumn -Rows $rows -Names @('TrustType')
    $colState = Resolve-InventoryColumn -Rows $rows -Names @('EntraRegisteredState')
    $colAltSec = Resolve-InventoryColumn -Rows $rows -Names @('AlternativeSecurityIdCount')
    $colEnabled = Resolve-InventoryColumn -Rows $rows -Names @('AccountEnabled')
    $colDeviceId = Resolve-InventoryColumn -Rows $rows -Names @('DeviceId')
    $colObjectId = Resolve-InventoryColumn -Rows $rows -Names @('EntraObjectId', 'Id')

    foreach ($row in $rows) {
        if ($null -eq $row -or -not $colIdentity) { continue }
        $identity = [string]$row.$colIdentity
        if ([string]::IsNullOrWhiteSpace($identity)) { continue }
        $key = $identity.Trim().Trim([char]34).Split('.')[0].ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $present = $true
        if ($colPresent) { $text = [string]$row.$colPresent; if (-not [string]::IsNullOrWhiteSpace($text) -and $text.Trim() -match '^(?i:false|0|no|disabled)$') { $present = $false } }
        $accountEnabled = $true
        if ($colEnabled) { $text = [string]$row.$colEnabled; if (-not [string]::IsNullOrWhiteSpace($text) -and $text.Trim() -match '^(?i:false|0|no|disabled)$') { $accountEnabled = $false } }
        [pscustomobject]@{
            Key                       = $key
            DisplayName               = $identity
            Present                   = $present
            TrustType                 = if ($colTrust) { [string]$row.$colTrust } else { '' }
            RegisteredState           = if ($colState) { [string]$row.$colState } else { '' }
            AlternativeSecurityIdCount = if ($colAltSec) { [string]$row.$colAltSec } else { '' }
            AccountEnabled            = $accountEnabled
            DeviceId                  = if ($colDeviceId) { [string]$row.$colDeviceId } else { '' }
            ObjectId                  = if ($colObjectId) { [string]$row.$colObjectId } else { '' }
        }
    }
}

function Get-AutomaticWaveInventoryFingerprint {
    <#
    Identity of the inventory files read by a selection (full path, size, last write time). Create
    reuses a Preview only when these files have not changed since that Preview.
    #>
    param([string[]]$Paths)

    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($path in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        $item = Get-Item -LiteralPath $path
        $parts.Add(('{0}|{1}|{2}' -f $item.FullName.ToLowerInvariant(), $item.Length, $item.LastWriteTimeUtc.Ticks))
    }
    return ($parts.ToArray() -join ';')
}

function Get-AutomaticWavePreviewReuse {
    <#
    Checks whether Create can reuse a Preview instead of computing the selection again: same campaign,
    engine version and filters, unchanged inventory files, less than 24 hours old, selection report
    present. Returns the preview summary (or $null) and the reason it cannot be reused.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$SelectionSignature,
        [Parameter(Mandatory = $true)][string]$InventoryFingerprint
    )

    $summaryPath = Join-Path $Path 'AutomaticWaveSummary.json'
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf) -or -not (Test-Path -LiteralPath (Join-Path $Path 'AutomaticWaveSelection.csv') -PathType Leaf)) {
        return [pscustomobject]@{ Summary = $null; Reason = 'the preview evidence is incomplete' }
    }
    $summary = Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $generatedValue = Get-SemtPropertyValue -InputObject $summary -Name 'GeneratedUtc' -Default $null
    $generatedUtc = [datetime]::MinValue
    if ($generatedValue -is [datetime]) { $generatedUtc = $generatedValue.ToUniversalTime() }
    elseif ([datetime]::TryParse([string]$generatedValue, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$generatedUtc)) { $generatedUtc = $generatedUtc.ToUniversalTime() }

    $reason = ''
    if ([string](Get-SemtPropertyValue -InputObject $summary -Name 'Campaign') -ne $script:AutomaticWaveCampaign -or [string](Get-SemtPropertyValue -InputObject $summary -Name 'Version') -ne $script:AutomaticWaveVersion) { $reason = 'it was written for another campaign or engine version' }
    elseif ([bool](Get-SemtPropertyValue -InputObject $summary -Name 'WaveCreated' -Default $true)) { $reason = 'the evidence is not a preview' }
    elseif ([string](Get-SemtPropertyValue -InputObject $summary -Name 'SelectionSignature') -ne $SelectionSignature) { $reason = 'the filters changed since the preview' }
    elseif ([string](Get-SemtPropertyValue -InputObject $summary -Name 'InventoryFingerprint') -ne $InventoryFingerprint) { $reason = 'an inventory changed since the preview' }
    elseif ($generatedUtc -lt [datetime]::UtcNow.AddHours(-24)) { $reason = 'the preview is older than 24 hours' }
    if ($reason) { return [pscustomobject]@{ Summary = $null; Reason = $reason } }
    return [pscustomobject]@{ Summary = $summary; Reason = '' }
}

function Assert-AutomaticWaveCreatable {
    param(
        [Parameter(Mandatory = $true)][int]$SelectedCount,
        [Parameter(Mandatory = $true)][string]$WaveName,
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$EmptySelectionMessage
    )

    if ($SelectedCount -eq 0) { throw $EmptySelectionMessage }
    if (-not (Test-SemtWaveName -Name $WaveName)) { throw ("SEMT_WAVE_NAME_INVALID: '{0}' is not a valid Wave name." -f $WaveName) }
    $plannedWavePath = Join-Path (Get-SemtWaveRoot -PackageRoot $PackageRoot -Campaign $script:AutomaticWaveCampaign) $WaveName
    if (Test-Path -LiteralPath $plannedWavePath) { throw ("SEMT_WAVE_EXISTS: Wave '{0}' already exists for {1}: {2}" -f $WaveName, $script:AutomaticWaveCampaign, $plannedWavePath) }
}

function New-AutomaticWaveFromPreview {
    <#
    Create from a reusable Preview: the Wave gets exactly the previewed selection without reading and
    analyzing the inventories again (minutes in a large forest). The preview reports and inventory
    copies are copied to the Create evidence.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$PreviewPath,
        [Parameter(Mandatory = $true)]$PreviewSummary,
        [Parameter(Mandatory = $true)][string]$WaveName,
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$DescriptionBase,
        [Parameter(Mandatory = $true)][string]$EmptySelectionMessage,
        [string]$EvidenceRoot,
        [switch]$NoEvidence
    )

    Publish-AutomaticWaveProgress -Stage 'Reusing the previewed selection...' -Detail $PreviewPath
    $selected = @(Import-Csv -LiteralPath (Join-Path $PreviewPath 'AutomaticWaveSelection.csv'))
    $excluded = @()
    $filterExcluded = @()
    if (Test-Path -LiteralPath (Join-Path $PreviewPath 'AutomaticWaveExclusions.csv') -PathType Leaf) { $excluded = @(Import-Csv -LiteralPath (Join-Path $PreviewPath 'AutomaticWaveExclusions.csv')) }
    if (Test-Path -LiteralPath (Join-Path $PreviewPath 'AutomaticWaveFilterExclusions.csv') -PathType Leaf) { $filterExcluded = @(Import-Csv -LiteralPath (Join-Path $PreviewPath 'AutomaticWaveFilterExclusions.csv')) }
    Assert-AutomaticWaveCreatable -SelectedCount $selected.Count -WaveName $WaveName -PackageRoot $PackageRoot -EmptySelectionMessage $EmptySelectionMessage

    $evidencePath = ''
    if (-not $NoEvidence) {
        if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
            $EvidenceRoot = Join-Path (Get-SemtStorageRoot -PackageRoot $PackageRoot -Kind Runs) ('{0}\_AutomaticWave' -f $script:AutomaticWaveCampaign)
        }
        $evidencePath = Join-Path $EvidenceRoot (Get-Date -Format 'yyyyMMdd-HHmmssfff')
        Publish-AutomaticWaveProgress -Stage 'Copying the preview evidence...' -Detail $evidencePath
        New-SemtDirectory -Path $evidencePath
        foreach ($file in @(Get-ChildItem -LiteralPath $PreviewPath -Filter '*.csv' -File)) {
            Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $evidencePath $file.Name) -Force
        }
    }

    Publish-AutomaticWaveProgress -Stage 'Creating the automatic Wave...' -Detail ("Wave: {0}; selected devices: {1}" -f $WaveName, $selected.Count)
    $waveComputers = @($selected | Sort-Object ComputerName | Select-Object -ExpandProperty ComputerName)
    $wave = New-SemtWave -PackageRoot $PackageRoot -Campaign $script:AutomaticWaveCampaign -Name $WaveName -Computers $waveComputers -Source Automatic -Description ('{0}; Selected={1}' -f $DescriptionBase, $selected.Count) -EvidencePath $evidencePath

    $values = [ordered]@{}
    foreach ($property in $PreviewSummary.PSObject.Properties) { $values[$property.Name] = $property.Value }
    $values['GeneratedUtc'] = [datetime]::UtcNow.ToString('o')
    $values['WaveCreated'] = $true
    $values['WaveName'] = $WaveName
    $values['WavePath'] = [string]$wave.Path
    $values['ComputersPath'] = [string]$wave.ComputersPath
    $values['EvidencePath'] = $evidencePath
    $values['PreviewEvidencePath'] = $PreviewPath
    $summary = [pscustomobject]$values
    if (-not $NoEvidence) {
        Write-SemtJsonFileAtomic -Path (Join-Path $evidencePath 'AutomaticWaveSummary.json') -InputObject $summary -Depth 5
    }
    Publish-AutomaticWaveProgress -Stage 'Automatic Wave created.' -Detail ([string]$wave.Path)

    return [pscustomobject]@{
        Summary = $summary
        SelectedDevices = $selected
        ExcludedDevices = $excluded
        FilterExcludedDevices = $filterExcluded
    }
}

function Export-AutomaticWaveCsv {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[]]$Columns,
        [object[]]$Rows
    )

    $items = @($Rows)
    if ($items.Count -gt 0) {
        $items | Select-Object -Property $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
        return
    }
    $header = @($Columns | ForEach-Object { '"{0}"' -f ([string]$_ -replace '"', '""') }) -join ','
    Set-Content -LiteralPath $Path -Value $header -Encoding UTF8
}

$effectivePackageRoot = Get-AutomaticWavePackageRoot -RequestedRoot $PackageRoot
if ([string]::IsNullOrWhiteSpace($AdInventoryCsv)) { $AdInventoryCsv = Get-SemtSharedInventoryPath -PackageRoot $effectivePackageRoot -Kind AD }
if ([string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) { $IntuneInventoryCsv = Get-SemtSharedInventoryPath -PackageRoot $effectivePackageRoot -Kind Intune }
if ([string]::IsNullOrWhiteSpace($EntraInventoryCsv)) { $EntraInventoryCsv = Get-SemtSharedInventoryPath -PackageRoot $effectivePackageRoot -Kind Entra }
foreach ($requiredPath in @($AdInventoryCsv, $IntuneInventoryCsv)) {
    if ([string]::IsNullOrWhiteSpace($requiredPath) -or -not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Required automatic Wave inventory not found: $requiredPath"
    }
}
$entraAvailable = (-not [string]::IsNullOrWhiteSpace($EntraInventoryCsv) -and (Test-Path -LiteralPath $EntraInventoryCsv -PathType Leaf))

$namePrefixes = @(Get-AutomaticWaveNamePrefixes -Prefixes $ComputerNamePrefix)
$nameContainsValues = @(Get-AutomaticWaveNameContainsValues -ContainsValues $ComputerNameContains)
$namePrefixDisplay = $namePrefixes -join ';'
$nameContainsDisplay = $nameContainsValues -join ';'
$adDomains = @(Get-AutomaticWaveAdDomains -Domains $AdDomainFilter)
$adDomainDisplay = $adDomains -join ';'
$availableSourcesText = if ($entraAvailable) { 'AD+Intune+Entra' } else { 'AD+Intune' }
$safeWaveName = Get-AutomaticWaveSafeName -Name $WaveName -NamePrefixes $namePrefixes
$emptySelectionMessage = 'No eligible Hybrid Join device was selected; the Wave was not created.'
$descriptionBase = 'Automatic HybridJoin Wave; Sources={0}; ComputerNamePrefix={1}; ComputerNameContains={2}; ADDomain={5}; ExcludeIntunePresent=False; ExcludeStaleAd={3}; ADLastLogonMaxAgeDays={4}' -f $availableSourcesText, $namePrefixDisplay, $nameContainsDisplay, [bool]$ExcludeStaleAd, $AdLastLogonMaxAgeDays, $adDomainDisplay
$selectionSignature = @($script:AutomaticWaveCampaign, $script:AutomaticWaveVersion, "Sources=$availableSourcesText", "ComputerNamePrefix=$namePrefixDisplay", "ComputerNameContains=$nameContainsDisplay", "ADDomain=$adDomainDisplay", "ExcludeStaleAd=$([bool]$ExcludeStaleAd)", "ADLastLogonMaxAgeDays=$AdLastLogonMaxAgeDays") -join '|'
$inventoryFingerprint = Get-AutomaticWaveInventoryFingerprint -Paths @($AdInventoryCsv, $IntuneInventoryCsv, $(if ($entraAvailable) { $EntraInventoryCsv } else { '' }))

if ($Create -and -not [string]::IsNullOrWhiteSpace($PreviewEvidencePath)) {
    $previewReuse = Get-AutomaticWavePreviewReuse -Path $PreviewEvidencePath -SelectionSignature $selectionSignature -InventoryFingerprint $inventoryFingerprint
    if ($previewReuse.Summary) {
        New-AutomaticWaveFromPreview -PreviewPath $PreviewEvidencePath -PreviewSummary $previewReuse.Summary -WaveName $safeWaveName -PackageRoot $effectivePackageRoot -DescriptionBase $descriptionBase -EmptySelectionMessage $emptySelectionMessage -EvidenceRoot $EvidenceRoot -NoEvidence:$NoEvidence
        return
    }
    Publish-AutomaticWaveProgress -Stage 'Computing the selection again...' -Detail ('The preview cannot be reused: {0}.' -f $previewReuse.Reason)
}

Publish-AutomaticWaveProgress -Stage 'Loading inventory records...' -Detail $(if ($entraAvailable) { 'AD + Intune + Entra' } else { 'AD + Intune; Entra inventory unavailable' })
$adRows = @(Import-Csv -LiteralPath $AdInventoryCsv)
$intuneRows = @(Import-Csv -LiteralPath $IntuneInventoryCsv)
$entraRows = @()
if ($entraAvailable) { $entraRows = @(Import-Csv -LiteralPath $EntraInventoryCsv) }
$adRecords = @(Get-AdAutomaticWaveRecord -Rows $adRows)
$intuneRecords = @(Get-IntuneAutomaticWaveRecord -Rows $intuneRows)
$entraRecords = @(Get-EntraAutomaticWaveRecord -Rows $entraRows)

$adByKey = @{}
foreach ($record in $adRecords) {
    if (-not $adByKey.ContainsKey($record.Key)) { $adByKey[$record.Key] = [System.Collections.Generic.List[object]]::new() }
    $adByKey[$record.Key].Add($record)
}
$intuneByKey = @{}
foreach ($record in $intuneRecords) {
    if (-not $intuneByKey.ContainsKey($record.Key)) { $intuneByKey[$record.Key] = [System.Collections.Generic.List[object]]::new() }
    $intuneByKey[$record.Key].Add($record)
}
$entraByKey = @{}
foreach ($record in $entraRecords) {
    if (-not $entraByKey.ContainsKey($record.Key)) { $entraByKey[$record.Key] = [System.Collections.Generic.List[object]]::new() }
    $entraByKey[$record.Key].Add($record)
}

Assert-AutomaticWaveAdDomainsInInventory -Domains $adDomains -AdRecords $adRecords -InventoryPath $AdInventoryCsv
$adLastLogonCutoffUtc = [datetime]::UtcNow.AddDays(-$AdLastLogonMaxAgeDays)
$selected = [System.Collections.Generic.List[object]]::new()
$excluded = [System.Collections.Generic.List[object]]::new()
$filterExcluded = [System.Collections.Generic.List[object]]::new()

# Per-computer checks use plain loops: Where-Object/Sort-Object pipelines per computer took minutes
# with tens of thousands of AD computers.
foreach ($key in @($adByKey.Keys | Sort-Object)) {
    $keyAdRecords = $adByKey[$key].ToArray()
    $keyIntuneRecords = @(if ($intuneByKey.ContainsKey($key)) { $intuneByKey[$key].ToArray() })
    $keyEntraRecords = @(if ($entraByKey.ContainsKey($key)) { $entraByKey[$key].ToArray() })
    $preferredAd = if ($keyAdRecords.Count -eq 1) { $keyAdRecords[0] } else { @($keyAdRecords | Sort-Object -Property DNSHostName | Select-Object -First 1)[0] }
    $filterReasons = [System.Collections.Generic.List[string]]::new()
    $filterTypes = [System.Collections.Generic.List[string]]::new()
    $filterValues = [System.Collections.Generic.List[string]]::new()

    if ($namePrefixes.Count -gt 0 -and -not (Test-AutomaticWaveNamePrefix -ComputerKey $key -Prefixes $namePrefixes)) {
        $filterTypes.Add('ComputerNamePrefix')
        $filterValues.Add($namePrefixDisplay)
        $filterReasons.Add('COMPUTER_NAME_PREFIX_NOT_MATCHED')
    }
    if ($nameContainsValues.Count -gt 0 -and -not (Test-AutomaticWaveNameContains -ComputerKey $key -ContainsValues $nameContainsValues)) {
        $filterTypes.Add('ComputerNameContains')
        $filterValues.Add($nameContainsDisplay)
        $filterReasons.Add('COMPUTER_NAME_CONTAINS_NOT_MATCHED')
    }
    $adDomainReason = if ($adDomains.Count -gt 0) { Get-AutomaticWaveAdDomainFilterReason -AdRecords $keyAdRecords -Domains $adDomains } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($adDomainReason)) {
        $filterTypes.Add('ADDomain')
        $filterValues.Add($adDomainDisplay)
        $filterReasons.Add($adDomainReason)
    }
    $latestLastSeen = [datetime]::MinValue
    foreach ($record in $keyAdRecords) { if ($record.LastSeenUtc -gt $latestLastSeen) { $latestLastSeen = [datetime]$record.LastSeenUtc } }
    if ($ExcludeStaleAd) {
        if ($latestLastSeen -eq [datetime]::MinValue) {
            $filterTypes.Add('ADLastLogon')
            $filterValues.Add("MaxAgeDays=$AdLastLogonMaxAgeDays")
            $filterReasons.Add('AD_LAST_LOGON_UNKNOWN')
        }
        elseif ($latestLastSeen -lt $adLastLogonCutoffUtc) {
            $filterTypes.Add('ADLastLogon')
            $filterValues.Add("MaxAgeDays=$AdLastLogonMaxAgeDays")
            $filterReasons.Add('AD_LAST_LOGON_OLDER_THAN_LIMIT')
        }
    }
    if ($filterReasons.Count -gt 0) {
        $filterExcluded.Add([pscustomobject]@{
            ComputerKey = $key
            ComputerName = [string]$preferredAd.PreferredName
            FilterType = $filterTypes.ToArray() -join ';'
            FilterValue = $filterValues.ToArray() -join ' | '
            FilterReason = $filterReasons.ToArray() -join ';'
            ADLastLogonTimestampUtc = if ($latestLastSeen -eq [datetime]::MinValue) { '' } else { $latestLastSeen.ToString('o') }
        })
        continue
    }

    $reasons = [System.Collections.Generic.List[string]]::new()
    $adObjectKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $adNotPresent = $false
    $adDisabled = $false
    foreach ($record in $keyAdRecords) {
        $objectKey = if (-not [string]::IsNullOrWhiteSpace([string]$record.DistinguishedName)) { [string]$record.DistinguishedName }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$record.DNSHostName)) { [string]$record.DNSHostName }
        else { [string]$record.Name }
        [void]$adObjectKeys.Add($objectKey.ToUpperInvariant())
        if (-not $record.Present) { $adNotPresent = $true }
        elseif (-not $record.Enabled) { $adDisabled = $true }
    }
    $intuneMatches = 0
    foreach ($record in $keyIntuneRecords) { if ($record.Present) { $intuneMatches++ } }
    $serverAdRecords = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $keyEntraRecords) { if ($record.Present -and $record.TrustType -ieq 'ServerAd') { $serverAdRecords.Add($record) } }

    if ($adObjectKeys.Count -gt 1) { $reasons.Add('AD_NAME_COLLISION') }
    if ($adNotPresent) { $reasons.Add('AD_INVENTORY_NOT_PRESENT') }
    if ($adDisabled) { $reasons.Add('AD_DISABLED') }
    if ($preferredAd.Classification -eq 'WindowsServer') { $reasons.Add('WINDOWS_SERVER_EXCLUDED') }
    elseif ($preferredAd.Classification -notin @('Windows10', 'Windows11')) { $reasons.Add('OS_UNKNOWN_OR_UNSUPPORTED') }
    # A name match is only a candidate; local identity must be checked before completion.

    if ($serverAdRecords.Count -gt 1) { $reasons.Add('ENTRA_SERVERAD_COLLISION') }
    if ($serverAdRecords.Count -eq 1 -and -not $serverAdRecords[0].AccountEnabled) { $reasons.Add('ENTRA_SERVERAD_DISABLED') }

    if ($reasons.Count -gt 0) {
        $fqdns = [System.Collections.Generic.List[string]]::new()
        foreach ($record in $keyAdRecords) { if (-not $fqdns.Contains([string]$record.DNSHostName)) { $fqdns.Add([string]$record.DNSHostName) } }
        $excluded.Add([pscustomobject]@{
            ComputerKey = $key
            ComputerName = [string]$preferredAd.PreferredName
            ExclusionReason = $reasons.ToArray() -join ';'
            ADFQDNs = $fqdns.ToArray() -join ';'
            ADOperatingSystem = [string]$preferredAd.OperatingSystem
            ADEnabled = [string]$preferredAd.Enabled
            IntuneMatches = $intuneMatches
            EntraServerAdMatches = $serverAdRecords.Count
        })
        continue
    }

    $selectionReason = 'NEEDS_HYBRID_JOIN'
    if (-not $entraAvailable) {
        $selectionReason = 'ENTRA_INVENTORY_UNAVAILABLE'
    }
    elseif ($serverAdRecords.Count -eq 1) {
        if ($serverAdRecords[0].RegisteredState -ieq 'Pending' -or [string]$serverAdRecords[0].AlternativeSecurityIdCount -eq '0') {
            $selectionReason = 'HYBRID_JOIN_PENDING'
        }
        else {
            $selectionReason = 'NEEDS_INTUNE_ENROLLMENT'
        }
    }

    $selected.Add([pscustomobject]@{
        ComputerKey = $key
        ComputerName = [string]$preferredAd.PreferredName
        SelectionReason = $selectionReason
        ADOperatingSystem = [string]$preferredAd.OperatingSystem
        ADOSVersion = [string]$preferredAd.OSVersion
        ADEnabled = [string]$preferredAd.Enabled
        ADDomain = [string]$preferredAd.Domain
        ADLastLogonTimestampUtc = if ($latestLastSeen -eq [datetime]::MinValue) { '' } else { $latestLastSeen.ToString('o') }
        EntraInventoryAvailable = $entraAvailable
        EntraRegisteredState = if ($serverAdRecords.Count -eq 1) { [string]$serverAdRecords[0].RegisteredState } else { '' }
        EntraTrustType = if ($serverAdRecords.Count -eq 1) { [string]$serverAdRecords[0].TrustType } else { '' }
        EntraDeviceId = if ($serverAdRecords.Count -eq 1) { [string]$serverAdRecords[0].DeviceId } else { '' }
    })
}

Publish-AutomaticWaveProgress -Stage 'Applying selection and safety filters...' -Detail ("Selected: {0}; safety exclusions: {1}; filtered out: {2}" -f $selected.Count, $excluded.Count, $filterExcluded.Count)
$wavePath = ''
$computersPath = ''
if ($Create) {
    Assert-AutomaticWaveCreatable -SelectedCount $selected.Count -WaveName $safeWaveName -PackageRoot $effectivePackageRoot -EmptySelectionMessage $emptySelectionMessage
}

$evidencePath = ''
if (-not $NoEvidence) {
    if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
        $EvidenceRoot = Join-Path (Get-SemtStorageRoot -PackageRoot $effectivePackageRoot -Kind Runs) ('{0}\_AutomaticWave' -f $script:AutomaticWaveCampaign)
    }
    $evidencePath = Join-Path $EvidenceRoot (Get-Date -Format 'yyyyMMdd-HHmmssfff')
    Publish-AutomaticWaveProgress -Stage 'Preparing automatic Wave evidence...' -Detail $evidencePath
    New-SemtDirectory -Path $evidencePath
    Copy-Item -LiteralPath $AdInventoryCsv -Destination (Join-Path $evidencePath 'DevicesAD.csv') -Force
    Copy-Item -LiteralPath $IntuneInventoryCsv -Destination (Join-Path $evidencePath 'DevicesIntune.csv') -Force
    if ($entraAvailable) { Copy-Item -LiteralPath $EntraInventoryCsv -Destination (Join-Path $evidencePath 'DevicesEntra.csv') -Force }
    Export-AutomaticWaveCsv -Path (Join-Path $evidencePath 'AutomaticWaveSelection.csv') -Columns @(
        'ComputerKey','ComputerName','SelectionReason','ADOperatingSystem','ADOSVersion','ADEnabled','ADDomain',
        'ADLastLogonTimestampUtc','EntraInventoryAvailable','EntraRegisteredState','EntraTrustType','EntraDeviceId'
    ) -Rows $selected.ToArray()
    Export-AutomaticWaveCsv -Path (Join-Path $evidencePath 'AutomaticWaveExclusions.csv') -Columns @(
        'ComputerKey','ComputerName','ExclusionReason','ADFQDNs','ADOperatingSystem','ADEnabled','IntuneMatches','EntraServerAdMatches'
    ) -Rows $excluded.ToArray()
    Export-AutomaticWaveCsv -Path (Join-Path $evidencePath 'AutomaticWaveFilterExclusions.csv') -Columns @(
        'ComputerKey','ComputerName','FilterType','FilterValue','FilterReason','ADLastLogonTimestampUtc'
    ) -Rows $filterExcluded.ToArray()
}

if ($Create) {
    $description = '{0}; Selected={1}' -f $descriptionBase, $selected.Count
    Publish-AutomaticWaveProgress -Stage 'Creating the automatic Wave...' -Detail ("Wave: {0}; selected devices: {1}" -f $safeWaveName, $selected.Count)
    $waveComputers = @($selected | Sort-Object ComputerName | Select-Object -ExpandProperty ComputerName)
    $wave = New-SemtWave -PackageRoot $effectivePackageRoot -Campaign $script:AutomaticWaveCampaign -Name $safeWaveName -Computers $waveComputers -Source Automatic -Description $description -EvidencePath $evidencePath
    $wavePath = [string]$wave.Path
    $computersPath = [string]$wave.ComputersPath
}

$summary = [pscustomobject]@{
    Version = $script:AutomaticWaveVersion
    Campaign = $script:AutomaticWaveCampaign
    GeneratedUtc = [datetime]::UtcNow.ToString('o')
    SelectionSignature = $selectionSignature
    InventoryFingerprint = $inventoryFingerprint
    AvailableSources = $availableSourcesText
    EntraInventoryAvailable = $entraAvailable
    ComputerNamePrefixes = $namePrefixDisplay
    ComputerNameContains = $nameContainsDisplay
    ADDomainFilter = $adDomainDisplay
    ExcludeIntunePresent = $false
    ExcludeStaleAd = [bool]$ExcludeStaleAd
    ADLastLogonMaxAgeDays = $AdLastLogonMaxAgeDays
    ADRows = $adRows.Count
    IntuneRows = $intuneRows.Count
    EntraRows = $entraRows.Count
    UniqueADDevices = $adByKey.Count
    SelectedDevices = $selected.Count
    ExcludedDevices = $excluded.Count
    FilterExcludedDevices = $filterExcluded.Count
    NeedsHybridJoin = $selected.Where({ $_.SelectionReason -eq 'NEEDS_HYBRID_JOIN' }).Count
    HybridJoinPending = $selected.Where({ $_.SelectionReason -eq 'HYBRID_JOIN_PENDING' }).Count
    NeedsIntuneEnrollment = $selected.Where({ $_.SelectionReason -eq 'NEEDS_INTUNE_ENROLLMENT' }).Count
    EntraUnavailableSelected = $selected.Where({ $_.SelectionReason -eq 'ENTRA_INVENTORY_UNAVAILABLE' }).Count
    IntunePresentExcluded = $excluded.Where({ $_.ExclusionReason -match 'INTUNE_DEVICE_PRESENT' }).Count
    ADDisabledExcluded = $excluded.Where({ $_.ExclusionReason -match 'AD_DISABLED' }).Count
    WindowsServerExcluded = $excluded.Where({ $_.ExclusionReason -match 'WINDOWS_SERVER_EXCLUDED' }).Count
    UnknownOSExcluded = $excluded.Where({ $_.ExclusionReason -match 'OS_UNKNOWN_OR_UNSUPPORTED' }).Count
    ADNameCollisions = $excluded.Where({ $_.ExclusionReason -match 'AD_NAME_COLLISION' }).Count
    EntraAmbiguousExcluded = $excluded.Where({ $_.ExclusionReason -match 'ENTRA_SERVERAD_' }).Count
    ADLastLogonFilterExcluded = $filterExcluded.Where({ $_.FilterReason -match 'AD_LAST_LOGON_' }).Count
    ADDomainFilterExcluded = $filterExcluded.Where({ $_.FilterReason -match 'AD_DOMAIN_' }).Count
    WaveCreated = [bool]$Create
    WaveName = $safeWaveName
    WavePath = $wavePath
    ComputersPath = $computersPath
    EvidencePath = $evidencePath
}

if (-not $NoEvidence) {
    Write-SemtJsonFileAtomic -Path (Join-Path $evidencePath 'AutomaticWaveSummary.json') -InputObject $summary -Depth 5
}
Publish-AutomaticWaveProgress -Stage $(if ($Create) { 'Automatic Wave created.' } else { 'Automatic Wave preview ready.' }) -Detail $(if ($Create) { $wavePath } else { "Selected devices: $($selected.Count)" })

[pscustomobject]@{
    Summary = $summary
    SelectedDevices = $selected.ToArray()
    ExcludedDevices = $excluded.ToArray()
    FilterExcludedDevices = $filterExcluded.ToArray()
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDWFXrQnPA8/LRG
# pq2YedGEcsPkx7rBwQNE/iNBI5NhQKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB6Qn2t815BSOU9PgtkKDHL
# gREhpl3xCSmPvy9/mx5bYzANBgkqhkiG9w0BAQEFAASCAYCTmvIGAfgVrrFV4HxE
# ZwwK1f3JkwfRLXcRYc2sN7wVFPjUtRgmm/jt0B9/lJujWuf1PIq5uGCEspkqt+7I
# bKOrNXx3e65/rFHZlEUJR0P8PnTIm2+rp4E689lBBMNYd5+I6BFZRo8/fg+kXvi7
# Eq4dvHf1v2RDTrjRcClhw7n1qVvDxffXHFNZhYDN6IWKkhKTTJEU43RK71ZLzTA0
# eP4gwyXxh93EjQz0bdlnSZi9vXIYwrQZLmdws1nptqZY04AQ9XCCMR7114L+r0yO
# 1vfmL5vlxau2wuQp19Cv1VdLu3SUzn3Idmivm/bNrAe3xvVtqzVb423OyqFHyAd3
# /rwolEmrsMcRYDjQb4ZlrmS594O0xn2s4j6+470rrf2UvbejftRLFvZWLSP+kU1p
# DILCTRgMVetaEvsbQqyE1FxeVAIk+DBzBxff3hU2Wv/qh9fopc66D7vaj1iGvcqC
# Ppsu8Pb/TXKJD2gKM9D44gIM0vhTtUD0GAGVCeys6RNo+xk=
# SIG # End signature block
