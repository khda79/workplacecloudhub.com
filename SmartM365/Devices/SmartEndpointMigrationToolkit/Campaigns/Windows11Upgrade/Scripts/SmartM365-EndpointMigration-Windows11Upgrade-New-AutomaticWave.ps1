<#
.SYNOPSIS
Builds a Windows11Upgrade Wave of Windows 10 devices from AD and/or Intune inventory CSV files.

.DESCRIPTION
Selects explicit Windows 10 devices from broad AD and Intune inventory CSV files,
deduplicates them, excludes Windows 11 evidence and unsafe inventory states, writes
selection evidence under Runs\Windows11Upgrade\_AutomaticWave, and can create a
standard Windows11Upgrade Wave through New-SemtWave without launching it. Default
inventory inputs are the shared caches in <Package>\Inventory.
Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/Scripts/SmartM365-Windows11Upgrade-New-AutomaticLot.ps1 1.3.2 (see Docs\PROVENANCE.json.txt).

.PARAMETER Source
Inventory sources required for the selection: AD, Intune or Both.

.PARAMETER AdInventoryCsv
AD inventory CSV. Defaults to the shared cache <Package>\Inventory\DevicesAD.csv.

.PARAMETER IntuneInventoryCsv
Intune inventory CSV. Defaults to the shared cache <Package>\Inventory\DevicesIntune.csv.

.PARAMETER PackageRoot
Smart Endpoint Migration Toolkit package root. Defaults to the package that contains this script.

.PARAMETER WaveName
Wave name. Defaults to WAVE-AUTO-W10-<prefix>-<yyyyMMdd-HHmmss>.

.PARAMETER AdDomainFilter
Optional AD domain DNS names (for example contoso.local). Only computers with an AD record in one of
these domains are kept; the others are listed in AutomaticWaveFilterExclusions.csv with
AD_DOMAIN_NOT_MATCHED, or AD_DOMAIN_UNKNOWN when they have no AD record (Intune only) or no domain. Requires the AD source.
Each domain must exist in the AD inventory. The shared AD inventory always covers the whole forest.

.PARAMETER Create
Creates the Wave (Computers.txt, Wave.local.json.txt, wrappers). Without it, only evidence is written.

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
    [ValidateSet('AD', 'Intune', 'Both')]
    [string]$Source = 'Both',
    [string]$AdInventoryCsv,
    [string]$IntuneInventoryCsv,
    [string]$PackageRoot,
    [string]$WaveName,
    [string[]]$ComputerNamePrefix,
    [string[]]$ComputerNameContains,
    [string[]]$AdDomainFilter,
    [switch]$ExcludeIntunePresent,
    [switch]$ExcludeStaleAd,
    [ValidateRange(1, 3650)]
    [int]$AdLastLogonMaxAgeDays = 45,
    [scriptblock]$ProgressCallback,
    [string]$EvidenceRoot,
    [switch]$Create,
    [string]$PreviewEvidencePath,
    [switch]$AllowPartialSource,
    [switch]$NoEvidence
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:AutomaticWaveVersion = '1.0.0'
$script:AutomaticWaveCampaign = 'Windows11Upgrade'
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
        $nameSegments = [System.Collections.Generic.List[string]]::new()
        foreach ($prefix in @($NamePrefixes)) {
            $segment = [regex]::Replace(([string]$prefix).ToUpperInvariant(), '[^A-Z0-9_-]+', '-').Trim('-_')
            if (-not [string]::IsNullOrWhiteSpace($segment) -and -not $nameSegments.Contains($segment)) {
                $nameSegments.Add($segment)
            }
        }

        $prefixSegment = if ($nameSegments.Count -gt 0) { "-$($nameSegments -join '-')" } else { '' }
        $Name = 'WAVE-AUTO-W10{0}-{1}' -f $prefixSegment, (Get-Date -Format 'yyyyMMdd-HHmmss')
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
                throw "Invalid computer name prefix '$value'. Use only letters, digits, hyphens, or underscores; wildcards and regular expressions are not supported."
            }
            if (-not $normalized.Contains($prefix)) { $normalized.Add($prefix) }
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

function Get-AutomaticWaveNameContainsValues {
    param([AllowNull()][string[]]$ContainsValues)

    $normalized = [System.Collections.Generic.List[string]]::new()
    foreach ($item in @($ContainsValues)) {
        foreach ($candidate in @(([string]$item) -split ';')) {
            $value = $candidate.Trim().Trim([char]34).ToUpperInvariant()
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            if ($value -notmatch '^[A-Z0-9_-]+$') {
                throw "Invalid computer name contains value '$candidate'. Use only letters, digits, hyphens, or underscores; wildcards and regular expressions are not supported."
            }
            if (-not $normalized.Contains($value)) { $normalized.Add($value) }
        }
    }

    return @($normalized.ToArray())
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

function Get-InventoryWindowsClassification {
    param(
        [AllowNull()][string]$OperatingSystem,
        [AllowNull()][string]$Version
    )

    $osText = ([string]$OperatingSystem).Trim()
    $versionText = ([string]$Version).Trim()
    if ($osText -match '(?i)Server') { return 'Unknown' }
    if (-not [string]::IsNullOrWhiteSpace($osText) -and $osText -notmatch '(?i)Windows') { return 'Unknown' }
    if ($osText -match '(?i)Windows\s*11') { return 'Windows11' }
    if ($osText -match '(?i)Windows\s*10') { return 'Windows10' }

    $numbers = @([regex]::Matches($versionText, '\d+') | ForEach-Object { [int64]$_.Value })
    if ($numbers.Count -ge 3 -and $numbers[0] -eq 10 -and $numbers[1] -eq 0) {
        $build = $numbers[2]
        if ($build -ge 22000) { return 'Windows11' }
        if ($build -ge 10240) { return 'Windows10' }
    }

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
    $classificationCache = @{}
    $staleCutoffUtc = [datetime]::UtcNow.AddDays(-90)
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
        $classificationKey = $operatingSystem + [char]0 + $operatingSystemVersion
        if (-not $classificationCache.ContainsKey($classificationKey)) { $classificationCache[$classificationKey] = Get-InventoryWindowsClassification -OperatingSystem $operatingSystem -Version $operatingSystemVersion }
        $classification = $classificationCache[$classificationKey]
        $lastSeen = [datetime]::MinValue
        if ($colLastSeen) {
            $text = [string]$row.$colLastSeen
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $parsed = [datetime]::MinValue
                if ([datetime]::TryParse($text, $invariant, $dateStyles, [ref]$parsed) -or [datetime]::TryParse($text, [ref]$parsed)) { $lastSeen = $parsed.ToUniversalTime() }
            }
        }

        [pscustomobject]@{
            Key             = $key
            Name            = $computerName
            DNSHostName     = $dnsHostName
            PreferredName   = $identity
            Domain          = if ($colDomain) { [string]$row.$colDomain } else { '' }
            Present         = $present
            Enabled         = $enabled
            Classification  = $classification
            OperatingSystem = $operatingSystem
            OSVersion       = $operatingSystemVersion
            LastSeenUtc     = $lastSeen
            IsStale         = ($lastSeen -ne [datetime]::MinValue -and $lastSeen -lt $staleCutoffUtc)
            Eligible        = ($present -and $enabled -and $classification -eq 'Windows10')
        }
    }
}

function Get-IntuneAutomaticWaveRecord {
    param([object[]]$Rows)

    $rows = @($Rows)
    $colDeviceName = Resolve-InventoryColumn -Rows $rows -Names @('DeviceName', 'ManagedDeviceName', 'ComputerName')
    $colComputerName = Resolve-InventoryColumn -Rows $rows -Names @('ComputerName')
    $colOs = Resolve-InventoryColumn -Rows $rows -Names @('OperatingSystem')
    $colOsVersion = Resolve-InventoryColumn -Rows $rows -Names @('OSVersion', 'OperatingSystemVersion')
    $colPresent = Resolve-InventoryColumn -Rows $rows -Names @('IntuneInventoryPresent')
    $colState = Resolve-InventoryColumn -Rows $rows -Names @('ManagementState')
    $colLastSync = Resolve-InventoryColumn -Rows $rows -Names @('LastSyncDateTime')
    $colId = Resolve-InventoryColumn -Rows $rows -Names @('IntuneManagedDeviceId', 'Id')
    $classificationCache = @{}
    $staleCutoffUtc = [datetime]::UtcNow.AddDays(-30)
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $dateStyles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces -bor [System.Globalization.DateTimeStyles]::AssumeUniversal

    foreach ($row in $rows) {
        if ($null -eq $row) { continue }
        $deviceName = if ($colDeviceName) { [string]$row.$colDeviceName } else { '' }
        $computerName = if ($colComputerName) { [string]$row.$colComputerName } else { '' }
        $identity = if (-not [string]::IsNullOrWhiteSpace($deviceName)) { $deviceName } else { $computerName }
        if ([string]::IsNullOrWhiteSpace($identity)) { continue }
        $key = $identity.Trim().Trim([char]34).Split('.')[0].ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($key)) { continue }

        $operatingSystem = if ($colOs) { [string]$row.$colOs } else { '' }
        $operatingSystemVersion = if ($colOsVersion) { [string]$row.$colOsVersion } else { '' }
        $present = $true
        if ($colPresent) { $text = [string]$row.$colPresent; if (-not [string]::IsNullOrWhiteSpace($text) -and $text.Trim() -match '^(?i:false|0|no|disabled)$') { $present = $false } }
        $managementState = if ($colState) { [string]$row.$colState } else { '' }
        $stateExcluded = (-not [string]::IsNullOrWhiteSpace($managementState) -and $managementState -match '(?i)(retire|wipe|delete)')
        $classificationKey = $operatingSystem + [char]0 + $operatingSystemVersion
        if (-not $classificationCache.ContainsKey($classificationKey)) { $classificationCache[$classificationKey] = Get-InventoryWindowsClassification -OperatingSystem $operatingSystem -Version $operatingSystemVersion }
        $classification = $classificationCache[$classificationKey]
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
            DeviceName      = $deviceName
            ComputerName    = $computerName
            PreferredName   = $identity
            ManagedDeviceId = if ($colId) { [string]$row.$colId } else { '' }
            Present         = $present
            ManagementState = $managementState
            StateExcluded   = $stateExcluded
            Classification  = $classification
            OperatingSystem = $operatingSystem
            OSVersion       = $operatingSystemVersion
            LastSyncUtc     = $lastSync
            IsStale         = ($lastSync -ne [datetime]::MinValue -and $lastSync -lt $staleCutoffUtc)
            Eligible        = ($present -and -not $stateExcluded -and $classification -eq 'Windows10')
        }
    }
}

function Get-AutomaticWaveExclusionReason {
    param(
        [object[]]$AdRecords,
        [object[]]$IntuneRecords,
        [bool]$AdCollision
    )

    $windows11 = $false; $adDisabled = $false; $stateExcluded = $false; $notPresent = $false; $unknownOs = $false
    foreach ($record in @($AdRecords)) {
        if ($null -eq $record) { continue }
        if ($record.Classification -eq 'Windows11') { $windows11 = $true }
        if ($record.Classification -eq 'Unknown') { $unknownOs = $true }
        if (-not $record.Present) { $notPresent = $true }
        elseif (-not $record.Enabled) { $adDisabled = $true }
    }
    foreach ($record in @($IntuneRecords)) {
        if ($null -eq $record) { continue }
        if ($record.Classification -eq 'Windows11') { $windows11 = $true }
        if ($record.Classification -eq 'Unknown') { $unknownOs = $true }
        if ($record.StateExcluded) { $stateExcluded = $true }
        if (-not $record.Present) { $notPresent = $true }
    }
    $reasons = [System.Collections.Generic.List[string]]::new()
    if ($AdCollision) { $reasons.Add('AD_NAME_COLLISION') }
    if ($windows11) { $reasons.Add('WINDOWS11_REPORTED') }
    if ($adDisabled) { $reasons.Add('AD_DISABLED') }
    if ($stateExcluded) { $reasons.Add('INTUNE_MANAGEMENT_STATE_EXCLUDED') }
    if ($notPresent) { $reasons.Add('INVENTORY_NOT_PRESENT') }
    if ($unknownOs) { $reasons.Add('OS_UNKNOWN_OR_UNSUPPORTED') }
    if ($reasons.Count -eq 0) { $reasons.Add('NOT_ELIGIBLE_WINDOWS10') }
    return $reasons.ToArray() -join ';'
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

    $summaryPath = Join-Path $Path 'AutomaticWaveSummary.json.txt'
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
        Write-SemtJsonFileAtomic -Path (Join-Path $evidencePath 'AutomaticWaveSummary.json.txt') -InputObject $summary -Depth 5
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
$requiredSources = switch ($Source) {
    'AD' { @('AD') }
    'Intune' { @('Intune') }
    default { @('AD', 'Intune') }
}

$availableSources = [System.Collections.Generic.List[string]]::new()
$missingSources = [System.Collections.Generic.List[string]]::new()
if ('AD' -in $requiredSources) {
    if (-not [string]::IsNullOrWhiteSpace($AdInventoryCsv) -and (Test-Path -LiteralPath $AdInventoryCsv -PathType Leaf)) {
        $availableSources.Add('AD')
    }
    else {
        $missingSources.Add('AD')
    }
}
if ('Intune' -in $requiredSources) {
    if (-not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv) -and (Test-Path -LiteralPath $IntuneInventoryCsv -PathType Leaf)) {
        $availableSources.Add('Intune')
    }
    else {
        $missingSources.Add('Intune')
    }
}

if ($missingSources.Count -gt 0 -and (-not $AllowPartialSource -or $availableSources.Count -eq 0)) {
    throw ('Required automatic Wave inventory source(s) missing: {0}.' -f ($missingSources.ToArray() -join ', '))
}
if ($ExcludeIntunePresent -and 'Intune' -notin $availableSources) {
    throw 'ExcludeIntunePresent requires an available Intune inventory source.'
}
if (@(Get-AutomaticWaveAdDomains -Domains $AdDomainFilter).Count -gt 0 -and 'AD' -notin $availableSources) {
    throw 'AdDomainFilter requires an available AD inventory source.'
}
if ($ExcludeStaleAd -and 'AD' -notin $availableSources) {
    throw 'ExcludeStaleAd requires an available AD inventory source.'
}

$namePrefixes = @(Get-AutomaticWaveNamePrefixes -Prefixes $ComputerNamePrefix)
$namePrefixDisplay = $namePrefixes -join ';'
$nameContainsValues = @(Get-AutomaticWaveNameContainsValues -ContainsValues $ComputerNameContains)
$nameContainsDisplay = $nameContainsValues -join ';'
$adDomains = @(Get-AutomaticWaveAdDomains -Domains $AdDomainFilter)
$adDomainDisplay = $adDomains -join ';'
$safeWaveName = Get-AutomaticWaveSafeName -Name $WaveName -NamePrefixes $namePrefixes
$emptySelectionMessage = 'No eligible Windows 10 device was selected; the Wave was not created.'
$descriptionBase = 'Automatic Windows11Upgrade Wave; Source={0}; AvailableSources={1}; ComputerNamePrefix={2}; ComputerNameContains={3}; ADDomain={7}; ExcludeIntunePresent={4}; ExcludeStaleAd={5}; ADLastLogonMaxAgeDays={6}' -f $Source,($availableSources.ToArray() -join '+'),$namePrefixDisplay,$nameContainsDisplay,[bool]$ExcludeIntunePresent,[bool]$ExcludeStaleAd,$AdLastLogonMaxAgeDays,$adDomainDisplay
$selectionSignature = @($script:AutomaticWaveCampaign, $script:AutomaticWaveVersion, "Source=$Source", "AvailableSources=$($availableSources.ToArray() -join '+')", "ComputerNamePrefix=$namePrefixDisplay", "ComputerNameContains=$nameContainsDisplay", "ADDomain=$adDomainDisplay", "ExcludeIntunePresent=$([bool]$ExcludeIntunePresent)", "ExcludeStaleAd=$([bool]$ExcludeStaleAd)", "ADLastLogonMaxAgeDays=$AdLastLogonMaxAgeDays") -join '|'
$inventoryFingerprint = Get-AutomaticWaveInventoryFingerprint -Paths @($(if ('AD' -in $availableSources) { $AdInventoryCsv } else { '' }), $(if ('Intune' -in $availableSources) { $IntuneInventoryCsv } else { '' }))

if ($Create -and -not [string]::IsNullOrWhiteSpace($PreviewEvidencePath)) {
    $previewReuse = Get-AutomaticWavePreviewReuse -Path $PreviewEvidencePath -SelectionSignature $selectionSignature -InventoryFingerprint $inventoryFingerprint
    if ($previewReuse.Summary) {
        New-AutomaticWaveFromPreview -PreviewPath $PreviewEvidencePath -PreviewSummary $previewReuse.Summary -WaveName $safeWaveName -PackageRoot $effectivePackageRoot -DescriptionBase $descriptionBase -EmptySelectionMessage $emptySelectionMessage -EvidenceRoot $EvidenceRoot -NoEvidence:$NoEvidence
        return
    }
    Publish-AutomaticWaveProgress -Stage 'Computing the selection again...' -Detail ('The preview cannot be reused: {0}.' -f $previewReuse.Reason)
}

Publish-AutomaticWaveProgress -Stage 'Loading inventory records...' -Detail ("Available sources: {0}" -f ($availableSources.ToArray() -join ' + '))
$adRows = if ('AD' -in $availableSources) { @(Import-Csv -LiteralPath $AdInventoryCsv) } else { @() }
$intuneRows = if ('Intune' -in $availableSources) { @(Import-Csv -LiteralPath $IntuneInventoryCsv) } else { @() }
$adRows = @($adRows)
$intuneRows = @($intuneRows)
$adRecords = @(Get-AdAutomaticWaveRecord -Rows $adRows)
$intuneRecords = @(Get-IntuneAutomaticWaveRecord -Rows $intuneRows)

$adByKey = @{}
foreach ($record in $adRecords) {
    if (-not $adByKey.ContainsKey($record.Key)) {
        $adByKey[$record.Key] = [System.Collections.Generic.List[object]]::new()
    }
    $adByKey[$record.Key].Add($record)
}

$intuneByKey = @{}
$intuneAllByKey = @{}
$intuneDuplicateCount = 0
# Hashtable grouping instead of Group-Object (slow with tens of thousands of rows); a key with
# several Intune rows keeps the newest sync first, as before.
$intuneGroups = @{}
foreach ($record in $intuneRecords) {
    if (-not $intuneGroups.ContainsKey($record.Key)) { $intuneGroups[$record.Key] = [System.Collections.Generic.List[object]]::new() }
    $intuneGroups[$record.Key].Add($record)
}
foreach ($groupKey in @($intuneGroups.Keys)) {
    $group = $intuneGroups[$groupKey]
    $ordered = @(if ($group.Count -eq 1) { $group.ToArray() } else { $group.ToArray() | Sort-Object -Property @{ Expression = 'LastSyncUtc'; Descending = $true } })
    if ($ordered.Count -gt 1) { $intuneDuplicateCount += ($ordered.Count - 1) }
    $intuneAllByKey[$groupKey] = $ordered
    $intuneByKey[$groupKey] = $ordered[0]
}

$allKeySet = @{}
foreach ($keyName in $adByKey.Keys) { $allKeySet[$keyName] = $true }
foreach ($keyName in $intuneByKey.Keys) { $allKeySet[$keyName] = $true }
$allKeys = @($allKeySet.Keys | Sort-Object)
Assert-AutomaticWaveAdDomainsInInventory -Domains $adDomains -AdRecords $adRecords -InventoryPath $AdInventoryCsv
$adLastLogonCutoffUtc = [datetime]::UtcNow.AddDays(-$AdLastLogonMaxAgeDays)
$selected = [System.Collections.Generic.List[object]]::new()
$excluded = [System.Collections.Generic.List[object]]::new()
$filterExcluded = [System.Collections.Generic.List[object]]::new()
$matchedKeyLookup = @{}

# Per-computer checks use plain loops: Where-Object/Sort-Object pipelines per computer took minutes
# with tens of thousands of computers.
foreach ($key in $allKeys) {
    $keyAdRecords = @()
    if ($adByKey.ContainsKey($key)) {
        $keyAdRecords = $adByKey[$key].ToArray()
    }
    $keyIntuneRecords = @(if ($intuneAllByKey.ContainsKey($key)) { $intuneAllByKey[$key] })
    $intuneRecord = if ($intuneByKey.ContainsKey($key)) { $intuneByKey[$key] } else { $null }

    $filterName = if ($keyAdRecords.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$keyAdRecords[0].PreferredName)) {
        [string]$keyAdRecords[0].PreferredName
    }
    elseif ($null -ne $intuneRecord) {
        [string]$intuneRecord.PreferredName
    }
    else {
        $key
    }
    $filterTypes = [System.Collections.Generic.List[string]]::new()
    $filterValues = [System.Collections.Generic.List[string]]::new()
    $filterReasons = [System.Collections.Generic.List[string]]::new()
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

    $intunePresent = $false
    $intuneWindows11 = $false
    $intuneStateExcluded = $false
    $intuneNotPresent = $false
    foreach ($record in $keyIntuneRecords) {
        if ($record.Present) { $intunePresent = $true } else { $intuneNotPresent = $true }
        if ($record.Classification -eq 'Windows11') { $intuneWindows11 = $true }
        if ($record.StateExcluded) { $intuneStateExcluded = $true }
    }
    if ($ExcludeIntunePresent -and $intunePresent) {
        $filterTypes.Add('IntunePresence')
        $filterValues.Add('Present=True')
        $filterReasons.Add('INTUNE_DEVICE_PRESENT')
    }

    $latestAdLastSeen = [datetime]::MinValue
    foreach ($record in $keyAdRecords) { if ($record.LastSeenUtc -gt $latestAdLastSeen) { $latestAdLastSeen = [datetime]$record.LastSeenUtc } }
    if ($ExcludeStaleAd) {
        if ($latestAdLastSeen -eq [datetime]::MinValue) {
            $filterTypes.Add('ADLastLogon')
            $filterValues.Add("MaxAgeDays=$AdLastLogonMaxAgeDays")
            $filterReasons.Add('AD_LAST_LOGON_UNKNOWN')
        }
        elseif ($latestAdLastSeen -lt $adLastLogonCutoffUtc) {
            $filterTypes.Add('ADLastLogon')
            $filterValues.Add("MaxAgeDays=$AdLastLogonMaxAgeDays")
            $filterReasons.Add('AD_LAST_LOGON_OLDER_THAN_LIMIT')
        }
    }

    if ($filterReasons.Count -gt 0) {
        $filterExcluded.Add([pscustomobject]@{
            ComputerKey                = $key
            ComputerName               = $filterName
            FilterType                 = $filterTypes.ToArray() -join ';'
            FilterValue                = $filterValues.ToArray() -join ' | '
            FilterReason               = $filterReasons.ToArray() -join ';'
            ADLastLogonTimestampUtc     = if ($latestAdLastSeen -eq [datetime]::MinValue) { '' } else { $latestAdLastSeen.ToString('o') }
            IntunePresent               = $intunePresent
        })
        continue
    }
    $matchedKeyLookup[$key] = $true

    $fqdnSet = @{}
    $adWindows11 = $false
    $adDisabledEvidence = $false
    $adNotPresent = $false
    $adStale = $false
    $eligibleAdRecords = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $keyAdRecords) {
        $fqdn = ([string]$record.DNSHostName).Trim().ToLowerInvariant()
        if ($fqdn -match '\.') { $fqdnSet[$fqdn] = $true }
        if ($record.Classification -eq 'Windows11') { $adWindows11 = $true }
        if (-not $record.Present) { $adNotPresent = $true }
        elseif (-not $record.Enabled) { $adDisabledEvidence = $true }
        if ($record.IsStale) { $adStale = $true }
        if ($record.Eligible) { $eligibleAdRecords.Add($record) }
    }
    $adFqdns = @($fqdnSet.Keys | Sort-Object)
    $adCollision = ($adFqdns.Count -gt 1)
    $windows11Evidence = ($adWindows11 -or $intuneWindows11)
    $inventoryNotPresent = ($adNotPresent -or $intuneNotPresent)
    $eligibleAd = @(if ($eligibleAdRecords.Count -le 1) { $eligibleAdRecords.ToArray() } else { $eligibleAdRecords.ToArray() | Sort-Object -Property DNSHostName | Select-Object -First 1 })
    $adEligible = ($eligibleAd.Count -gt 0)
    $intuneEligible = ($null -ne $intuneRecord -and $intuneRecord.Eligible)
    $unsafeInventoryState = ($adDisabledEvidence -or $intuneStateExcluded -or $inventoryNotPresent)

    if (-not $adCollision -and -not $windows11Evidence -and -not $unsafeInventoryState -and ($adEligible -or $intuneEligible)) {
        $preferredName = if ($adEligible -and -not [string]::IsNullOrWhiteSpace([string]$eligibleAd[0].PreferredName)) {
            [string]$eligibleAd[0].PreferredName
        }
        else {
            [string]$intuneRecord.PreferredName
        }
        $sourceEvidence = if ($adEligible -and $intuneEligible) { 'AD+Intune' } elseif ($adEligible) { 'AD' } else { 'Intune' }
        $selected.Add([pscustomobject]@{
            ComputerKey        = $key
            ComputerName       = $preferredName
            SourceEvidence     = $sourceEvidence
            ADOperatingSystem  = if ($keyAdRecords.Count -gt 0) { [string]$keyAdRecords[0].OperatingSystem } else { '' }
            ADOSVersion        = if ($keyAdRecords.Count -gt 0) { [string]$keyAdRecords[0].OSVersion } else { '' }
            ADEnabled          = if ($keyAdRecords.Count -gt 0) { [string]$keyAdRecords[0].Enabled } else { '' }
            ADLastLogonStale   = $adStale
            IntuneOperatingSystem = if ($null -ne $intuneRecord) { [string]$intuneRecord.OperatingSystem } else { '' }
            IntuneOSVersion    = if ($null -ne $intuneRecord) { [string]$intuneRecord.OSVersion } else { '' }
            IntuneManagementState = if ($null -ne $intuneRecord) { [string]$intuneRecord.ManagementState } else { '' }
            IntuneLastSyncStale = ($null -ne $intuneRecord -and $intuneRecord.IsStale)
            SelectionReason    = 'EXPLICIT_WINDOWS10'
        })
        continue
    }

    $fallbackName = if ($keyAdRecords.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$keyAdRecords[0].PreferredName)) {
        [string]$keyAdRecords[0].PreferredName
    }
    elseif ($null -ne $intuneRecord) {
        [string]$intuneRecord.PreferredName
    }
    else {
        $key
    }
    $excluded.Add([pscustomobject]@{
        ComputerKey        = $key
        ComputerName       = $fallbackName
        ExclusionReason    = Get-AutomaticWaveExclusionReason -AdRecords $keyAdRecords -IntuneRecords $keyIntuneRecords -AdCollision $adCollision
        ADFQDNs             = $adFqdns -join ';'
        ADOperatingSystem   = if ($keyAdRecords.Count -gt 0) { [string]$keyAdRecords[0].OperatingSystem } else { '' }
        IntuneOperatingSystem = if ($null -ne $intuneRecord) { [string]$intuneRecord.OperatingSystem } else { '' }
        IntuneOSVersion     = if ($null -ne $intuneRecord) { [string]$intuneRecord.OSVersion } else { '' }
        IntuneManagementState = if ($null -ne $intuneRecord) { [string]$intuneRecord.ManagementState } else { '' }
    })
}

Publish-AutomaticWaveProgress -Stage 'Applying selection and safety filters...' -Detail ("Selected: {0}; safety exclusions: {1}; filtered out: {2}" -f $selected.Count,$excluded.Count,$filterExcluded.Count)
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
    Publish-AutomaticWaveProgress -Stage 'Preparing the evidence folder...' -Detail $EvidenceRoot
    $evidencePath = Join-Path $EvidenceRoot (Get-Date -Format 'yyyyMMdd-HHmmssfff')
    New-SemtDirectory -Path $evidencePath

    Publish-AutomaticWaveProgress -Stage 'Copying inventory evidence...' -Detail ("Sources: {0}" -f ($availableSources.ToArray() -join ' + '))
    if ('AD' -in $availableSources) {
        Copy-Item -LiteralPath $AdInventoryCsv -Destination (Join-Path $evidencePath 'DevicesAD.csv') -Force
    }
    if ('Intune' -in $availableSources) {
        Copy-Item -LiteralPath $IntuneInventoryCsv -Destination (Join-Path $evidencePath 'DevicesIntune.csv') -Force
    }

    $selectionColumns = @(
        'ComputerKey', 'ComputerName', 'SourceEvidence', 'ADOperatingSystem', 'ADOSVersion',
        'ADEnabled', 'ADLastLogonStale', 'IntuneOperatingSystem', 'IntuneOSVersion',
        'IntuneManagementState', 'IntuneLastSyncStale', 'SelectionReason'
    )
    $exclusionColumns = @(
        'ComputerKey', 'ComputerName', 'ExclusionReason', 'ADFQDNs', 'ADOperatingSystem',
        'IntuneOperatingSystem', 'IntuneOSVersion', 'IntuneManagementState'
    )
    $filterExclusionColumns = @(
        'ComputerKey', 'ComputerName', 'FilterType', 'FilterValue', 'FilterReason',
        'ADLastLogonTimestampUtc', 'IntunePresent'
    )
    Publish-AutomaticWaveProgress -Stage 'Writing selection reports...' -Detail $evidencePath
    Export-AutomaticWaveCsv -Path (Join-Path $evidencePath 'AutomaticWaveSelection.csv') -Columns $selectionColumns -Rows $selected.ToArray()
    Export-AutomaticWaveCsv -Path (Join-Path $evidencePath 'AutomaticWaveExclusions.csv') -Columns $exclusionColumns -Rows $excluded.ToArray()
    Export-AutomaticWaveCsv -Path (Join-Path $evidencePath 'AutomaticWaveFilterExclusions.csv') -Columns $filterExclusionColumns -Rows $filterExcluded.ToArray()
}

if ($Create) {
    $description = '{0}; Selected={1}' -f $descriptionBase,$selected.Count
    Publish-AutomaticWaveProgress -Stage 'Creating the automatic Wave...' -Detail ("Wave: {0}; selected devices: {1}" -f $safeWaveName,$selected.Count)
    $waveComputers = @($selected | Sort-Object ComputerName | Select-Object -ExpandProperty ComputerName)
    $wave = New-SemtWave -PackageRoot $effectivePackageRoot -Campaign $script:AutomaticWaveCampaign -Name $safeWaveName -Computers $waveComputers -Source Automatic -Description $description -EvidencePath $evidencePath
    $wavePath = [string]$wave.Path
    $computersPath = [string]$wave.ComputersPath
}

$summary = [pscustomobject]@{
    Version                    = $script:AutomaticWaveVersion
    Campaign                   = $script:AutomaticWaveCampaign
    GeneratedUtc               = [datetime]::UtcNow.ToString('o')
    SelectionSignature         = $selectionSignature
    InventoryFingerprint       = $inventoryFingerprint
    RequestedSource            = $Source
    AvailableSources           = $availableSources.ToArray() -join '+'
    MissingSources             = $missingSources.ToArray() -join ','
    PartialSource              = ($missingSources.Count -gt 0)
    ComputerNamePrefixes       = $namePrefixDisplay
    ComputerNameContains       = $nameContainsDisplay
    ADDomainFilter             = $adDomainDisplay
    ExcludeIntunePresent       = [bool]$ExcludeIntunePresent
    ExcludeStaleAd             = [bool]$ExcludeStaleAd
    ADLastLogonMaxAgeDays      = $AdLastLogonMaxAgeDays
    NameFilterEnabled          = ($namePrefixes.Count -gt 0 -or $nameContainsValues.Count -gt 0 -or $adDomains.Count -gt 0 -or $ExcludeIntunePresent -or $ExcludeStaleAd)
    UniqueInventoryDevices     = $allKeys.Count
    NameFilterMatchedDevices   = $matchedKeyLookup.Count
    NameFilterExcludedDevices  = $filterExcluded.Count
    PrefixFilterExcluded       = $filterExcluded.Where({ $_.FilterReason -match 'COMPUTER_NAME_PREFIX_NOT_MATCHED' }).Count
    ContainsFilterExcluded     = $filterExcluded.Where({ $_.FilterReason -match 'COMPUTER_NAME_CONTAINS_NOT_MATCHED' }).Count
    IntunePresentFilterExcluded = $filterExcluded.Where({ $_.FilterReason -match 'INTUNE_DEVICE_PRESENT' }).Count
    ADLastLogonFilterExcluded  = $filterExcluded.Where({ $_.FilterReason -match 'AD_LAST_LOGON_' }).Count
    ADLastLogonUnknownExcluded = $filterExcluded.Where({ $_.FilterReason -match 'AD_LAST_LOGON_UNKNOWN' }).Count
    ADDomainFilterExcluded     = $filterExcluded.Where({ $_.FilterReason -match 'AD_DOMAIN_' }).Count
    ADRows                     = $adRows.Count
    IntuneRows                 = $intuneRows.Count
    ADWindows10Candidates      = $adRecords.Where({ $_.Eligible -and $matchedKeyLookup.ContainsKey($_.Key) }).Count
    IntuneWindows10Candidates  = $intuneRecords.Where({ $_.Eligible -and $matchedKeyLookup.ContainsKey($_.Key) }).Count
    SelectedDevices            = $selected.Count
    ExcludedDevices            = $excluded.Count
    Windows11Excluded          = $excluded.Where({ $_.ExclusionReason -match 'WINDOWS11_REPORTED' }).Count
    ADDisabledExcluded         = $excluded.Where({ $_.ExclusionReason -match 'AD_DISABLED' }).Count
    IntuneStateExcluded        = $excluded.Where({ $_.ExclusionReason -match 'INTUNE_MANAGEMENT_STATE_EXCLUDED' }).Count
    UnknownOSExcluded          = $excluded.Where({ $_.ExclusionReason -match 'OS_UNKNOWN_OR_UNSUPPORTED' }).Count
    ADNameCollisions           = $excluded.Where({ $_.ExclusionReason -match 'AD_NAME_COLLISION' }).Count
    IntuneDuplicateRowsIgnored = $intuneDuplicateCount
    ADStaleWarnings            = $adRecords.Where({ $_.IsStale -and $matchedKeyLookup.ContainsKey($_.Key) }).Count
    IntuneStaleWarnings        = $intuneRecords.Where({ $_.IsStale -and $matchedKeyLookup.ContainsKey($_.Key) }).Count
    WaveCreated                = [bool]$Create
    WaveName                   = $safeWaveName
    WavePath                   = $wavePath
    ComputersPath              = $computersPath
    EvidencePath               = $evidencePath
}

if (-not $NoEvidence) {
    Publish-AutomaticWaveProgress -Stage 'Finalizing automatic Wave evidence...' -Detail $evidencePath
    Write-SemtJsonFileAtomic -Path (Join-Path $evidencePath 'AutomaticWaveSummary.json.txt') -InputObject $summary -Depth 5
}

Publish-AutomaticWaveProgress -Stage $(if ($Create) { 'Automatic Wave created.' } else { 'Automatic Wave preview ready.' }) -Detail $(if ($Create) { $wavePath } else { "Selected devices: $($selected.Count)" })

[pscustomobject]@{
    Summary               = $summary
    SelectedDevices       = $selected.ToArray()
    ExcludedDevices       = $excluded.ToArray()
    FilterExcludedDevices = $filterExcluded.ToArray()
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCNArgNw+ubqFik
# Fv8Udlz1KfKzuR8yhGZvIVfzBOD9HKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCAOrdw4SpQTcHMZLn46nNs
# KaZSan6A5qlwhuG5C5BwSjANBgkqhkiG9w0BAQEFAASCAYAiD9uw8sLV0LqlwktL
# Ebw2KnEvx+uwsH3a1kVKG3i6T1j9A4q9bDeVpHIFGRuO7sz0opoGNCnZsf+X0Q+B
# 8nX4vSeo5D5B9IAAdgrk8sEXJyV2L8LhNz2pKGDKGuf2WnJhqYnNF67J5skWzZLF
# JHF0IZ1pEa/k6UIIYehSUmIGhK0zsK788hTrhLodbJkFZ22TJzzfHNQUaat5Ai4R
# YBng/bgcmV54wrXhFXWf6oSkkDUXXkYt0X9zU0qWrOccX5f63AyiMQY962WrkxW5
# +utUPQkrnGjicWMYba3utc+mTmGg3XZ32FmByMgBL1VQ7UpMNJsZDQVk1oBneEho
# ebpjRRJb2PoMzwn1elqAabepIilThyKqrzqkRl9eJkksuDusnxc4Xgqd9Bc8FRdO
# /mhOS6JG1NkrUR3aUURAoybQcCu4oA98s4zGyfbtXsMnET00sJUh2uFWXoLirvLz
# rexXOIz+QNH8F7HlWDgnJWqP3cY0f5rep0wamytCiVePT80=
# SIG # End signature block
