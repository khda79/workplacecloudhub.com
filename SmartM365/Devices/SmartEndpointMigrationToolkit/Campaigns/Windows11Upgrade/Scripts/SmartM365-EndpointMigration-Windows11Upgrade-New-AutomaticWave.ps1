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

.PARAMETER Create
Creates the Wave (Computers.txt, Wave.local.json, wrappers). Without it, only evidence is written.

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
    [switch]$ExcludeIntunePresent,
    [switch]$ExcludeStaleAd,
    [ValidateRange(1, 3650)]
    [int]$AdLastLogonMaxAgeDays = 45,
    [scriptblock]$ProgressCallback,
    [string]$EvidenceRoot,
    [switch]$Create,
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
        $nameSegments = New-Object System.Collections.Generic.List[string]
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

    $normalized = New-Object System.Collections.Generic.List[string]
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

    $normalized = New-Object System.Collections.Generic.List[string]
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

function Get-AdAutomaticWaveRecord {
    param([object[]]$Rows)

    foreach ($row in @($Rows)) {
        $computerName = [string](Get-InventoryValue -Row $row -Names @('ComputerName', 'Name'))
        $dnsHostName = [string](Get-InventoryValue -Row $row -Names @('DNSHostName'))
        $identity = if (-not [string]::IsNullOrWhiteSpace($dnsHostName)) { $dnsHostName } else { $computerName }
        $key = Get-InventoryShortName -Name $identity
        if ([string]::IsNullOrWhiteSpace($key)) { continue }

        $operatingSystem = [string](Get-InventoryValue -Row $row -Names @('OperatingSystem'))
        $operatingSystemVersion = [string](Get-InventoryValue -Row $row -Names @('OperatingSystemVersion', 'OSVersion'))
        $present = ConvertTo-InventoryBoolean -Value (Get-InventoryValue -Row $row -Names @('ADInventoryPresent')) -Default $true
        $enabled = ConvertTo-InventoryBoolean -Value (Get-InventoryValue -Row $row -Names @('Enabled')) -Default $true
        $classification = Get-InventoryWindowsClassification -OperatingSystem $operatingSystem -Version $operatingSystemVersion
        $lastSeen = ConvertTo-InventoryDate -Value (Get-InventoryValue -Row $row -Names @('LastLogonTimestampUtc'))

        [pscustomobject]@{
            Key             = $key
            Name            = $computerName
            DNSHostName     = $dnsHostName
            PreferredName   = $identity
            Domain          = [string](Get-InventoryValue -Row $row -Names @('ADDomain'))
            Present         = $present
            Enabled         = $enabled
            Classification  = $classification
            OperatingSystem = $operatingSystem
            OSVersion       = $operatingSystemVersion
            LastSeenUtc     = $lastSeen
            IsStale         = ($lastSeen -ne [datetime]::MinValue -and $lastSeen -lt [datetime]::UtcNow.AddDays(-90))
            Eligible        = ($present -and $enabled -and $classification -eq 'Windows10')
        }
    }
}

function Get-IntuneAutomaticWaveRecord {
    param([object[]]$Rows)

    foreach ($row in @($Rows)) {
        $deviceName = [string](Get-InventoryValue -Row $row -Names @('DeviceName', 'ManagedDeviceName', 'ComputerName'))
        $computerName = [string](Get-InventoryValue -Row $row -Names @('ComputerName'))
        $identity = if (-not [string]::IsNullOrWhiteSpace($deviceName)) { $deviceName } else { $computerName }
        $key = Get-InventoryShortName -Name $identity
        if ([string]::IsNullOrWhiteSpace($key)) { continue }

        $operatingSystem = [string](Get-InventoryValue -Row $row -Names @('OperatingSystem'))
        $operatingSystemVersion = [string](Get-InventoryValue -Row $row -Names @('OSVersion', 'OperatingSystemVersion'))
        $present = ConvertTo-InventoryBoolean -Value (Get-InventoryValue -Row $row -Names @('IntuneInventoryPresent')) -Default $true
        $managementState = [string](Get-InventoryValue -Row $row -Names @('ManagementState'))
        $stateExcluded = (-not [string]::IsNullOrWhiteSpace($managementState) -and $managementState -match '(?i)(retire|wipe|delete)')
        $classification = Get-InventoryWindowsClassification -OperatingSystem $operatingSystem -Version $operatingSystemVersion
        $lastSync = ConvertTo-InventoryDate -Value (Get-InventoryValue -Row $row -Names @('LastSyncDateTime'))

        [pscustomobject]@{
            Key             = $key
            DeviceName      = $deviceName
            ComputerName    = $computerName
            PreferredName   = $identity
            ManagedDeviceId = [string](Get-InventoryValue -Row $row -Names @('IntuneManagedDeviceId', 'Id'))
            Present         = $present
            ManagementState = $managementState
            StateExcluded   = $stateExcluded
            Classification  = $classification
            OperatingSystem = $operatingSystem
            OSVersion       = $operatingSystemVersion
            LastSyncUtc     = $lastSync
            IsStale         = ($lastSync -ne [datetime]::MinValue -and $lastSync -lt [datetime]::UtcNow.AddDays(-30))
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

    $AdRecords = @($AdRecords | Where-Object { $null -ne $_ })
    $IntuneRecords = @($IntuneRecords | Where-Object { $null -ne $_ })
    $reasons = New-Object System.Collections.Generic.List[string]
    if ($AdCollision) { $reasons.Add('AD_NAME_COLLISION') }
    if (@($AdRecords | Where-Object { $_.Classification -eq 'Windows11' }).Count -gt 0 -or
        @($IntuneRecords | Where-Object { $_.Classification -eq 'Windows11' }).Count -gt 0) {
        $reasons.Add('WINDOWS11_REPORTED')
    }
    if (@($AdRecords | Where-Object { $_.Present -and -not $_.Enabled }).Count -gt 0) {
        $reasons.Add('AD_DISABLED')
    }
    if (@($IntuneRecords | Where-Object { $_.StateExcluded }).Count -gt 0) {
        $reasons.Add('INTUNE_MANAGEMENT_STATE_EXCLUDED')
    }
    if (@($AdRecords | Where-Object { -not $_.Present }).Count -gt 0 -or
        @($IntuneRecords | Where-Object { -not $_.Present }).Count -gt 0) {
        $reasons.Add('INVENTORY_NOT_PRESENT')
    }
    if (@($AdRecords | Where-Object { $_.Classification -eq 'Unknown' }).Count -gt 0 -or
        @($IntuneRecords | Where-Object { $_.Classification -eq 'Unknown' }).Count -gt 0) {
        $reasons.Add('OS_UNKNOWN_OR_UNSUPPORTED')
    }
    if ($reasons.Count -eq 0) { $reasons.Add('NOT_ELIGIBLE_WINDOWS10') }
    return @($reasons | Select-Object -Unique) -join ';'
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

$availableSources = New-Object System.Collections.Generic.List[string]
$missingSources = New-Object System.Collections.Generic.List[string]
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
if ($ExcludeStaleAd -and 'AD' -notin $availableSources) {
    throw 'ExcludeStaleAd requires an available AD inventory source.'
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
        $adByKey[$record.Key] = New-Object System.Collections.Generic.List[object]
    }
    $adByKey[$record.Key].Add($record)
}

$intuneByKey = @{}
$intuneAllByKey = @{}
$intuneDuplicateCount = 0
foreach ($group in @($intuneRecords | Group-Object -Property Key)) {
    $ordered = @($group.Group | Sort-Object -Property @{ Expression = 'LastSyncUtc'; Descending = $true })
    if ($ordered.Count -gt 1) { $intuneDuplicateCount += ($ordered.Count - 1) }
    $intuneAllByKey[$group.Name] = $ordered
    $intuneByKey[$group.Name] = $ordered[0]
}

$allKeys = @(@($adByKey.Keys) + @($intuneByKey.Keys) | Sort-Object -Unique)
$namePrefixes = @(Get-AutomaticWaveNamePrefixes -Prefixes $ComputerNamePrefix)
$namePrefixDisplay = $namePrefixes -join ';'
$nameContainsValues = @(Get-AutomaticWaveNameContainsValues -ContainsValues $ComputerNameContains)
$nameContainsDisplay = $nameContainsValues -join ';'
$adLastLogonCutoffUtc = [datetime]::UtcNow.AddDays(-$AdLastLogonMaxAgeDays)
$selected = New-Object System.Collections.Generic.List[object]
$excluded = New-Object System.Collections.Generic.List[object]
$filterExcluded = New-Object System.Collections.Generic.List[object]
$matchedKeyLookup = @{}

foreach ($key in $allKeys) {
    $keyAdRecords = @()
    if ($adByKey.ContainsKey($key)) {
        $keyAdRecords = $adByKey[$key].ToArray()
    }
    $keyIntuneRecords = if ($intuneAllByKey.ContainsKey($key)) { @($intuneAllByKey[$key]) } else { @() }
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
    $filterTypes = New-Object System.Collections.Generic.List[string]
    $filterValues = New-Object System.Collections.Generic.List[string]
    $filterReasons = New-Object System.Collections.Generic.List[string]
    if (-not (Test-AutomaticWaveNamePrefix -ComputerKey $key -Prefixes $namePrefixes)) {
        $filterTypes.Add('ComputerNamePrefix')
        $filterValues.Add($namePrefixDisplay)
        $filterReasons.Add('COMPUTER_NAME_PREFIX_NOT_MATCHED')
    }
    if (-not (Test-AutomaticWaveNameContains -ComputerKey $key -ContainsValues $nameContainsValues)) {
        $filterTypes.Add('ComputerNameContains')
        $filterValues.Add($nameContainsDisplay)
        $filterReasons.Add('COMPUTER_NAME_CONTAINS_NOT_MATCHED')
    }

    $intunePresent = (@($keyIntuneRecords | Where-Object { $_.Present }).Count -gt 0)
    if ($ExcludeIntunePresent -and $intunePresent) {
        $filterTypes.Add('IntunePresence')
        $filterValues.Add('Present=True')
        $filterReasons.Add('INTUNE_DEVICE_PRESENT')
    }

    $validAdLastSeen = @($keyAdRecords | Where-Object { $_.LastSeenUtc -ne [datetime]::MinValue } | Sort-Object -Property LastSeenUtc -Descending)
    $latestAdLastSeen = if ($validAdLastSeen.Count -gt 0) { [datetime]$validAdLastSeen[0].LastSeenUtc } else { [datetime]::MinValue }
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

    $adFqdns = @($keyAdRecords | ForEach-Object { ([string]$_.DNSHostName).Trim().ToLowerInvariant() } | Where-Object { $_ -match '\.' } | Sort-Object -Unique)
    $adCollision = ($adFqdns.Count -gt 1)
    $windows11Evidence = (
        @($keyAdRecords | Where-Object { $_.Classification -eq 'Windows11' }).Count -gt 0 -or
        @($keyIntuneRecords | Where-Object { $_.Classification -eq 'Windows11' }).Count -gt 0
    )
    $adDisabledEvidence = (@($keyAdRecords | Where-Object { $_.Present -and -not $_.Enabled }).Count -gt 0)
    $intuneStateExcluded = (@($keyIntuneRecords | Where-Object { $_.StateExcluded }).Count -gt 0)
    $inventoryNotPresent = (
        @($keyAdRecords | Where-Object { -not $_.Present }).Count -gt 0 -or
        @($keyIntuneRecords | Where-Object { -not $_.Present }).Count -gt 0
    )
    $eligibleAd = @($keyAdRecords | Where-Object { $_.Eligible } | Sort-Object -Property DNSHostName | Select-Object -First 1)
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
            ADLastLogonStale   = (@($keyAdRecords | Where-Object { $_.IsStale }).Count -gt 0)
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
$safeWaveName = Get-AutomaticWaveSafeName -Name $WaveName -NamePrefixes $namePrefixes
$wavePath = ''
$computersPath = ''
if ($Create) {
    if ($selected.Count -eq 0) { throw 'No eligible Windows 10 device was selected; the Wave was not created.' }
    if (-not (Test-SemtWaveName -Name $safeWaveName)) { throw ("SEMT_WAVE_NAME_INVALID: '{0}' is not a valid Wave name." -f $safeWaveName) }
    $plannedWavePath = Join-Path (Get-SemtWaveRoot -PackageRoot $effectivePackageRoot -Campaign $script:AutomaticWaveCampaign) $safeWaveName
    if (Test-Path -LiteralPath $plannedWavePath) { throw ("SEMT_WAVE_EXISTS: Wave '{0}' already exists for {1}: {2}" -f $safeWaveName,$script:AutomaticWaveCampaign,$plannedWavePath) }
}

$evidencePath = ''
if (-not $NoEvidence) {
    if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
        $EvidenceRoot = Join-Path $effectivePackageRoot ('Runs\{0}\_AutomaticWave' -f $script:AutomaticWaveCampaign)
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
    $description = 'Automatic Windows11Upgrade Wave; Source={0}; AvailableSources={1}; ComputerNamePrefix={2}; ComputerNameContains={3}; ExcludeIntunePresent={4}; ExcludeStaleAd={5}; ADLastLogonMaxAgeDays={6}; Selected={7}' -f $Source,($availableSources.ToArray() -join '+'),$namePrefixDisplay,$nameContainsDisplay,[bool]$ExcludeIntunePresent,[bool]$ExcludeStaleAd,$AdLastLogonMaxAgeDays,$selected.Count
    Publish-AutomaticWaveProgress -Stage 'Creating the automatic Wave...' -Detail ("Wave: {0}; selected devices: {1}" -f $safeWaveName,$selected.Count)
    $waveComputers = @($selected | Sort-Object ComputerName | Select-Object -ExpandProperty ComputerName)
    $wave = New-SemtWave -PackageRoot $effectivePackageRoot -Campaign $script:AutomaticWaveCampaign -Name $safeWaveName -Computers $waveComputers -Source Automatic -Description $description -EvidencePath $evidencePath
    $wavePath = [string]$wave.Path
    $computersPath = [string]$wave.ComputersPath
}

$summary = [pscustomobject]@{
    Version                    = $script:AutomaticWaveVersion
    Campaign                   = $script:AutomaticWaveCampaign
    RequestedSource            = $Source
    AvailableSources           = $availableSources.ToArray() -join '+'
    MissingSources             = $missingSources.ToArray() -join ','
    PartialSource              = ($missingSources.Count -gt 0)
    ComputerNamePrefixes       = $namePrefixDisplay
    ComputerNameContains       = $nameContainsDisplay
    ExcludeIntunePresent       = [bool]$ExcludeIntunePresent
    ExcludeStaleAd             = [bool]$ExcludeStaleAd
    ADLastLogonMaxAgeDays      = $AdLastLogonMaxAgeDays
    NameFilterEnabled          = ($namePrefixes.Count -gt 0 -or $nameContainsValues.Count -gt 0 -or $ExcludeIntunePresent -or $ExcludeStaleAd)
    UniqueInventoryDevices     = $allKeys.Count
    NameFilterMatchedDevices   = $matchedKeyLookup.Count
    NameFilterExcludedDevices  = $filterExcluded.Count
    PrefixFilterExcluded       = @($filterExcluded | Where-Object { $_.FilterReason -match 'COMPUTER_NAME_PREFIX_NOT_MATCHED' }).Count
    ContainsFilterExcluded     = @($filterExcluded | Where-Object { $_.FilterReason -match 'COMPUTER_NAME_CONTAINS_NOT_MATCHED' }).Count
    IntunePresentFilterExcluded = @($filterExcluded | Where-Object { $_.FilterReason -match 'INTUNE_DEVICE_PRESENT' }).Count
    ADLastLogonFilterExcluded  = @($filterExcluded | Where-Object { $_.FilterReason -match 'AD_LAST_LOGON_' }).Count
    ADLastLogonUnknownExcluded = @($filterExcluded | Where-Object { $_.FilterReason -match 'AD_LAST_LOGON_UNKNOWN' }).Count
    ADRows                     = $adRows.Count
    IntuneRows                 = $intuneRows.Count
    ADWindows10Candidates      = @($adRecords | Where-Object { $_.Eligible -and $matchedKeyLookup.ContainsKey($_.Key) }).Count
    IntuneWindows10Candidates  = @($intuneRecords | Where-Object { $_.Eligible -and $matchedKeyLookup.ContainsKey($_.Key) }).Count
    SelectedDevices            = $selected.Count
    ExcludedDevices            = $excluded.Count
    Windows11Excluded          = @($excluded | Where-Object { $_.ExclusionReason -match 'WINDOWS11_REPORTED' }).Count
    ADDisabledExcluded         = @($excluded | Where-Object { $_.ExclusionReason -match 'AD_DISABLED' }).Count
    IntuneStateExcluded        = @($excluded | Where-Object { $_.ExclusionReason -match 'INTUNE_MANAGEMENT_STATE_EXCLUDED' }).Count
    UnknownOSExcluded          = @($excluded | Where-Object { $_.ExclusionReason -match 'OS_UNKNOWN_OR_UNSUPPORTED' }).Count
    ADNameCollisions           = @($excluded | Where-Object { $_.ExclusionReason -match 'AD_NAME_COLLISION' }).Count
    IntuneDuplicateRowsIgnored = $intuneDuplicateCount
    ADStaleWarnings            = @($adRecords | Where-Object { $_.IsStale -and $matchedKeyLookup.ContainsKey($_.Key) }).Count
    IntuneStaleWarnings        = @($intuneRecords | Where-Object { $_.IsStale -and $matchedKeyLookup.ContainsKey($_.Key) }).Count
    WaveCreated                = [bool]$Create
    WaveName                   = $safeWaveName
    WavePath                   = $wavePath
    ComputersPath              = $computersPath
    EvidencePath               = $evidencePath
}

if (-not $NoEvidence) {
    Publish-AutomaticWaveProgress -Stage 'Finalizing automatic Wave evidence...' -Detail $evidencePath
    Write-SemtJsonFileAtomic -Path (Join-Path $evidencePath 'AutomaticWaveSummary.json') -InputObject $summary -Depth 5
}

Publish-AutomaticWaveProgress -Stage $(if ($Create) { 'Automatic Wave created.' } else { 'Automatic Wave preview ready.' }) -Detail $(if ($Create) { $wavePath } else { "Selected devices: $($selected.Count)" })

[pscustomobject]@{
    Summary               = $summary
    SelectedDevices       = $selected.ToArray()
    ExcludedDevices       = $excluded.ToArray()
    FilterExcludedDevices = $filterExcluded.ToArray()
}
