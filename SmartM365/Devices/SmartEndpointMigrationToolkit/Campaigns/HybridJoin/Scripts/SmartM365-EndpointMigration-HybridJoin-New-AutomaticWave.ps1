<#
.SYNOPSIS
Builds a guarded HybridJoin Wave from AD, Intune, and Entra inventories.

.DESCRIPTION
Selects enabled Windows 10 and Windows 11 client computers from Active Directory,
excludes devices already present in Intune and ambiguous inventory states, classifies
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
    [string]$AdInventoryCsv,
    [string]$IntuneInventoryCsv,
    [string]$EntraInventoryCsv,
    [string]$PackageRoot,
    [string]$WaveName,
    [string[]]$ComputerNamePrefix,
    [string[]]$ComputerNameContains,
    [switch]$ExcludeStaleAd,
    [ValidateRange(1, 3650)][int]$AdLastLogonMaxAgeDays = 45,
    [scriptblock]$ProgressCallback,
    [string]$EvidenceRoot,
    [switch]$Create,
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
        $segments = New-Object System.Collections.Generic.List[string]
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

    $normalized = New-Object System.Collections.Generic.List[string]
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

    $normalized = New-Object System.Collections.Generic.List[string]
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
        $lastSeen = ConvertTo-InventoryDate -Value (Get-InventoryValue -Row $row -Names @('LastLogonTimestampUtc'))
        [pscustomobject]@{
            Key                = $key
            Name               = $computerName
            DNSHostName        = $dnsHostName
            PreferredName      = $identity
            Domain             = [string](Get-InventoryValue -Row $row -Names @('ADDomain'))
            DistinguishedName  = [string](Get-InventoryValue -Row $row -Names @('DistinguishedName'))
            Present            = $present
            Enabled            = $enabled
            Classification     = Get-WindowsClientClassification -OperatingSystem $operatingSystem
            OperatingSystem    = $operatingSystem
            OSVersion          = $operatingSystemVersion
            LastSeenUtc        = $lastSeen
        }
    }
}

function Get-IntuneAutomaticWaveRecord {
    param([object[]]$Rows)

    foreach ($row in @($Rows)) {
        $identity = [string](Get-InventoryValue -Row $row -Names @('DeviceName', 'ManagedDeviceName', 'ComputerName'))
        $key = Get-InventoryShortName -Name $identity
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        [pscustomobject]@{
            Key             = $key
            DeviceName      = $identity
            Present         = ConvertTo-InventoryBoolean -Value (Get-InventoryValue -Row $row -Names @('IntuneInventoryPresent')) -Default $true
            ManagedDeviceId = [string](Get-InventoryValue -Row $row -Names @('IntuneManagedDeviceId', 'Id'))
            LastSyncUtc     = ConvertTo-InventoryDate -Value (Get-InventoryValue -Row $row -Names @('LastSyncDateTime'))
        }
    }
}

function Get-EntraAutomaticWaveRecord {
    param([object[]]$Rows)

    foreach ($row in @($Rows)) {
        $identity = [string](Get-InventoryValue -Row $row -Names @('DisplayName', 'ComputerName'))
        $key = Get-InventoryShortName -Name $identity
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $present = ConvertTo-InventoryBoolean -Value (Get-InventoryValue -Row $row -Names @('EntraInventoryPresent')) -Default $true
        [pscustomobject]@{
            Key                       = $key
            DisplayName               = $identity
            Present                   = $present
            TrustType                 = [string](Get-InventoryValue -Row $row -Names @('TrustType'))
            RegisteredState           = [string](Get-InventoryValue -Row $row -Names @('EntraRegisteredState'))
            AlternativeSecurityIdCount = [string](Get-InventoryValue -Row $row -Names @('AlternativeSecurityIdCount'))
            AccountEnabled            = ConvertTo-InventoryBoolean -Value (Get-InventoryValue -Row $row -Names @('AccountEnabled')) -Default $true
            DeviceId                  = [string](Get-InventoryValue -Row $row -Names @('DeviceId'))
            ObjectId                  = [string](Get-InventoryValue -Row $row -Names @('EntraObjectId', 'Id'))
        }
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
    if (-not $adByKey.ContainsKey($record.Key)) { $adByKey[$record.Key] = New-Object System.Collections.Generic.List[object] }
    $adByKey[$record.Key].Add($record)
}
$intuneByKey = @{}
foreach ($record in $intuneRecords) {
    if (-not $intuneByKey.ContainsKey($record.Key)) { $intuneByKey[$record.Key] = New-Object System.Collections.Generic.List[object] }
    $intuneByKey[$record.Key].Add($record)
}
$entraByKey = @{}
foreach ($record in $entraRecords) {
    if (-not $entraByKey.ContainsKey($record.Key)) { $entraByKey[$record.Key] = New-Object System.Collections.Generic.List[object] }
    $entraByKey[$record.Key].Add($record)
}

$namePrefixes = @(Get-AutomaticWaveNamePrefixes -Prefixes $ComputerNamePrefix)
$nameContainsValues = @(Get-AutomaticWaveNameContainsValues -ContainsValues $ComputerNameContains)
$namePrefixDisplay = $namePrefixes -join ';'
$nameContainsDisplay = $nameContainsValues -join ';'
$adLastLogonCutoffUtc = [datetime]::UtcNow.AddDays(-$AdLastLogonMaxAgeDays)
$selected = New-Object System.Collections.Generic.List[object]
$excluded = New-Object System.Collections.Generic.List[object]
$filterExcluded = New-Object System.Collections.Generic.List[object]

foreach ($key in @($adByKey.Keys | Sort-Object)) {
    $keyAdRecords = @($adByKey[$key].ToArray())
    $keyIntuneRecords = if ($intuneByKey.ContainsKey($key)) { @($intuneByKey[$key].ToArray()) } else { @() }
    $keyEntraRecords = if ($entraByKey.ContainsKey($key)) { @($entraByKey[$key].ToArray()) } else { @() }
    $preferredAd = @($keyAdRecords | Sort-Object -Property DNSHostName | Select-Object -First 1)[0]
    $filterReasons = New-Object System.Collections.Generic.List[string]
    $filterTypes = New-Object System.Collections.Generic.List[string]
    $filterValues = New-Object System.Collections.Generic.List[string]

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
    $validLastSeen = @($keyAdRecords | Where-Object { $_.LastSeenUtc -ne [datetime]::MinValue } | Sort-Object LastSeenUtc -Descending)
    $latestLastSeen = if ($validLastSeen.Count -gt 0) { [datetime]$validLastSeen[0].LastSeenUtc } else { [datetime]::MinValue }
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

    $reasons = New-Object System.Collections.Generic.List[string]
    $adObjectKeys = @(
        $keyAdRecords |
            ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace([string]$_.DistinguishedName)) { ([string]$_.DistinguishedName).ToUpperInvariant() }
                elseif (-not [string]::IsNullOrWhiteSpace([string]$_.DNSHostName)) { ([string]$_.DNSHostName).ToUpperInvariant() }
                else { ([string]$_.Name).ToUpperInvariant() }
            } |
            Sort-Object -Unique
    )
    if ($adObjectKeys.Count -gt 1) { $reasons.Add('AD_NAME_COLLISION') }
    if (@($keyAdRecords | Where-Object { -not $_.Present }).Count -gt 0) { $reasons.Add('AD_INVENTORY_NOT_PRESENT') }
    if (@($keyAdRecords | Where-Object { $_.Present -and -not $_.Enabled }).Count -gt 0) { $reasons.Add('AD_DISABLED') }
    if ($preferredAd.Classification -eq 'WindowsServer') { $reasons.Add('WINDOWS_SERVER_EXCLUDED') }
    elseif ($preferredAd.Classification -notin @('Windows10', 'Windows11')) { $reasons.Add('OS_UNKNOWN_OR_UNSUPPORTED') }
    if (@($keyIntuneRecords | Where-Object { $_.Present }).Count -gt 0) { $reasons.Add('INTUNE_DEVICE_PRESENT') }

    $serverAdRecords = @($keyEntraRecords | Where-Object { $_.Present -and $_.TrustType -ieq 'ServerAd' })
    if ($serverAdRecords.Count -gt 1) { $reasons.Add('ENTRA_SERVERAD_COLLISION') }
    if ($serverAdRecords.Count -eq 1 -and -not $serverAdRecords[0].AccountEnabled) { $reasons.Add('ENTRA_SERVERAD_DISABLED') }

    if ($reasons.Count -gt 0) {
        $excluded.Add([pscustomobject]@{
            ComputerKey = $key
            ComputerName = [string]$preferredAd.PreferredName
            ExclusionReason = $reasons.ToArray() -join ';'
            ADFQDNs = @($keyAdRecords | Select-Object -ExpandProperty DNSHostName -Unique) -join ';'
            ADOperatingSystem = [string]$preferredAd.OperatingSystem
            ADEnabled = [string]$preferredAd.Enabled
            IntuneMatches = @($keyIntuneRecords | Where-Object { $_.Present }).Count
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
$safeWaveName = Get-AutomaticWaveSafeName -Name $WaveName -NamePrefixes $namePrefixes
$wavePath = ''
$computersPath = ''
if ($Create) {
    if ($selected.Count -eq 0) { throw 'No eligible Hybrid Join device was selected; the Wave was not created.' }
    if (-not (Test-SemtWaveName -Name $safeWaveName)) { throw ("SEMT_WAVE_NAME_INVALID: '{0}' is not a valid Wave name." -f $safeWaveName) }
    $plannedWavePath = Join-Path (Get-SemtWaveRoot -PackageRoot $effectivePackageRoot -Campaign $script:AutomaticWaveCampaign) $safeWaveName
    if (Test-Path -LiteralPath $plannedWavePath) { throw ("SEMT_WAVE_EXISTS: Wave '{0}' already exists for {1}: {2}" -f $safeWaveName, $script:AutomaticWaveCampaign, $plannedWavePath) }
}

$evidencePath = ''
if (-not $NoEvidence) {
    if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
        $EvidenceRoot = Join-Path $effectivePackageRoot ('Runs\{0}\_AutomaticWave' -f $script:AutomaticWaveCampaign)
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

$availableSourcesText = if ($entraAvailable) { 'AD+Intune+Entra' } else { 'AD+Intune' }
if ($Create) {
    $description = 'Automatic HybridJoin Wave; Sources={0}; ComputerNamePrefix={1}; ComputerNameContains={2}; ExcludeIntunePresent=True; ExcludeStaleAd={3}; ADLastLogonMaxAgeDays={4}; Selected={5}' -f $availableSourcesText, $namePrefixDisplay, $nameContainsDisplay, [bool]$ExcludeStaleAd, $AdLastLogonMaxAgeDays, $selected.Count
    Publish-AutomaticWaveProgress -Stage 'Creating the automatic Wave...' -Detail ("Wave: {0}; selected devices: {1}" -f $safeWaveName, $selected.Count)
    $waveComputers = @($selected | Sort-Object ComputerName | Select-Object -ExpandProperty ComputerName)
    $wave = New-SemtWave -PackageRoot $effectivePackageRoot -Campaign $script:AutomaticWaveCampaign -Name $safeWaveName -Computers $waveComputers -Source Automatic -Description $description -EvidencePath $evidencePath
    $wavePath = [string]$wave.Path
    $computersPath = [string]$wave.ComputersPath
}

$summary = [pscustomobject]@{
    Version = $script:AutomaticWaveVersion
    Campaign = $script:AutomaticWaveCampaign
    AvailableSources = $availableSourcesText
    EntraInventoryAvailable = $entraAvailable
    ComputerNamePrefixes = $namePrefixDisplay
    ComputerNameContains = $nameContainsDisplay
    ExcludeIntunePresent = $true
    ExcludeStaleAd = [bool]$ExcludeStaleAd
    ADLastLogonMaxAgeDays = $AdLastLogonMaxAgeDays
    ADRows = $adRows.Count
    IntuneRows = $intuneRows.Count
    EntraRows = $entraRows.Count
    UniqueADDevices = $adByKey.Count
    SelectedDevices = $selected.Count
    ExcludedDevices = $excluded.Count
    FilterExcludedDevices = $filterExcluded.Count
    NeedsHybridJoin = @($selected | Where-Object { $_.SelectionReason -eq 'NEEDS_HYBRID_JOIN' }).Count
    HybridJoinPending = @($selected | Where-Object { $_.SelectionReason -eq 'HYBRID_JOIN_PENDING' }).Count
    NeedsIntuneEnrollment = @($selected | Where-Object { $_.SelectionReason -eq 'NEEDS_INTUNE_ENROLLMENT' }).Count
    EntraUnavailableSelected = @($selected | Where-Object { $_.SelectionReason -eq 'ENTRA_INVENTORY_UNAVAILABLE' }).Count
    IntunePresentExcluded = @($excluded | Where-Object { $_.ExclusionReason -match 'INTUNE_DEVICE_PRESENT' }).Count
    ADDisabledExcluded = @($excluded | Where-Object { $_.ExclusionReason -match 'AD_DISABLED' }).Count
    WindowsServerExcluded = @($excluded | Where-Object { $_.ExclusionReason -match 'WINDOWS_SERVER_EXCLUDED' }).Count
    UnknownOSExcluded = @($excluded | Where-Object { $_.ExclusionReason -match 'OS_UNKNOWN_OR_UNSUPPORTED' }).Count
    ADNameCollisions = @($excluded | Where-Object { $_.ExclusionReason -match 'AD_NAME_COLLISION' }).Count
    EntraAmbiguousExcluded = @($excluded | Where-Object { $_.ExclusionReason -match 'ENTRA_SERVERAD_' }).Count
    ADLastLogonFilterExcluded = @($filterExcluded | Where-Object { $_.FilterReason -match 'AD_LAST_LOGON_' }).Count
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
