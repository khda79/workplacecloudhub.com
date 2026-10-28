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

.PARAMETER AdDomainFilter
Optional AD domain DNS names (for example contoso.local). Only computers with an AD record in one of
these domains are kept; the others are listed in AutomaticWaveFilterExclusions.csv with
AD_DOMAIN_NOT_MATCHED, or AD_DOMAIN_UNKNOWN when they have no AD record or no domain. Each domain must
exist in the AD inventory. The shared AD inventory always covers the whole forest.

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
    [string[]]$AdDomainFilter,
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

function Get-AutomaticWaveAdDomains {
    param([AllowNull()][string[]]$Domains)

    $normalized = New-Object System.Collections.Generic.List[string]
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
    $recordDomains = @(@($AdRecords) | Where-Object { $null -ne $_ } | ForEach-Object { ([string]$_.Domain).Trim().Trim('.').ToLowerInvariant() } | Where-Object { $_ })
    if ($recordDomains.Count -eq 0) { return 'AD_DOMAIN_UNKNOWN' }
    foreach ($domain in $recordDomains) { if ($Domains -contains $domain) { return '' } }
    return 'AD_DOMAIN_NOT_MATCHED'
}

function Assert-AutomaticWaveAdDomainsInInventory {
    param(
        [string[]]$Domains,
        [object[]]$AdRecords,
        [string]$InventoryPath
    )

    if (@($Domains).Count -eq 0) { return }
    $known = @(@($AdRecords) | ForEach-Object { ([string]$_.Domain).Trim().Trim('.').ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
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
$adDomains = @(Get-AutomaticWaveAdDomains -Domains $AdDomainFilter)
$adDomainDisplay = $adDomains -join ';'
Assert-AutomaticWaveAdDomainsInInventory -Domains $adDomains -AdRecords $adRecords -InventoryPath $AdInventoryCsv
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
    $adDomainReason = Get-AutomaticWaveAdDomainFilterReason -AdRecords $keyAdRecords -Domains $adDomains
    if (-not [string]::IsNullOrWhiteSpace($adDomainReason)) {
        $filterTypes.Add('ADDomain')
        $filterValues.Add($adDomainDisplay)
        $filterReasons.Add($adDomainReason)
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
    $description = 'Automatic HybridJoin Wave; Sources={0}; ComputerNamePrefix={1}; ComputerNameContains={2}; ADDomain={6}; ExcludeIntunePresent=True; ExcludeStaleAd={3}; ADLastLogonMaxAgeDays={4}; Selected={5}' -f $availableSourcesText, $namePrefixDisplay, $nameContainsDisplay, [bool]$ExcludeStaleAd, $AdLastLogonMaxAgeDays, $selected.Count, $adDomainDisplay
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
    ADDomainFilter = $adDomainDisplay
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
    ADDomainFilterExcluded = @($filterExcluded | Where-Object { $_.FilterReason -match 'AD_DOMAIN_' }).Count
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDlUngywxOuQl6K
# hYBWk5KA+qoIsABkJbYipRgIyK+G26CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIF/ELE5jzAQvXEk1aKGuCYWZmM9VQYNov9E1GTeuc6hyMA0GCSqG
# SIb3DQEBAQUABIIBgAh3WQ7ksPV3uCfFXzG2euyD6I9KkJPlbx/ByP2erLXZaECe
# /IJhg95YgsAYEWUqS3VURxq9HCy0B4tW7BbkvXa5EKgZOKxbKFzQNiEAmrypCnuE
# cT1oqMvNC/nTa6hUI5A3orcJE+alPEI6/L/JhthQ5POVhNK/iMICgim0MdV7kf+z
# F///FuLOq3TnmOP+LhEycFdgrsSo/iKWCIV5BFNg8DGeC5fbNLpDihm+sxB1qtYo
# b4Dvn4XEZEccQUUPobREeZNPyPs6aR70ZvTvo0KsKMGuDKKZwCESI6a7LXJk134u
# W21UyCimag/3N3ud6AgD3hlrjd1/6Y7uF8PdKTYhWHzfRIOT8IEssGVhJX8dAvsK
# 9WnOVT0X3I1bwLOOjEUm52ayY00IlzafdEMWwR4ShHvjgUtGlTRof4E8h54Mhmhz
# 3S3b5rnokhbSHjcMdGgEjzAaDbG8pyvunsX4ghNV2JCwONO8V/69csy/shiEvMXk
# z4BEBhY3JpqPk+xuyaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NDdaMC8GCSqGSIb3DQEJBDEiBCCE+vV6bz7nLvXzGqvM8aYwxzjfkxMOu7UDeLvu
# HeUqLjANBgkqhkiG9w0BAQEFAASCAgCZRU1crHCqEcGH9k79lyKMs5m2WOV2cJtG
# MilVkfVHqtt+ACuvbUjOtuSoBE3zcdrIjs4yV3cVmVY1+fU6rWohVhoagKcqgaFk
# D71q/6gm3nRQW9onf919StP8KSGFo2O4gmsDARPwASWbU5RLEfs3AXzGlh+2pvk2
# tYIsEszFhiFGhM3VjjQ7gi2nTyFrb5BgCgvsrYQFNwrcg9w9FnXDozgkfuTqgZVo
# 3xtorNrNqULwvv5I2k0xcGMORIfBpGQhb2d1leHrA5Oq42VOR6HpZbIIG70kyqNA
# rvQ+gJv45eL5Co7xgf2gg/x80m8tDOoFxjnOTbUx0xGBbulOvmAtSeOIM8G2Vrhu
# ukLaX4nrW/GKXO2YZh0yhLkHFwh25/A2k3f7n97kOnlvOSFSSXxmasqh9zDPXbHH
# jl/OL3efvTpUeiw2MbV5cre8AsnZbTGsngst4V6ZtiDicUeJ2fG4/fIOjHGwCt+L
# TaCJM6FFCNjzUqTSxMWSvqDXgLlqTwzckLOvH+X1p23WdmLSwksH3Ln5p44mrvNh
# xk/rc9QQsrhFB+ZXsNk7jwZPaBWWNGYAzwYTyKbrdApIPM13Vu7W+U8jFQfRNOqc
# cOdr7w++AqT2e7ApbuXWUamqYAB4xoG00A9XlU00nZ7olaWFc+JtXi2L+oyEAPd2
# WfkwrYaN4g==
# SIG # End signature block
