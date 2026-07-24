<#
.SYNOPSIS
Builds a Windows 10 LOT from AD and/or Intune inventory CSV files.

.DESCRIPTION
Selects explicit Windows 10 devices from broad AD and Intune inventory CSV files,
deduplicates them, excludes Windows 11 evidence and unsafe inventory states, writes
selection evidence under Runs, and can create a standard operational LOT.

.VERSION
1.0.0
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet('AD', 'Intune', 'Both')]
    [string]$Source = 'Both',
    [string]$AdInventoryCsv,
    [string]$IntuneInventoryCsv,
    [string]$ToolkitRoot,
    [string]$LotName,
    [string]$EvidenceRoot,
    [switch]$Create,
    [switch]$AllowPartialSource,
    [switch]$SkipWrapperRefresh,
    [switch]$NoEvidence
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:AutomaticLotVersion = '1.0.0'

function Get-AutomaticLotToolkitRoot {
    param([AllowNull()][string]$RequestedRoot)

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        return [System.IO.Path]::GetFullPath($RequestedRoot)
    }

    $scriptsRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    return Split-Path -Parent $scriptsRoot
}

function Get-AutomaticLotSafeName {
    param([AllowNull()][string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = 'LOT-AUTO-W10-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    }

    $safeName = [regex]::Replace($Name.Trim(), '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        throw 'The automatic LOT name is empty after sanitization.'
    }
    if ($safeName -notmatch '^(?i)LOT-') {
        $safeName = "LOT-$safeName"
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

function Get-AdAutomaticLotRecord {
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

function Get-IntuneAutomaticLotRecord {
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

function Get-AutomaticLotExclusionReason {
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

function Export-AutomaticLotCsv {
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

$effectiveToolkitRoot = Get-AutomaticLotToolkitRoot -RequestedRoot $ToolkitRoot
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
    throw ('Required automatic LOT inventory source(s) missing: {0}.' -f ($missingSources.ToArray() -join ', '))
}

$adRows = if ('AD' -in $availableSources) { @(Import-Csv -LiteralPath $AdInventoryCsv) } else { @() }
$intuneRows = if ('Intune' -in $availableSources) { @(Import-Csv -LiteralPath $IntuneInventoryCsv) } else { @() }
$adRows = @($adRows)
$intuneRows = @($intuneRows)
$adRecords = @(Get-AdAutomaticLotRecord -Rows $adRows)
$intuneRecords = @(Get-IntuneAutomaticLotRecord -Rows $intuneRows)

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
$selected = New-Object System.Collections.Generic.List[object]
$excluded = New-Object System.Collections.Generic.List[object]

foreach ($key in $allKeys) {
    $keyAdRecords = @()
    if ($adByKey.ContainsKey($key)) {
        $keyAdRecords = $adByKey[$key].ToArray()
    }
    $keyIntuneRecords = if ($intuneAllByKey.ContainsKey($key)) { @($intuneAllByKey[$key]) } else { @() }
    $intuneRecord = if ($intuneByKey.ContainsKey($key)) { $intuneByKey[$key] } else { $null }
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
        ExclusionReason    = Get-AutomaticLotExclusionReason -AdRecords $keyAdRecords -IntuneRecords $keyIntuneRecords -AdCollision $adCollision
        ADFQDNs             = $adFqdns -join ';'
        ADOperatingSystem   = if ($keyAdRecords.Count -gt 0) { [string]$keyAdRecords[0].OperatingSystem } else { '' }
        IntuneOperatingSystem = if ($null -ne $intuneRecord) { [string]$intuneRecord.OperatingSystem } else { '' }
        IntuneOSVersion     = if ($null -ne $intuneRecord) { [string]$intuneRecord.OSVersion } else { '' }
        IntuneManagementState = if ($null -ne $intuneRecord) { [string]$intuneRecord.ManagementState } else { '' }
    })
}

$safeLotName = Get-AutomaticLotSafeName -Name $LotName
$lotPath = ''
$computersPath = ''
if ($Create) {
    if ($selected.Count -eq 0) { throw 'No eligible Windows 10 device was selected; the LOT was not created.' }

    $lotsRoot = Join-Path $effectiveToolkitRoot 'Lots'
    New-Item -ItemType Directory -Path $lotsRoot -Force | Out-Null
    $lotPath = Join-Path $lotsRoot $safeLotName
    if (Test-Path -LiteralPath $lotPath) { throw "LOT folder already exists: $lotPath" }

    New-Item -ItemType Directory -Path $lotPath -Force | Out-Null
    $computersPath = Join-Path $lotPath 'Computers.txt'
    @($selected | Sort-Object ComputerName | Select-Object -ExpandProperty ComputerName) |
        Set-Content -LiteralPath $computersPath -Encoding ASCII

    $configTemplate = Join-Path $effectiveToolkitRoot 'Windows11UpgradeToolkit.config.template'
    $lotConfigPath = Join-Path $lotPath 'Windows11UpgradeToolkit.config'
    if (Test-Path -LiteralPath $configTemplate -PathType Leaf) {
        Copy-Item -LiteralPath $configTemplate -Destination $lotConfigPath -Force
    }
    else {
        Set-Content -LiteralPath $lotConfigPath -Encoding ASCII -Value @(
            '# SmartM365 Windows 11 Upgrade Toolkit LOT defaults.'
            'W11UT_SETUP_SOURCE='
        )
    }

    if (-not $SkipWrapperRefresh) {
        $wrapperRefresh = Join-Path $effectiveToolkitRoot 'Scripts\SmartM365-Windows11Upgrade-Update-LotCmdWrappers.ps1'
        if (-not (Test-Path -LiteralPath $wrapperRefresh -PathType Leaf)) {
            throw "Wrapper refresh script not found: $wrapperRefresh"
        }
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $wrapperRefresh -ToolkitRoot $effectiveToolkitRoot
        if ($LASTEXITCODE -ne 0) {
            throw "LOT wrapper refresh failed with exit code $LASTEXITCODE."
        }
    }
}

$evidencePath = ''
if (-not $NoEvidence) {
    if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
        $EvidenceRoot = Join-Path $effectiveToolkitRoot 'Runs\AutomaticLotInventory'
    }
    $evidencePath = Join-Path $EvidenceRoot (Get-Date -Format 'yyyyMMdd-HHmmssfff')
    New-Item -ItemType Directory -Path $evidencePath -Force | Out-Null

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
    Export-AutomaticLotCsv -Path (Join-Path $evidencePath 'AutomaticLotSelection.csv') -Columns $selectionColumns -Rows $selected.ToArray()
    Export-AutomaticLotCsv -Path (Join-Path $evidencePath 'AutomaticLotExclusions.csv') -Columns $exclusionColumns -Rows $excluded.ToArray()
}

$summary = [pscustomobject]@{
    Version                    = $script:AutomaticLotVersion
    RequestedSource            = $Source
    AvailableSources           = $availableSources.ToArray() -join '+'
    MissingSources             = $missingSources.ToArray() -join ','
    PartialSource              = ($missingSources.Count -gt 0)
    ADRows                     = $adRows.Count
    IntuneRows                 = $intuneRows.Count
    ADWindows10Candidates      = @($adRecords | Where-Object { $_.Eligible }).Count
    IntuneWindows10Candidates  = @($intuneRecords | Where-Object { $_.Eligible }).Count
    SelectedDevices            = $selected.Count
    ExcludedDevices            = $excluded.Count
    Windows11Excluded          = @($excluded | Where-Object { $_.ExclusionReason -match 'WINDOWS11_REPORTED' }).Count
    ADDisabledExcluded         = @($excluded | Where-Object { $_.ExclusionReason -match 'AD_DISABLED' }).Count
    IntuneStateExcluded        = @($excluded | Where-Object { $_.ExclusionReason -match 'INTUNE_MANAGEMENT_STATE_EXCLUDED' }).Count
    UnknownOSExcluded          = @($excluded | Where-Object { $_.ExclusionReason -match 'OS_UNKNOWN_OR_UNSUPPORTED' }).Count
    ADNameCollisions           = @($excluded | Where-Object { $_.ExclusionReason -match 'AD_NAME_COLLISION' }).Count
    IntuneDuplicateRowsIgnored = $intuneDuplicateCount
    ADStaleWarnings            = @($adRecords | Where-Object { $_.IsStale }).Count
    IntuneStaleWarnings        = @($intuneRecords | Where-Object { $_.IsStale }).Count
    LotName                    = $safeLotName
    LotPath                    = $lotPath
    ComputersPath              = $computersPath
    EvidencePath               = $evidencePath
}

if (-not $NoEvidence) {
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $evidencePath 'AutomaticLotSummary.json') -Encoding UTF8
}

[pscustomobject]@{
    Summary         = $summary
    SelectedDevices = $selected.ToArray()
    ExcludedDevices = $excluded.ToArray()
}
