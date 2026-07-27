<#
.SYNOPSIS
Builds a Windows 10 LOT from AD and/or Intune inventory CSV files.

.DESCRIPTION
Selects explicit Windows 10 devices from broad AD and Intune inventory CSV files,
deduplicates them, excludes Windows 11 evidence and unsafe inventory states, writes
selection evidence under Runs, and can create a standard operational LOT.

.VERSION
1.3.1
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
    [switch]$SkipWrapperRefresh,
    [switch]$NoEvidence
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:AutomaticLotVersion = '1.3.2'
$script:AutomaticLotProgressCallback = $ProgressCallback

function Publish-AutomaticLotProgress {
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [string]$Detail = ''
    )

    if ($script:AutomaticLotProgressCallback) {
        & $script:AutomaticLotProgressCallback $Stage $Detail | Out-Null
    }
}

function Invoke-AutomaticLotWrapperRefresh {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$EffectiveToolkitRoot
    )

    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) { $powerShellPath = 'powershell.exe' }
    $quotedScriptPath = '"{0}"' -f $ScriptPath.Replace('"', '\"')
    $quotedToolkitRoot = '"{0}"' -f $EffectiveToolkitRoot.Replace('"', '\"')
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $powerShellPath
    $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File $quotedScriptPath -ToolkitRoot $quotedToolkitRoot"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $startedUtc = [datetime]::UtcNow
    try {
        if (-not $process.Start()) { throw 'The LOT wrapper refresh process did not start.' }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        while (-not $process.WaitForExit(1000)) {
            $elapsedSeconds = [math]::Max(0, [math]::Floor(([datetime]::UtcNow - $startedUtc).TotalSeconds))
            Publish-AutomaticLotProgress -Stage 'Refreshing LOT command wrappers...' -Detail ("Generating operational launch commands; elapsed: {0} s." -f $elapsedSeconds)
        }
        $process.WaitForExit()
        $standardOutput = [string]$stdoutTask.Result
        $standardError = [string]$stderrTask.Result
        $outputLines = @(
            @($standardOutput -split '\r?\n')
            @($standardError -split '\r?\n')
        )
        return [pscustomobject]@{
            ExitCode = [int]$process.ExitCode
            Output = @($outputLines | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        }
    }
    finally {
        $process.Dispose()
    }
}

function Get-AutomaticLotToolkitRoot {
    param([AllowNull()][string]$RequestedRoot)

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        return [System.IO.Path]::GetFullPath($RequestedRoot)
    }

    $scriptsRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    return Split-Path -Parent $scriptsRoot
}

function Get-AutomaticLotSafeName {
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
        $Name = 'LOT-AUTO-W10{0}-{1}' -f $prefixSegment, (Get-Date -Format 'yyyyMMdd-HHmmss')
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

function Get-AutomaticLotNamePrefixes {
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

function Test-AutomaticLotNamePrefix {
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

function Get-AutomaticLotNameContainsValues {
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

function Test-AutomaticLotNameContains {
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
if ($ExcludeIntunePresent -and 'Intune' -notin $availableSources) {
    throw 'ExcludeIntunePresent requires an available Intune inventory source.'
}
if ($ExcludeStaleAd -and 'AD' -notin $availableSources) {
    throw 'ExcludeStaleAd requires an available AD inventory source.'
}

Publish-AutomaticLotProgress -Stage 'Loading inventory records...' -Detail ("Available sources: {0}" -f ($availableSources.ToArray() -join ' + '))
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
$namePrefixes = @(Get-AutomaticLotNamePrefixes -Prefixes $ComputerNamePrefix)
$namePrefixDisplay = $namePrefixes -join ';'
$nameContainsValues = @(Get-AutomaticLotNameContainsValues -ContainsValues $ComputerNameContains)
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
    if (-not (Test-AutomaticLotNamePrefix -ComputerKey $key -Prefixes $namePrefixes)) {
        $filterTypes.Add('ComputerNamePrefix')
        $filterValues.Add($namePrefixDisplay)
        $filterReasons.Add('COMPUTER_NAME_PREFIX_NOT_MATCHED')
    }
    if (-not (Test-AutomaticLotNameContains -ComputerKey $key -ContainsValues $nameContainsValues)) {
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
        ExclusionReason    = Get-AutomaticLotExclusionReason -AdRecords $keyAdRecords -IntuneRecords $keyIntuneRecords -AdCollision $adCollision
        ADFQDNs             = $adFqdns -join ';'
        ADOperatingSystem   = if ($keyAdRecords.Count -gt 0) { [string]$keyAdRecords[0].OperatingSystem } else { '' }
        IntuneOperatingSystem = if ($null -ne $intuneRecord) { [string]$intuneRecord.OperatingSystem } else { '' }
        IntuneOSVersion     = if ($null -ne $intuneRecord) { [string]$intuneRecord.OSVersion } else { '' }
        IntuneManagementState = if ($null -ne $intuneRecord) { [string]$intuneRecord.ManagementState } else { '' }
    })
}

Publish-AutomaticLotProgress -Stage 'Applying selection and safety filters...' -Detail ("Selected: {0}; safety exclusions: {1}; filtered out: {2}" -f $selected.Count,$excluded.Count,$filterExcluded.Count)
$safeLotName = Get-AutomaticLotSafeName -Name $LotName -NamePrefixes $namePrefixes
$lotPath = ''
$computersPath = ''
if ($Create) {
    if ($selected.Count -eq 0) { throw 'No eligible Windows 10 device was selected; the LOT was not created.' }

    Publish-AutomaticLotProgress -Stage 'Creating the automatic LOT folder...' -Detail ("LOT: {0}; selected devices: {1}" -f $safeLotName,$selected.Count)
    $lotsRoot = Join-Path $effectiveToolkitRoot 'Lots'
    New-Item -ItemType Directory -Path $lotsRoot -Force | Out-Null
    $lotPath = Join-Path $lotsRoot $safeLotName
    if (Test-Path -LiteralPath $lotPath) { throw "LOT folder already exists: $lotPath" }

    New-Item -ItemType Directory -Path $lotPath -Force | Out-Null
    $computersPath = Join-Path $lotPath 'Computers.txt'
    Publish-AutomaticLotProgress -Stage 'Writing Computers.txt...' -Detail ("{0} selected device(s)" -f $selected.Count)
    @($selected | Sort-Object ComputerName | Select-Object -ExpandProperty ComputerName) |
        Set-Content -LiteralPath $computersPath -Encoding ASCII

    Publish-AutomaticLotProgress -Stage 'Creating the LOT configuration...' -Detail $lotPath
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
        Publish-AutomaticLotProgress -Stage 'Refreshing LOT command wrappers...' -Detail 'Generating the launch commands for operational LOT folders. This step can take some time.'
        $wrapperRefresh = Join-Path $effectiveToolkitRoot 'Scripts\SmartM365-Windows11Upgrade-Update-LotCmdWrappers.ps1'
        if (-not (Test-Path -LiteralPath $wrapperRefresh -PathType Leaf)) {
            throw "Wrapper refresh script not found: $wrapperRefresh"
        }
        $wrapperRefreshResult = Invoke-AutomaticLotWrapperRefresh -ScriptPath $wrapperRefresh -EffectiveToolkitRoot $effectiveToolkitRoot
        $wrapperRefreshOutput = @($wrapperRefreshResult.Output)
        $wrapperRefreshExitCode = [int]$wrapperRefreshResult.ExitCode
        if ($wrapperRefreshExitCode -ne 0) {
            $wrapperRefreshDetail = @(
                $wrapperRefreshOutput |
                    ForEach-Object { ([string]$_).Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Select-Object -Last 10
            ) -join ' | '
            if ([string]::IsNullOrWhiteSpace($wrapperRefreshDetail)) {
                throw "LOT wrapper refresh failed with exit code $wrapperRefreshExitCode."
            }
            throw "LOT wrapper refresh failed with exit code $wrapperRefreshExitCode. Detail: $wrapperRefreshDetail"
        }
    }
}

$evidencePath = ''
if (-not $NoEvidence) {
    if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
        $EvidenceRoot = Join-Path $effectiveToolkitRoot 'Runs\AutomaticLotInventory'
    }
    Publish-AutomaticLotProgress -Stage 'Preparing the evidence folder...' -Detail $EvidenceRoot
    $evidencePath = Join-Path $EvidenceRoot (Get-Date -Format 'yyyyMMdd-HHmmssfff')
    New-Item -ItemType Directory -Path $evidencePath -Force | Out-Null

    Publish-AutomaticLotProgress -Stage 'Copying inventory evidence...' -Detail ("Sources: {0}" -f ($availableSources.ToArray() -join ' + '))
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
    Publish-AutomaticLotProgress -Stage 'Writing selection reports...' -Detail $evidencePath
    Export-AutomaticLotCsv -Path (Join-Path $evidencePath 'AutomaticLotSelection.csv') -Columns $selectionColumns -Rows $selected.ToArray()
    Export-AutomaticLotCsv -Path (Join-Path $evidencePath 'AutomaticLotExclusions.csv') -Columns $exclusionColumns -Rows $excluded.ToArray()
    Export-AutomaticLotCsv -Path (Join-Path $evidencePath 'AutomaticLotFilterExclusions.csv') -Columns $filterExclusionColumns -Rows $filterExcluded.ToArray()
}

$summary = [pscustomobject]@{
    Version                    = $script:AutomaticLotVersion
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
    LotName                    = $safeLotName
    LotPath                    = $lotPath
    ComputersPath              = $computersPath
    EvidencePath               = $evidencePath
}

if (-not $NoEvidence) {
    Publish-AutomaticLotProgress -Stage 'Finalizing automatic LOT evidence...' -Detail $evidencePath
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $evidencePath 'AutomaticLotSummary.json') -Encoding UTF8
}

Publish-AutomaticLotProgress -Stage $(if ($Create) { 'Automatic LOT created.' } else { 'Automatic LOT preview ready.' }) -Detail $(if ($Create) { $lotPath } else { "Selected devices: $($selected.Count)" })

[pscustomobject]@{
    Summary               = $summary
    SelectedDevices       = $selected.ToArray()
    ExcludedDevices       = $excluded.ToArray()
    FilterExcludedDevices = $filterExcluded.ToArray()
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCK40sWkisvdNcC
# rF0wRdJLNDHY4VIZF9wsMpuEoZtGHaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOWnGyhyyHsXAx3QeRCbYlxLnE8Q00yL1yuJu9xhjdMFMA0GCSqG
# SIb3DQEBAQUABIIBgHgCI9is4tFcAgRasOvImZVbz13xJ1vz/m2rEuBQpITrsjPQ
# LD4QeY2gh7xDwQl4dP5nDUt43WKcfel3huf0PcabkCay3eNj8PSkDdainP+U1twb
# THBq9Pg2bSFGTjyX2bfKufqPnJVzVyKrGhk4VguZQWTWST7nOZdUnJZcL53tfHZF
# k+YhmJRtuE1G3i4elINMQ5sdftyFRKNj8Mg7oDqRDobaMiPkuQNlCgr1HeTVAFrQ
# 3yJ0HQIMpp/hcHS9xm/TcpL+7fxAdcD15G7wfxx6ow2skgqSmCsr7CFfqJOe7Rbi
# 4mF3Hh8/6sEHVjvTdFfxdJi7iGGwZeEDG9oc6r217uzXynLm4Ej0DTsvDYC4QkHR
# pSXO3YDiw84RuJGUOxjHnSzV9jKK8qdmvsnd3DkMJgPKA0896ehgBJxGR3U4+99C
# NMHt8at+/B/d5t+65F82vmUDHzHqBdH4r5nqGSNCq7/I6MBZfKN1T2eJh1mUsn+3
# wGUC2ewQ55GCdZ6FYqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjcxMzU5
# MzVaMC8GCSqGSIb3DQEJBDEiBCCP6xOfpiMhoiUMWRI3yiOgzPqW+DfG+4LQo3kV
# HVzgvjANBgkqhkiG9w0BAQEFAASCAgAGppQRK5XPxnGkzdIkKzm2Zqf+vbocKUjM
# oUcXI9I29eP/s5ujGRrrDZ+R70L3554AmDFtfqPtqHYQQ2VRn/iK5Q04/LPcvt48
# ApSbtHx2idjgNLURHcGRShZu0kAni6m2tNfr9p4nTHLqtFCIGKftY6wE5rSYGJZT
# wdR1C0Yxn1w57EG3HwRnZ1pOYT44CFstDC+sfaQT99xid5oZw/5gqKi9IcbmMU4q
# nHwzQmaSIx2MiHRhD3R1k11HTs8w4jb+0fn6IOd5cd60RVryN7TXJRxJ7+kK5CCV
# OltF0pdXfoOkzey6lXDho6mrNuHp29VLEO1i9asarQfNenDn6T2J2hoN7UsfHLGX
# 4i9MEkFqdUZCVmrLUQYtnXRxaPZehTBT57zNzYCL6Epfansh68Bnk5qDj2jI5MaO
# eXL8HAiw06VIe1BmK0abu1Jmi+PIIfdsVlD6lAC9II3ZWj3/w8zaa+jNckorA37z
# OxBIi1m98GT18vIxOa0DV0frs6tmIvFPAkYysuj0W5f4Y4S1CdXWHItG3fMyovhG
# KIUeiHIpcNGNC5byWIgayl0fY3TMuuupcz/cSKlnNUK21hT2Vcbgu7P/WUOz8Sht
# yeLBiH+4ZL68mAbvqCYcNJQ8vXYZSFaW9KKFSVhEvoogFTRIR94xuBTPTi20qCjV
# I9DvaCWHTg==
# SIG # End signature block
