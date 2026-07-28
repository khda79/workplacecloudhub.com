<#
.SYNOPSIS
Builds a guarded Intune Hybrid Join repair LOT from AD, Intune, and Entra inventories.

.DESCRIPTION
Selects enabled Windows 10 and Windows 11 client computers from Active Directory,
excludes devices already present in Intune and ambiguous inventory states, classifies
the remaining devices with Entra evidence, writes selection evidence under Runs, and
can create a standard operational LOT without launching it.

.VERSION
1.0.0
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$AdInventoryCsv,
    [Parameter(Mandatory = $true)][string]$IntuneInventoryCsv,
    [string]$EntraInventoryCsv,
    [string]$ToolkitRoot,
    [string]$LotName,
    [string[]]$ComputerNamePrefix,
    [string[]]$ComputerNameContains,
    [switch]$ExcludeStaleAd,
    [ValidateRange(1, 3650)][int]$AdLastLogonMaxAgeDays = 45,
    [scriptblock]$ProgressCallback,
    [string]$EvidenceRoot,
    [switch]$Create,
    [switch]$SkipWrapperRefresh,
    [switch]$NoEvidence
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:AutomaticLotVersion = '1.0.0'
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
        $segments = New-Object System.Collections.Generic.List[string]
        foreach ($prefix in @($NamePrefixes)) {
            $segment = [regex]::Replace(([string]$prefix).ToUpperInvariant(), '[^A-Z0-9_-]+', '-').Trim('-_')
            if (-not [string]::IsNullOrWhiteSpace($segment) -and -not $segments.Contains($segment)) {
                $segments.Add($segment)
            }
        }
        $prefixSegment = if ($segments.Count -gt 0) { "-$($segments -join '-')" } else { '' }
        $Name = 'LOT-AUTO-IHJ{0}-{1}' -f $prefixSegment, (Get-Date -Format 'yyyyMMdd-HHmmss')
    }

    $safeName = [regex]::Replace($Name.Trim(), '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        throw 'The automatic LOT name is empty after sanitization.'
    }
    if ($safeName -notmatch '^(?i)LOT-') {
        $safeName = "LOT-$safeName"
    }
    if ($safeName.Length -gt 96) {
        $safeName = $safeName.Substring(0, 96).TrimEnd('-._')
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
                throw "Invalid computer name prefix '$value'. Use only letters, digits, hyphens, or underscores."
            }
            if (-not $normalized.Contains($prefix)) { $normalized.Add($prefix) }
        }
    }
    return @($normalized.ToArray())
}

function Get-AutomaticLotNameContainsValues {
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

function Get-WindowsClientClassification {
    param(
        [AllowNull()][string]$OperatingSystem,
        [AllowNull()][string]$Version
    )

    $osText = ([string]$OperatingSystem).Trim()
    $versionText = ([string]$Version).Trim()
    if ($osText -match '(?i)Server') { return 'WindowsServer' }
    if (-not [string]::IsNullOrWhiteSpace($osText) -and $osText -notmatch '(?i)Windows') { return 'Unsupported' }
    if ($osText -match '(?i)Windows\s*11') { return 'Windows11' }
    if ($osText -match '(?i)Windows\s*10') { return 'Windows10' }

    $numbers = @([regex]::Matches($versionText, '\d+') | ForEach-Object { [int64]$_.Value })
    if ($numbers.Count -ge 3 -and $numbers[0] -eq 10 -and $numbers[1] -eq 0) {
        if ($numbers[2] -ge 22000) { return 'Windows11' }
        if ($numbers[2] -ge 10240) { return 'Windows10' }
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
            Classification     = Get-WindowsClientClassification -OperatingSystem $operatingSystem -Version $operatingSystemVersion
            OperatingSystem    = $operatingSystem
            OSVersion          = $operatingSystemVersion
            LastSeenUtc        = $lastSeen
        }
    }
}

function Get-IntuneAutomaticLotRecord {
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

function Get-EntraAutomaticLotRecord {
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

function Invoke-AutomaticLotWrapperRefresh {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$EffectiveToolkitRoot
    )

    $engine = (Get-Command pwsh.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Source)
    if ([string]::IsNullOrWhiteSpace($engine)) { $engine = 'powershell.exe' }
    $outputPath = Join-Path ([System.IO.Path]::GetTempPath()) ('SmartM365-IHJ-AutoWrapper-{0}.out' -f [guid]::NewGuid().ToString('N'))
    $errorPath = "$outputPath.err"
    try {
        $process = Start-Process -FilePath $engine -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath,
            '-RootPath', $EffectiveToolkitRoot
        ) -PassThru -WindowStyle Hidden -RedirectStandardOutput $outputPath -RedirectStandardError $errorPath
        $started = [datetime]::UtcNow
        while (-not $process.HasExited) {
            $elapsed = [math]::Floor(([datetime]::UtcNow - $started).TotalSeconds)
            Publish-AutomaticLotProgress -Stage 'Refreshing LOT command wrappers...' -Detail ("Generating operational launch commands; elapsed: {0} s." -f $elapsed)
            Start-Sleep -Milliseconds 500
            $process.Refresh()
        }
        $output = @()
        if (Test-Path -LiteralPath $outputPath) { $output += @(Get-Content -LiteralPath $outputPath -ErrorAction SilentlyContinue) }
        if (Test-Path -LiteralPath $errorPath) { $output += @(Get-Content -LiteralPath $errorPath -ErrorAction SilentlyContinue) }
        return [pscustomobject]@{ ExitCode = [int]$process.ExitCode; Output = $output }
    }
    finally {
        Remove-Item -LiteralPath $outputPath, $errorPath -Force -ErrorAction SilentlyContinue
    }
}

$effectiveToolkitRoot = Get-AutomaticLotToolkitRoot -RequestedRoot $ToolkitRoot
foreach ($requiredPath in @($AdInventoryCsv, $IntuneInventoryCsv)) {
    if ([string]::IsNullOrWhiteSpace($requiredPath) -or -not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Required automatic LOT inventory not found: $requiredPath"
    }
}
$entraAvailable = (-not [string]::IsNullOrWhiteSpace($EntraInventoryCsv) -and (Test-Path -LiteralPath $EntraInventoryCsv -PathType Leaf))

Publish-AutomaticLotProgress -Stage 'Loading inventory records...' -Detail $(if ($entraAvailable) { 'AD + Intune + Entra' } else { 'AD + Intune; Entra inventory unavailable' })
$adRows = @(Import-Csv -LiteralPath $AdInventoryCsv)
$intuneRows = @(Import-Csv -LiteralPath $IntuneInventoryCsv)
$entraRows = @()
if ($entraAvailable) { $entraRows = @(Import-Csv -LiteralPath $EntraInventoryCsv) }
$adRecords = @(Get-AdAutomaticLotRecord -Rows $adRows)
$intuneRecords = @(Get-IntuneAutomaticLotRecord -Rows $intuneRows)
$entraRecords = @(Get-EntraAutomaticLotRecord -Rows $entraRows)

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

$namePrefixes = @(Get-AutomaticLotNamePrefixes -Prefixes $ComputerNamePrefix)
$nameContainsValues = @(Get-AutomaticLotNameContainsValues -ContainsValues $ComputerNameContains)
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

Publish-AutomaticLotProgress -Stage 'Applying selection and safety filters...' -Detail ("Selected: {0}; safety exclusions: {1}; filtered out: {2}" -f $selected.Count, $excluded.Count, $filterExcluded.Count)
$safeLotName = Get-AutomaticLotSafeName -Name $LotName -NamePrefixes $namePrefixes
$lotPath = ''
$computersPath = ''
if ($Create) {
    if ($selected.Count -eq 0) { throw 'No eligible Hybrid Join device was selected; the LOT was not created.' }
    $lotsRoot = Join-Path $effectiveToolkitRoot 'Lots'
    New-Item -ItemType Directory -Path $lotsRoot -Force | Out-Null
    $lotPath = Join-Path $lotsRoot $safeLotName
    if (Test-Path -LiteralPath $lotPath) { throw "LOT folder already exists: $lotPath" }
    Publish-AutomaticLotProgress -Stage 'Creating the automatic LOT folder...' -Detail $lotPath
    New-Item -ItemType Directory -Path $lotPath -Force | Out-Null
    $computersPath = Join-Path $lotPath 'Computers.txt'
    @($selected | Sort-Object ComputerName | Select-Object -ExpandProperty ComputerName) | Set-Content -LiteralPath $computersPath -Encoding ASCII
    New-Item -ItemType File -Path (Join-Path $lotPath 'AdDomain.txt') -Force | Out-Null

    if (-not $SkipWrapperRefresh) {
        $wrapperScript = Join-Path $effectiveToolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1'
        if (-not (Test-Path -LiteralPath $wrapperScript -PathType Leaf)) { throw "Wrapper refresh script not found: $wrapperScript" }
        $wrapperResult = Invoke-AutomaticLotWrapperRefresh -ScriptPath $wrapperScript -EffectiveToolkitRoot $effectiveToolkitRoot
        if ($wrapperResult.ExitCode -ne 0) {
            $detail = @($wrapperResult.Output | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Last 10) -join ' | '
            throw "LOT wrapper refresh failed with exit code $($wrapperResult.ExitCode). Detail: $detail"
        }
    }
}

$evidencePath = ''
if (-not $NoEvidence) {
    if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
        $EvidenceRoot = Join-Path $effectiveToolkitRoot 'Runs\AutomaticLotInventory'
    }
    $evidencePath = Join-Path $EvidenceRoot (Get-Date -Format 'yyyyMMdd-HHmmssfff')
    Publish-AutomaticLotProgress -Stage 'Preparing automatic LOT evidence...' -Detail $evidencePath
    New-Item -ItemType Directory -Path $evidencePath -Force | Out-Null
    Copy-Item -LiteralPath $AdInventoryCsv -Destination (Join-Path $evidencePath 'DevicesAD.csv') -Force
    Copy-Item -LiteralPath $IntuneInventoryCsv -Destination (Join-Path $evidencePath 'DevicesIntune.csv') -Force
    if ($entraAvailable) { Copy-Item -LiteralPath $EntraInventoryCsv -Destination (Join-Path $evidencePath 'DevicesEntra.csv') -Force }
    Export-AutomaticLotCsv -Path (Join-Path $evidencePath 'AutomaticLotSelection.csv') -Columns @(
        'ComputerKey','ComputerName','SelectionReason','ADOperatingSystem','ADOSVersion','ADEnabled','ADDomain',
        'ADLastLogonTimestampUtc','EntraInventoryAvailable','EntraRegisteredState','EntraTrustType','EntraDeviceId'
    ) -Rows $selected.ToArray()
    Export-AutomaticLotCsv -Path (Join-Path $evidencePath 'AutomaticLotExclusions.csv') -Columns @(
        'ComputerKey','ComputerName','ExclusionReason','ADFQDNs','ADOperatingSystem','ADEnabled','IntuneMatches','EntraServerAdMatches'
    ) -Rows $excluded.ToArray()
    Export-AutomaticLotCsv -Path (Join-Path $evidencePath 'AutomaticLotFilterExclusions.csv') -Columns @(
        'ComputerKey','ComputerName','FilterType','FilterValue','FilterReason','ADLastLogonTimestampUtc'
    ) -Rows $filterExcluded.ToArray()
}

$summary = [pscustomobject]@{
    Version = $script:AutomaticLotVersion
    AvailableSources = if ($entraAvailable) { 'AD+Intune+Entra' } else { 'AD+Intune' }
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
    LotName = $safeLotName
    LotPath = $lotPath
    ComputersPath = $computersPath
    EvidencePath = $evidencePath
}

if (-not $NoEvidence) {
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $evidencePath 'AutomaticLotSummary.json') -Encoding UTF8
}
Publish-AutomaticLotProgress -Stage $(if ($Create) { 'Automatic LOT created.' } else { 'Automatic LOT preview ready.' }) -Detail $(if ($Create) { $lotPath } else { "Selected devices: $($selected.Count)" })

[pscustomobject]@{
    Summary = $summary
    SelectedDevices = $selected.ToArray()
    ExcludedDevices = $excluded.ToArray()
    FilterExcludedDevices = $filterExcluded.ToArray()
}
