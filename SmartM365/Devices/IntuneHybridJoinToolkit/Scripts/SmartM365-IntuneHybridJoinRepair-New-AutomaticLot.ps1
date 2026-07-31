<#
.SYNOPSIS
Builds a guarded Intune Hybrid Join repair LOT from AD, Intune, and Entra inventories.

.DESCRIPTION
Selects enabled Windows 10 and Windows 11 client computers from Active Directory,
excludes devices already present in Intune and ambiguous inventory states, classifies
the remaining devices with Entra evidence, writes selection evidence under Runs, and
can create a standard operational LOT without launching it.

.VERSION
1.0.1
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
$script:AutomaticLotVersion = '1.0.1'
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
        [AllowNull()][string]$OperatingSystem
    )

    $osText = ([string]$OperatingSystem).Trim()
    if ($osText -match '(?i)Server') { return 'WindowsServer' }
    if (-not [string]::IsNullOrWhiteSpace($osText) -and $osText -notmatch '(?i)Windows') { return 'Unsupported' }
    if ($osText -match '(?i)Windows\s*11\b') { return 'Windows11' }
    if ($osText -match '(?i)Windows\s*10\b') { return 'Windows10' }
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
            Classification     = Get-WindowsClientClassification -OperatingSystem $operatingSystem
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBt5KHddZQPJPpe
# 6rOIpzqKAYa8IDW/Sijh/cab2/0vCqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPFVzN9zGXe4o44pAo0j0WjqDcHFXMEm8458J1uMGZzBMA0GCSqG
# SIb3DQEBAQUABIIBgGzWkGSNSBLiQKiJytOo48w9tfrGF9lzQAQu0k5JPz9SNw2H
# pDhNGg02U9J1rPNCuzYkNGnDx/6aKYNTvXpDL3JSO1mkBFoUGwiYw96/ezul6pcl
# TRJ/Op6CS0ufkSlsZQMJ72o+JDoMxO8ru9Aa70paN7l15yUJNOiN9Py8bPhLWu8h
# 2erbzDP2/OGUJ5KS82rRULMNt1jtjxQkcX4GhOpslQR5d/3FdK/afWPBOa1kIXDt
# uMrwGpaxW1qdA/oR2ePJ1Sjzkn/C0T98BGDg04tYtc5VIYGZAsNUHThhqXzH2XLM
# HPLceBqFAybHZ3L1SiUM90GqgcScLqch9ZT2gNDVRJJbAb0KYhBhJtsM4VUDGUn5
# QV2WgiKmkWYxdkPRgaWjBmbw+2rByyKoioeaPaiOClrYwa5m1enbGMAlGqUUKCPd
# xdnigWXfK8N6L8d6g3N9fWIFU2rzZuC7D0eJGS1vWDxo7edO+hSeTEF2SVrN3jdF
# DZw3Y+ovPKFb0wB7JKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MzEwNjU4
# NTJaMC8GCSqGSIb3DQEJBDEiBCBq3zoBzt6FQ65yamFU8sH5dc88zytFymNjWWT7
# UzsIOzANBgkqhkiG9w0BAQEFAASCAgBwCyMbB7e4Lkb59COHtRWNKIp16jSTQvst
# xkYWUqQj3OcKVuHGd4viuYE0mU+UK2jVA7V9OxjQgzFrqFdBFqXRrwjAyCRgUlPx
# 7FifkpS6CioG8aEBuLY8K6582IpHz8vBoPPzq+JEQe+D6WCUBC3pOqs2z9pmUZIz
# tvX1AHxXQO3jPx8yuL8UatvbfS1gslPwO7H52oByphjvPeojNGn2L5KA04TqPZXu
# hQBCRxxRfhv0V4OS2SxYkYm16wDvIJAQXbHlTuo7o2YRcH8nUghNDQZBeP27hKqa
# iPfAj3iByJWBIhVdDWbv0l5FqFRIjVgV/InffwVMO/yqraCSsybmzfTd5dDssT8L
# 4vLOx/MVi5qU96zdRspLWR2IAAaWwn1UAWeDofDuEExiU5RTiSem6karGgZipLeA
# ArAUkVWCPh29YnQsXr6rbZ3g2d1Xg3+FXMgnpzIARWj6Lln5+t5cjPUtQFPhmn/V
# keZCYFgU+u8Rnofk2+ciT/9TxUFmjq09OCiLO7SPWdOcxW59moZbIKyzbKxSw3Ss
# MfFqP1oSQ6TPPuxYi66p6GWT9oxsWsm6iS2njsn2gBkZM0nyimh6AB1vHHhIwlpi
# IBmqXbM6hFnA0FKRIbUEM8eEf90o0QE79fDUWq0EYJetUSp0OS3lDWGcQ8uWU3bD
# 4vwmyurw4g==
# SIG # End signature block
