# Device readiness gap facts (docs/HYBRID-IDENTITY-DEVICE-READINESS-FINDINGS.md): raw facts only,
# no verdict. Thresholds and applicability are evaluated downstream.
Set-StrictMode -Version Latest

function Get-GapText([AllowNull()][object]$Row, [string[]]$Names) {
    if ($null -eq $Row) { return '' }
    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return ([string]$property.Value).Trim() }
    }
    return ''
}

function Get-GapKey([AllowNull()][object]$Value) { ([string]$Value).Trim().Trim('{', '}').ToLowerInvariant() }

function New-GapIndex([object[]]$Rows, [string[]]$Names) {
    # First row wins, as in the legacy producer.
    $index = @{}
    foreach ($row in $Rows) { $key = Get-GapKey (Get-GapText $row $Names); if ($key -and -not $index.ContainsKey($key)) { $index[$key] = $row } }
    return $index
}

function ConvertTo-GapNumber([AllowNull()][object]$Value) {
    $text = ([string]$Value).Trim() -replace ',', '.'
    $number = 0.0
    if ($text -and [double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and -not [double]::IsNaN($number) -and -not [double]::IsInfinity($number)) { return $number }
    return $null
}

function ConvertTo-RebootFact {
    # Offset-less source times are never converted to UTC: the age uses whole calendar days
    # between the source date and the reference UTC date, and says so.
    param([string]$Value, [Parameter(Mandatory)][datetimeoffset]$Reference)
    $result = [ordered]@{Utc = ''; Basis = ''; Days = $null; AgeBasis = '' }
    if (-not $Value) { $result.Basis = 'Missing'; return [pscustomobject]$result }
    $referenceDate = $Reference.UtcDateTime.Date
    if ($Value -match '(Z|[+-]\d\d:\d\d)$') {
        $dto = [datetimeoffset]::MinValue
        if ([datetimeoffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$dto)) {
            $result.Basis = 'Explicit offset'
            $result.Utc = $dto.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
            if ($dto -gt $Reference) { $result.AgeBasis = 'Future' } else { $result.Days = ($referenceDate - $dto.UtcDateTime.Date).Days; $result.AgeBasis = 'UTC calendar days' }
            return [pscustomobject]$result
        }
    } else {
        $local = [datetime]::MinValue
        $formats = [string[]]@('yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm:ss.FFFFFFF', 'dd/MM/yyyy HH:mm:ss', 'yyyy-MM-dd')
        if ([datetime]::TryParseExact($Value, $formats, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$local)) {
            $result.Basis = 'Offset-less source time'
            # A local date can lead UTC by up to one calendar day; anything later is in the future.
            if ($local.Date -gt $referenceDate.AddDays(1)) { $result.AgeBasis = 'Future' }
            else { $result.Days = [math]::Max(0, ($referenceDate - $local.Date).Days); $result.AgeBasis = 'Source calendar days (offset unknown, +/-1 day)' }
            return [pscustomobject]$result
        }
    }
    $result.Basis = 'Invalid'
    return [pscustomobject]$result
}

function Get-DeviceReadinessGapColumns {
    @('TenantKey', 'Inventory Scope', 'Device Source ID', 'Device Name', 'Azure AD Device ID', 'AD Object GUID',
      'Entra Registration Pending', 'Physical Memory (GB)', 'Firmware Type (source)', 'Secure Boot Status (source)',
      'Last Reboot (source)', 'Last Reboot Source', 'Last Reboot Time Basis', 'Last Reboot UTC',
      'Days Since Last Reboot', 'Reboot Age Basis', 'Evidence Date')
}

function New-DeviceReadinessGapEvidence {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$IntuneRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$EntraRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AdRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$LocalSystemRows,
        [Parameter(Mandatory)][datetimeoffset]$Reference
    )
    $entraById = New-GapIndex $EntraRows @('DeviceId')
    $adByIntuneId = New-GapIndex $AdRows @('IntuneDeviceId')
    $adByEntraId = New-GapIndex $AdRows @('ObjectGuidNormalized', 'EntraDeviceIdNormalized', 'AzureADDeviceId')
    $localByEntraId = New-GapIndex $LocalSystemRows @('AzureADDeviceId', 'DeviceId')
    $evidenceDate = $Reference.UtcDateTime.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    $matchedAd = [Collections.Generic.HashSet[object]]::new()
    $rows = [Collections.Generic.List[object]]::new()
    foreach ($device in $IntuneRows) {
        $deviceId = Get-GapKey (Get-GapText $device @('Device ID'))
        $entraId = Get-GapKey (Get-GapText $device @('Azure AD Device ID', 'Entra DeviceId'))
        $entraRow = if ($entraId -and $entraById.ContainsKey($entraId)) { $entraById[$entraId] } else { $null }
        $adRow = if ($deviceId -and $adByIntuneId.ContainsKey($deviceId)) { $adByIntuneId[$deviceId] } elseif ($entraId -and $adByEntraId.ContainsKey($entraId)) { $adByEntraId[$entraId] } else { $null }
        $localRow = if ($entraId -and $localByEntraId.ContainsKey($entraId)) { $localByEntraId[$entraId] } else { $null }
        if ($null -ne $adRow) { [void]$matchedAd.Add($adRow) }
        $rows.Add((New-GapRow 'Intune device' $device $entraId $entraRow $adRow $localRow $Reference $evidenceDate))
    }
    # AD computers without an Intune device keep their directory facts (hybrid join pending, reboot).
    foreach ($adRow in $AdRows) {
        if ($matchedAd.Contains($adRow)) { continue }
        $entraId = Get-GapKey (Get-GapText $adRow @('ObjectGuidNormalized', 'EntraDeviceIdNormalized', 'AzureADDeviceId'))
        $entraRow = if ($entraId -and $entraById.ContainsKey($entraId)) { $entraById[$entraId] } else { $null }
        $localRow = if ($entraId -and $localByEntraId.ContainsKey($entraId)) { $localByEntraId[$entraId] } else { $null }
        $rows.Add((New-GapRow 'AD computer only' $null $entraId $entraRow $adRow $localRow $Reference $evidenceDate))
    }
    return $rows.ToArray()
}

function New-GapRow {
    param([string]$Scope, [AllowNull()][object]$Device, [string]$EntraId, [AllowNull()][object]$EntraRow, [AllowNull()][object]$AdRow, [AllowNull()][object]$LocalRow, [datetimeoffset]$Reference, [string]$EvidenceDate)
    $entraRow = $EntraRow; $adRow = $AdRow; $localRow = $LocalRow; $entraId = $EntraId; $device = $Device
    # Legacy priority: the AD computer value, then the Intune local system value only when AD has none.
    $rebootRaw = Get-GapText $adRow @('LastRebootDateTime', 'LastRebootDate', 'LastBootUpTime')
    $rebootSource = if ($rebootRaw) { 'AD computer' } else { '' }
    if (-not $rebootRaw) { $rebootRaw = Get-GapText $localRow @('Last Reboot Date', 'LastRebootDate', 'LastBootUpTime'); if ($rebootRaw) { $rebootSource = 'Intune local system' } }
    $reboot = ConvertTo-RebootFact -Value $rebootRaw -Reference $Reference
    $pending = Get-GapText $entraRow @('IsPending')
    $memory = ConvertTo-GapNumber (Get-GapText $device @('PhysicalMemoryGB'))
    $tenant = Get-GapText $device @('TenantKey'); if (-not $tenant) { $tenant = Get-GapText $adRow @('TenantKey') }
    $name = Get-GapText $device @('Device name'); if (-not $name) { $name = (Get-GapText $adRow @('Name', 'SamAccountName')).TrimEnd('$') }
    [pscustomobject][ordered]@{
        'TenantKey' = $tenant
        'Inventory Scope' = $Scope
        'Device Source ID' = Get-GapText $device @('Device ID')
        'Device Name' = $name
        'Azure AD Device ID' = $entraId
        'AD Object GUID' = Get-GapKey (Get-GapText $adRow @('ObjectGuidNormalized', 'ObjectGUID'))
        'Entra Registration Pending' = if (-not $pending) { '' } elseif ($pending -match '^(?i:true|1|yes)$') { 'True' } elseif ($pending -match '^(?i:false|0|no)$') { 'False' } else { '' }
        'Physical Memory (GB)' = if ($null -ne $memory) { $memory.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) } else { '' }
        'Firmware Type (source)' = Get-GapText $localRow @('FirmwareType')
        'Secure Boot Status (source)' = Get-GapText $localRow @('SecureBootStatus')
        'Last Reboot (source)' = $rebootRaw
        'Last Reboot Source' = $rebootSource
        'Last Reboot Time Basis' = $reboot.Basis
        'Last Reboot UTC' = $reboot.Utc
        'Days Since Last Reboot' = if ($null -ne $reboot.Days) { [string]$reboot.Days } else { '' }
        'Reboot Age Basis' = $reboot.AgeBasis
        'Evidence Date' = $EvidenceDate
    }
}

Export-ModuleMember -Function New-DeviceReadinessGapEvidence, Get-DeviceReadinessGapColumns, ConvertTo-RebootFact
