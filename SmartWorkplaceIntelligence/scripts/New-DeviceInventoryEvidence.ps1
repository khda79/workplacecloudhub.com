[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\DeviceInventoryEvidence.csv'),

    [string]$SignalsOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\DeviceRiskEvidenceSummary.csv'),

    [string]$DirectorySummaryOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\DeviceDirectorySummary.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'EndpointAnalyticsEvidence.psm1') -Force -ErrorAction Stop

function Get-RequiredCsv {
    param([string]$RelativePath)

    $path = Join-Path $DataRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required source CSV was not found: $path"
    }
    return $path
}

function Get-NormalizedKey {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Convert-ToDateTimeOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

    $parsed = [datetime]::MinValue
    $culture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
    if ([datetime]::TryParse([string]$Value, $culture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
        return $parsed
    }
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Convert-ToDoubleOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

    $parsed = 0.0
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Convert-ToInvariantDecimalText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    return ([double]$Value).ToString('0.0', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-UniqueIndex {
    param(
        [object[]]$Rows,
        [string]$PropertyName
    )

    $groups = @{}
    foreach ($row in $Rows) {
        $key = Get-NormalizedKey $row.$PropertyName
        if (-not $key) { continue }
        if (-not $groups.ContainsKey($key)) { $groups[$key] = [System.Collections.Generic.List[object]]::new() }
        $groups[$key].Add($row)
    }

    $index = @{}
    foreach ($entry in $groups.GetEnumerator()) {
        if ($entry.Value.Count -eq 1) { $index[$entry.Key] = $entry.Value[0] }
    }
    return $index
}

function Get-WindowsUpdatePriorityScore {
    param([AllowNull()][object]$Row)
    if ($null -eq $Row) { return -1 }

    $stateScore = switch (Get-NormalizedKey $Row.AggregateState) {
        'rollback' { 500 }
        'error' { 490 }
        'cancelled' { 400 }
        'inprogress' { 300 }
        'success' { 100 }
        default { 0 }
    }
    $priorityScore = switch (Get-NormalizedKey $Row.ActionPriority) {
        'high' { 30 }
        'medium' { 20 }
        'low' { 10 }
        default { 0 }
    }
    return $stateScore + $priorityScore
}

function Get-WindowsUpdateIndex {
    param([object[]]$Rows)

    $index = @{}
    foreach ($row in $Rows) {
        $key = Get-NormalizedKey $row.DeviceId
        if (-not $key) { continue }
        if (-not $index.ContainsKey($key) -or (Get-WindowsUpdatePriorityScore $row) -gt (Get-WindowsUpdatePriorityScore $index[$key])) {
            $index[$key] = $row
        }
    }
    return $index
}

function Get-PreferredLatestIndex {
    param(
        [object[]]$Rows,
        [string]$KeyProperty,
        [string]$PreferredProperty,
        [string]$DateProperty
    )

    Assert-EndpointAnalyticsDeviceGrain -Rows $Rows
    $index = @{}
    foreach ($row in $Rows) {
        $key = Get-NormalizedKey $row.$KeyProperty
        if (-not $key) { continue }
        if (-not $index.ContainsKey($key)) {
            $index[$key] = $row
            continue
        }

        $current = $index[$key]
        $rowHasPreferred = -not [string]::IsNullOrWhiteSpace([string]$row.$PreferredProperty)
        $currentHasPreferred = -not [string]::IsNullOrWhiteSpace([string]$current.$PreferredProperty)
        $rowDate = Convert-ToDateTimeOrNull $row.$DateProperty
        $currentDate = Convert-ToDateTimeOrNull $current.$DateProperty
        if (($rowHasPreferred -and -not $currentHasPreferred) -or
            ($rowHasPreferred -eq $currentHasPreferred -and $null -ne $rowDate -and ($null -eq $currentDate -or $rowDate -gt $currentDate))) {
            $index[$key] = $row
        }
    }
    return $index
}

function Get-NormalizedValueSet {
    param([object[]]$Values)

    $set = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($value in $Values) {
        $key = Get-NormalizedKey $value
        if ($key) { [void]$set.Add($key) }
    }
    return ,$set
}

function Get-Platform {
    param([string]$OperatingSystem)
    $os = Get-NormalizedKey $OperatingSystem
    if ($os -match 'windows') { return 'Windows' }
    if ($os -match 'mac|os x') { return 'macOS' }
    if ($os -match 'android') { return 'Android' }
    if ($os -match 'ios|iphone|ipad') { return 'iOS' }
    return 'Other'
}

function Get-DeviceCategory {
    param([string]$Platform)
    switch ($Platform) {
        'Windows' { return 'PC Windows' }
        'macOS' { return 'MAC' }
        'Android' { return 'Mobile Android/iOS' }
        'iOS' { return 'Mobile Android/iOS' }
        default { return 'Other' }
    }
}

function Get-WindowsRelease {
    param([string]$OperatingSystemVersion)
    if ([string]::IsNullOrWhiteSpace($OperatingSystemVersion)) { return 'Unknown' }
    $parts = $OperatingSystemVersion.Split('.')
    if ($parts.Count -lt 3) { return 'Unknown' }
    $build = 0
    if (-not [int]::TryParse($parts[2], [ref]$build)) { return 'Unknown' }
    if ($build -ge 28000) { return 'Windows 11 26H1' }
    if ($build -ge 26200) { return 'Windows 11 25H2' }
    if ($build -ge 26100) { return 'Windows 11 24H2' }
    if ($build -ge 22631) { return 'Windows 11 23H2' }
    if ($build -ge 22621) { return 'Windows 11 22H2' }
    if ($build -ge 22000) { return 'Windows 11 21H2' }
    if ($build -ge 19045) { return 'Windows 10 22H2' }
    return 'Other / legacy'
}

function Get-WindowsGeneration {
    param(
        [string]$Platform,
        [string]$OperatingSystemVersion
    )
    if ($Platform -ne 'Windows') { return 'Not applicable' }
    if ([string]::IsNullOrWhiteSpace($OperatingSystemVersion)) { return 'Unknown' }
    $parts = $OperatingSystemVersion.Split('.')
    if ($parts.Count -lt 3) { return 'Unknown' }
    $build = 0
    if (-not [int]::TryParse($parts[2], [ref]$build)) { return 'Unknown' }
    if ($build -ge 22000) { return 'Windows 11' }
    if ($build -ge 10240) { return 'Windows 10' }
    return 'Windows legacy'
}

function Get-Windows11UpgradeEligibility {
    param(
        [string]$WindowsGeneration,
        [AllowNull()][object]$UpgradeRow
    )
    if ($WindowsGeneration -eq 'Not applicable') { return 'Not applicable' }
    if ($WindowsGeneration -eq 'Windows 11') { return 'Already upgraded' }
    if ($null -eq $UpgradeRow) { return 'Not assessed' }
    switch (Get-NormalizedKey $UpgradeRow.UpgradeEligibility) {
        'capable' { return 'Capable' }
        'notcapable' { return 'Not capable' }
        'upgraded' { return 'Already upgraded' }
        default { return 'Unknown' }
    }
}

function Get-Windows11BlockingReasons {
    param([AllowNull()][object]$UpgradeRow)
    if ($null -eq $UpgradeRow -or (Get-NormalizedKey $UpgradeRow.UpgradeEligibility) -ne 'notcapable') { return $null }

    $checks = [ordered]@{
        RamCheckFailed = 'RAM'
        StorageCheckFailed = 'Storage'
        ProcessorCoreCountCheckFailed = 'Processor cores'
        ProcessorSpeedCheckFailed = 'Processor speed'
        TPMCheckFailed = 'TPM'
        SecureBootCheckFailed = 'Secure Boot'
        ProcessorFamilyCheckFailed = 'Processor family'
        Processor64BitCheckFailed = '64-bit processor'
        OSCheckFailed = 'Operating system'
    }
    $reasons = foreach ($entry in $checks.GetEnumerator()) {
        if ((Get-NormalizedKey $UpgradeRow.($entry.Key)) -in @('true', '1', 'yes')) { $entry.Value }
    }
    if (@($reasons).Count -eq 0) { return 'Not reported' }
    return ($reasons -join '; ')
}

function Get-ComplianceState {
    param([string]$Compliance)
    $value = Get-NormalizedKey $Compliance
    if ($value -eq 'compliant') { return 'Compliant' }
    if ($value -eq 'noncompliant') { return 'Noncompliant' }
    if ($value -match 'grace') { return 'In Grace Period' }
    return 'Unknown'
}

function Get-Ownership {
    param([string]$Ownership)
    $value = Get-NormalizedKey $Ownership
    if ($value -match 'company|corporate') { return 'Corporate' }
    if ($value -match 'personal') { return 'Personal' }
    return 'Unknown'
}

function Get-ManagementState {
    param([string]$ManagedBy)
    $value = Get-NormalizedKey $ManagedBy
    if (-not $value -or $value -match 'unknown|unmanaged|none') { return 'Unmanaged' }
    if ($value -match 'co') { return 'Co-managed' }
    return 'Managed'
}

$intunePath = Get-RequiredCsv 'DATA-LAST\Intune_Devices_Inventory.csv'
$usersPath = Get-RequiredCsv 'DATA-LAST\M365_Users_Active.csv'
$entraPath = Get-RequiredCsv 'DATA-LAST\M365_Entra_Devices.csv'
$adPath = Get-RequiredCsv 'DATA-LAST\AD_Computers_AllDomains.csv'
$upgradeEligibilityPath = Get-RequiredCsv 'DATA-LAST\Intune_Devices_UpgradeEligibility.csv'
$windowsUpdatePath = Get-RequiredCsv 'DATA-LAST\Intune_WindowsUpdate_Status.csv'
$endpointPerformancePath = Get-RequiredCsv 'DATA-LAST\Intune_EndpointAnalytics_DevicePerformance.csv'
$endpointStartupPath = Get-RequiredCsv 'DATA-LAST\Intune_EndpointAnalytics_StartupDevices.csv'
$hardwareConflictPath = Get-RequiredCsv 'DATA-LAST\M365_Entra_Devices_HardwareIdConflicts.csv'
$removalCandidatePath = Get-RequiredCsv 'DATA-LAST\M365_Entra_Devices_RemovalCandidates.csv'

$intune = @(Import-Csv -LiteralPath $intunePath)
$users = @(Import-Csv -LiteralPath $usersPath)
$entra = @(Import-Csv -LiteralPath $entraPath)
$ad = @(Import-Csv -LiteralPath $adPath)
$upgradeEligibility = @(Import-Csv -LiteralPath $upgradeEligibilityPath)
$windowsUpdate = @(Import-Csv -LiteralPath $windowsUpdatePath)
$endpointPerformance = @(Import-Csv -LiteralPath $endpointPerformancePath)
$endpointStartup = @(Import-Csv -LiteralPath $endpointStartupPath)
$hardwareConflicts = @(Import-Csv -LiteralPath $hardwareConflictPath)
$removalCandidates = @(Import-Csv -LiteralPath $removalCandidatePath)

$userIndex = Get-UniqueIndex -Rows $users -PropertyName 'User principal name'
$entraIndex = Get-UniqueIndex -Rows $entra -PropertyName 'DeviceId'
$adIndex = Get-UniqueIndex -Rows $ad -PropertyName 'IntuneDeviceId'
$upgradeEligibilityIdIndex = Get-UniqueIndex -Rows $upgradeEligibility -PropertyName 'GraphId'
$upgradeEligibilityNameIndex = Get-UniqueIndex -Rows $upgradeEligibility -PropertyName 'NormalizedDeviceName'
$windowsUpdateIndex = Get-WindowsUpdateIndex -Rows $windowsUpdate
$endpointPerformanceIndex = Get-PreferredLatestIndex -Rows $endpointPerformance -KeyProperty 'DeviceId' -PreferredProperty 'EndpointAnalyticsScore' -DateProperty 'ReportRefreshDate'
$endpointStartupIndex = Get-PreferredLatestIndex -Rows $endpointStartup -KeyProperty 'DeviceId' -PreferredProperty 'StopErrorCount' -DateProperty 'ReportRefreshDate'
$hardwareConflictDeviceIds = Get-NormalizedValueSet -Values @($hardwareConflicts | ForEach-Object { @([string]$_.DeviceIds -split '[,;|\s]+' | Where-Object { $_ }) })
$removalCandidateDeviceIds = Get-NormalizedValueSet -Values @($removalCandidates | ForEach-Object { $_.CandidateDeviceId })
$now = Get-Date

$evidence = @(foreach ($device in $intune) {
    $deviceId = Get-NormalizedKey $device.'Device ID'
    $deviceName = Get-NormalizedKey $device.'Device name'
    $entraId = Get-NormalizedKey $device.'Azure AD Device ID'
    if (-not $entraId) { $entraId = Get-NormalizedKey $device.'Entra DeviceId' }
    $upn = Get-NormalizedKey $device.'Primary user UPN'

    $entraRow = if ($entraId -and $entraIndex.ContainsKey($entraId)) { $entraIndex[$entraId] } else { $null }
    $adRow = if ($deviceId -and $adIndex.ContainsKey($deviceId)) { $adIndex[$deviceId] } else { $null }
    $userRow = if ($upn -and $userIndex.ContainsKey($upn)) { $userIndex[$upn] } else { $null }
    $upgradeRow = if ($deviceId -and $upgradeEligibilityIdIndex.ContainsKey($deviceId)) {
        $upgradeEligibilityIdIndex[$deviceId]
    } elseif ($deviceName -and $upgradeEligibilityNameIndex.ContainsKey($deviceName)) {
        $upgradeEligibilityNameIndex[$deviceName]
    } else {
        $null
    }
    $windowsUpdateRow = if ($deviceId -and $windowsUpdateIndex.ContainsKey($deviceId)) { $windowsUpdateIndex[$deviceId] } else { $null }
    $endpointPerformanceRow = if ($deviceId -and $endpointPerformanceIndex.ContainsKey($deviceId)) { $endpointPerformanceIndex[$deviceId] } else { $null }
    $endpointStartupRow = if ($deviceId -and $endpointStartupIndex.ContainsKey($deviceId)) { $endpointStartupIndex[$deviceId] } else { $null }

    $platform = Get-Platform $device.OS
    $windowsGeneration = Get-WindowsGeneration -Platform $platform -OperatingSystemVersion $device.'OS version'
    $windows11UpgradeEligibility = Get-Windows11UpgradeEligibility -WindowsGeneration $windowsGeneration -UpgradeRow $upgradeRow
    $windows11BlockingReasons = Get-Windows11BlockingReasons -UpgradeRow $upgradeRow
    $totalBytes = Convert-ToDoubleOrNull $device.'Total storage'
    $freeBytes = Convert-ToDoubleOrNull $device.'Free storage'
    $totalGiB = if ($null -ne $totalBytes) { [math]::Round($totalBytes / 1GB, 1) } else { $null }
    $freeGiB = if ($null -ne $freeBytes) { [math]::Round($freeBytes / 1GB, 1) } else { $null }
    $freePercent = if ($null -ne $totalBytes -and $totalBytes -gt 0 -and $null -ne $freeBytes) { [math]::Round(100 * $freeBytes / $totalBytes, 1) } else { $null }
    $diskState = if (($null -ne $freeGiB -and $freeGiB -lt 20) -or ($null -ne $freePercent -and $freePercent -lt 15)) {
        'Low'
    } elseif ($null -eq $freeGiB -or $null -eq $freePercent) {
        'Unknown'
    } elseif ($null -ne $freeGiB -and $freeGiB -lt 40) {
        'Approaching Low'
    } else {
        'Healthy'
    }

    $enrollmentDate = Convert-ToDateTimeOrNull $device.'Enrollment date'
    $entraRegistrationDate = if ($null -ne $entraRow) { Convert-ToDateTimeOrNull $entraRow.RegistrationDateTime } else { $null }
    $adCreationDate = if ($null -ne $adRow) { Convert-ToDateTimeOrNull $adRow.WhenCreated } else { $null }
    $observedDates = @(@($enrollmentDate, $entraRegistrationDate, $adCreationDate) | Where-Object { $null -ne $_ })
    $observedSince = if ($observedDates.Count -gt 0) { ($observedDates | Sort-Object | Select-Object -First 1) } else { $null }
    $tenureYears = if ($null -ne $observedSince) { [math]::Round(($now - $observedSince).TotalDays / 365.25, 1) } else { $null }
    $tenureBucket = if ($null -eq $tenureYears) { 'Unknown' } elseif ($tenureYears -lt 1) { '< 1 year' } elseif ($tenureYears -lt 2) { '1-2 years' } elseif ($tenureYears -lt 4) { '2-4 years' } else { '4+ years' }
    $tenureConfidence = switch ($observedDates.Count) { 3 { 'High' } 2 { 'Medium' } 1 { 'Low' } default { 'Unknown' } }
    $country = if ($null -ne $userRow -and -not [string]::IsNullOrWhiteSpace($userRow.'Usage location')) { $userRow.'Usage location'.Trim().ToUpperInvariant() } elseif ($null -ne $userRow -and -not [string]::IsNullOrWhiteSpace($userRow.CountryOrRegion)) { $userRow.CountryOrRegion.Trim() } else { 'Unknown' }
    $sourceScope = @('Intune')
    if ($null -ne $entraRow) { $sourceScope += 'Entra device' }
    if ($null -ne $adRow) { $sourceScope += 'AD computer' }
    if ($null -ne $userRow) { $sourceScope += 'Entra user' }
    if ($null -ne $upgradeRow) { $sourceScope += 'Windows 11 eligibility' }
    if ($null -ne $windowsUpdateRow) { $sourceScope += 'Windows Update status' }
    if ($null -ne $endpointPerformanceRow) { $sourceScope += 'Endpoint Analytics performance' }
    if ($null -ne $endpointStartupRow) { $sourceScope += 'Endpoint Analytics startup' }

    $appReliabilityScore = if ($null -ne $endpointPerformanceRow) { Convert-ToDoubleOrNull $endpointPerformanceRow.AppReliabilityScore } else { $null }
    if ($null -ne $appReliabilityScore -and $appReliabilityScore -lt 0) { $appReliabilityScore = $null }

    [pscustomobject][ordered]@{
        'Device Source ID' = $device.'Device ID'
        'Device Name' = $device.'Device name'
        'Primary User UPN' = $device.'Primary user UPN'
        'Country' = $country
        'Platform' = $platform
        'Device Category' = Get-DeviceCategory $platform
        'Ownership' = Get-Ownership $device.Ownership
        'Management State' = Get-ManagementState $device.'Managed by'
        'Compliance State' = Get-ComplianceState $device.Compliance
        'Manufacturer' = if ([string]::IsNullOrWhiteSpace($device.Manufacturer)) { 'Unknown' } else { $device.Manufacturer.Trim() }
        'Model' = if ([string]::IsNullOrWhiteSpace($device.Model)) { 'Unknown' } else { $device.Model.Trim() }
        'Operating System Family' = $platform
        'Operating System Version' = $device.'OS version'
        'Windows Release' = if ($platform -eq 'Windows') { Get-WindowsRelease $device.'OS version' } else { 'Not applicable' }
        'Windows Generation' = $windowsGeneration
        'Windows 11 Upgrade Eligibility' = $windows11UpgradeEligibility
        'Windows 11 Blocking Reasons' = $windows11BlockingReasons
        'System Drive Capacity (GB)' = Convert-ToInvariantDecimalText $totalGiB
        'Free Disk Space (GB)' = Convert-ToInvariantDecimalText $freeGiB
        'Free Disk Space (%)' = Convert-ToInvariantDecimalText $freePercent
        'Disk Space State' = $diskState
        'AD to Intune Match' = if ($null -ne $adRow) { 'Matched' } else { 'Unmatched' }
        'Entra Match' = if ($null -ne $entraRow) { 'Matched' } else { 'Unmatched' }
        'Entra Account State' = if ($null -eq $entraRow) { 'Not observed' } elseif ((Get-NormalizedKey $entraRow.AccountEnabled) -eq 'true') { 'Enabled' } else { 'Disabled' }
        'Intune Encryption State' = switch (Get-NormalizedKey $device.Encrypted) { 'true' { 'Encrypted' } 'false' { 'Not encrypted' } default { 'Not observed' } }
        'Enrollment DateTime' = if ($null -ne $enrollmentDate) { $enrollmentDate.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
        'Last Sync DateTime' = (Convert-ToDateTimeOrNull $device.LastSyncDateTime)?.ToString('yyyy-MM-dd HH:mm:ss')
        'Observed Since' = if ($null -ne $observedSince) { $observedSince.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
        'Managed Tenure (Years)' = Convert-ToInvariantDecimalText $tenureYears
        'Tenure Bucket' = $tenureBucket
        'Tenure Confidence' = $tenureConfidence
        'Windows Update State' = if ($null -ne $windowsUpdateRow -and -not [string]::IsNullOrWhiteSpace($windowsUpdateRow.AggregateState)) { $windowsUpdateRow.AggregateState.Trim() } else { 'Not observed' }
        'Windows Update Risk' = if ($null -ne $windowsUpdateRow -and -not [string]::IsNullOrWhiteSpace($windowsUpdateRow.RiskBucket)) { $windowsUpdateRow.RiskBucket.Trim() } else { 'Not observed' }
        'Windows Update Action Code' = if ($null -ne $windowsUpdateRow -and -not [string]::IsNullOrWhiteSpace($windowsUpdateRow.ActionCode)) { $windowsUpdateRow.ActionCode.Trim() } else { 'None' }
        'Secure Boot State' = if ($null -ne $adRow -and -not [string]::IsNullOrWhiteSpace($adRow.SecureBootStatus)) { $adRow.SecureBootStatus.Trim() } else { 'Not observed' }
        'Endpoint Analytics Score' = if ($null -ne $endpointPerformanceRow) { Convert-ToInvariantDecimalText (Convert-ToDoubleOrNull $endpointPerformanceRow.EndpointAnalyticsScore) } else { $null }
        'Endpoint Analytics State' = if ($null -eq $endpointPerformanceRow -or [string]::IsNullOrWhiteSpace($endpointPerformanceRow.EndpointAnalyticsScore)) { 'Not observed' } elseif ([double]$endpointPerformanceRow.EndpointAnalyticsScore -lt 50) { 'Below 50' } else { '50 or above' }
        'Startup Score' = if ($null -ne $endpointStartupRow) { Convert-ToInvariantDecimalText (Convert-ToDoubleOrNull $endpointStartupRow.StartupScore) } else { $null }
        'Startup Score State' = if ($null -eq $endpointStartupRow -or [string]::IsNullOrWhiteSpace($endpointStartupRow.StartupScore)) { 'Not observed' } elseif ([double]$endpointStartupRow.StartupScore -lt 50) { 'Below 50' } else { '50 or above' }
        'App Reliability Score' = Convert-ToInvariantDecimalText $appReliabilityScore
        'App Reliability State' = if ($null -eq $appReliabilityScore) { 'Not observed' } elseif ($appReliabilityScore -lt 50) { 'Below 50' } else { '50 or above' }
        'Boot Score' = if ($null -ne $endpointStartupRow) { Convert-ToInvariantDecimalText (Convert-ToDoubleOrNull $endpointStartupRow.BootScore) } else { $null }
        'Sign-in Score' = if ($null -ne $endpointStartupRow) { Convert-ToInvariantDecimalText (Convert-ToDoubleOrNull $endpointStartupRow.SignInScore) } else { $null }
        'Core Boot Time (s)' = if ($null -ne $endpointStartupRow) { Convert-ToInvariantDecimalText (Convert-ToDoubleOrNull $endpointStartupRow.CoreBootTime) } else { $null }
        'Core Sign-in Time (s)' = if ($null -ne $endpointStartupRow) { Convert-ToInvariantDecimalText (Convert-ToDoubleOrNull $endpointStartupRow.CoreSignInTime) } else { $null }
        'Restart Count' = if ($null -ne $endpointStartupRow -and -not [string]::IsNullOrWhiteSpace($endpointStartupRow.RestartCount)) { [int][double]$endpointStartupRow.RestartCount } else { $null }
        'Stop Error Count' = if ($null -ne $endpointStartupRow -and -not [string]::IsNullOrWhiteSpace($endpointStartupRow.StopErrorCount)) { [int][double]$endpointStartupRow.StopErrorCount } else { $null }
        'Stop Error State' = if ($null -eq $endpointStartupRow -or [string]::IsNullOrWhiteSpace($endpointStartupRow.StopErrorCount)) { 'Not observed' } elseif ([double]$endpointStartupRow.StopErrorCount -gt 0) { 'Stop errors observed' } else { 'No stop errors' }
        'Crash Count' = if ($null -ne $endpointPerformanceRow -and -not [string]::IsNullOrWhiteSpace($endpointPerformanceRow.CrashCount)) { [int][double]$endpointPerformanceRow.CrashCount } else { $null }
        'Crash State' = if ($null -eq $endpointPerformanceRow -or [string]::IsNullOrWhiteSpace($endpointPerformanceRow.CrashCount)) { 'Not observed' } elseif ([double]$endpointPerformanceRow.CrashCount -gt 0) { 'Crashes observed' } else { 'No crashes' }
        'Entra Hardware ID Conflict' = if ($entraId -and $hardwareConflictDeviceIds.Contains($entraId)) { 'Yes' } else { 'No' }
        'Removal Candidate' = if ($entraId -and $removalCandidateDeviceIds.Contains($entraId)) { 'Yes' } else { 'No' }
        'Windows 11 Release State' = if ($windowsGeneration -ne 'Windows 11') { 'Not applicable' } elseif ((Get-WindowsRelease $device.'OS version') -in @('Windows 11 21H2', 'Windows 11 22H2', 'Windows 11 23H2')) { 'Older release' } else { 'Current release' }
        'Source Scope' = ($sourceScope -join ' + ')
    }
})

$outputDirectory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$evidence | Sort-Object 'Device Source ID' | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8NoBOM

$enabledAdWindowsPcs = @($ad | Where-Object {
    (Get-NormalizedKey $_.Enabled) -eq 'true' -and
    ([string]$_.OperatingSystem -match '(?i)windows') -and
    ([string]$_.OperatingSystem -notmatch '(?i)server')
})
$enabledAdWindowsPcsManaged = @($enabledAdWindowsPcs | Where-Object {
    (Get-NormalizedKey $_.ExistsInIntune) -eq 'true' -or
    (Get-NormalizedKey $_.IsRegisteredInIntune) -eq 'true' -or
    -not [string]::IsNullOrWhiteSpace([string]$_.IntuneDeviceId)
})
# Match AD Windows 10 computers against the current Intune inventory. A unique
# device name can reconcile a re-enrolled computer; ambiguous names are excluded
# from the AD-only count rather than silently classified as unmanaged.
$intuneIds = Get-NormalizedValueSet -Values @($intune | ForEach-Object { $_.'Device ID' })
$intuneNameGroups = @{}
foreach ($device in $intune) {
    $nameKey = Get-NormalizedKey $device.'Device name'
    if ($nameKey) { $intuneNameGroups[$nameKey] = 1 + [int]$intuneNameGroups[$nameKey] }
}
$enabledAdWindows10 = @($ad | Where-Object {
    (Get-NormalizedKey $_.Enabled) -eq 'true' -and
    (Get-NormalizedKey $_.OperatingSystemShortName) -eq 'windows 10'
})
$adOnlyWindows10 = 0
$ambiguousWindows10Matches = 0
foreach ($computer in $enabledAdWindows10) {
    $idKey = Get-NormalizedKey $computer.IntuneDeviceId
    $nameKey = Get-NormalizedKey $computer.Name
    if ($idKey -and $intuneIds.Contains($idKey)) { continue }
    if ($nameKey -and $intuneNameGroups.ContainsKey($nameKey)) {
        if ($intuneNameGroups[$nameKey] -eq 1) { continue }
        $ambiguousWindows10Matches++
        continue
    }
    $adOnlyWindows10++
}
$intuneWindows10Undetermined = @($evidence | Where-Object {
    $_.'Windows Generation' -eq 'Windows 10' -and
    $_.'Windows 11 Upgrade Eligibility' -in @('Not assessed', 'Unknown')
}).Count
$disabledAdComputerObjects = @($ad | Where-Object { (Get-NormalizedKey $_.Enabled) -eq 'false' }).Count
$disabledEntraDevices = @($entra | Where-Object { (Get-NormalizedKey $_.AccountEnabled) -eq 'false' }).Count
$directorySummary = [pscustomobject][ordered]@{
    'Snapshot Date' = (Get-Item -LiteralPath $adPath).LastWriteTime.Date.ToString('yyyy-MM-dd')
    'Enabled AD Windows PCs' = $enabledAdWindowsPcs.Count
    'Enabled AD Windows PCs Managed by Intune' = $enabledAdWindowsPcsManaged.Count
    'Enabled AD Windows PCs Unmanaged' = $enabledAdWindowsPcs.Count - $enabledAdWindowsPcsManaged.Count
    'Enabled AD PCs Managed by Intune (%)' = if ($enabledAdWindowsPcs.Count -gt 0) { ($enabledAdWindowsPcsManaged.Count / [double]$enabledAdWindowsPcs.Count).ToString('0.######', [Globalization.CultureInfo]::InvariantCulture) } else { $null }
    'Disabled AD Computer Objects' = $disabledAdComputerObjects
    'Disabled Entra Devices' = $disabledEntraDevices
    'Enabled AD Windows 10 PCs' = $enabledAdWindows10.Count
    'Enabled AD Windows 10 PCs Without Intune' = $adOnlyWindows10
    'AD Windows 10 Ambiguous Intune Matches' = $ambiguousWindows10Matches
    'Intune Windows 10 Eligibility Undetermined' = $intuneWindows10Undetermined
    'Windows 10 Eligibility Undetermined' = $intuneWindows10Undetermined + $adOnlyWindows10
    'Country Definition' = 'Managed-device country uses the primary user UsageLocation or CountryOrRegion.'
    'Evidence Status' = 'Observed'
}
$directorySummary | Export-Csv -LiteralPath $DirectorySummaryOutputPath -NoTypeInformation -Encoding utf8NoBOM

$managedDevices = $evidence.Count
$entraCoveredManaged = @($evidence | Where-Object 'Entra Match' -eq 'Matched').Count
$windows11Covered = @($evidence | Where-Object 'Windows Generation' -eq 'Windows 11').Count
$signalEvidence = @(
    [pscustomobject][ordered]@{ 'Signal Sort' = 9; 'Signal' = 'Secure Boot disabled'; 'Affected Devices' = @($ad | Where-Object SecureBootStatus -eq 'Disabled').Count; 'Covered Devices' = @($ad | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SecureBootStatus) }).Count; 'Managed Devices Affected' = @($evidence | Where-Object 'Secure Boot State' -eq 'Disabled').Count; 'Managed Devices Covered' = @($evidence | Where-Object 'Secure Boot State' -ne 'Not observed').Count; 'Source Findings' = @($ad | Where-Object SecureBootStatus -eq 'Disabled').Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 10; 'Signal' = 'Endpoint Analytics score below 50'; 'Affected Devices' = @($endpointPerformanceIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.EndpointAnalyticsScore) -and [double]$_.EndpointAnalyticsScore -lt 50 }).Count; 'Covered Devices' = @($endpointPerformanceIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.EndpointAnalyticsScore) }).Count; 'Managed Devices Affected' = @($evidence | Where-Object 'Endpoint Analytics State' -eq 'Below 50').Count; 'Managed Devices Covered' = @($evidence | Where-Object 'Endpoint Analytics State' -ne 'Not observed').Count; 'Source Findings' = @($endpointPerformanceIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.EndpointAnalyticsScore) -and [double]$_.EndpointAnalyticsScore -lt 50 }).Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 11; 'Signal' = 'Devices with stop errors'; 'Affected Devices' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.StopErrorCount) -and [double]$_.StopErrorCount -gt 0 }).Count; 'Covered Devices' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.StopErrorCount) }).Count; 'Managed Devices Affected' = @($evidence | Where-Object 'Stop Error State' -eq 'Stop errors observed').Count; 'Managed Devices Covered' = @($evidence | Where-Object 'Stop Error State' -ne 'Not observed').Count; 'Source Findings' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.StopErrorCount) -and [double]$_.StopErrorCount -gt 0 }).Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 12; 'Signal' = 'Entra hardware ID conflicts'; 'Affected Devices' = $hardwareConflictDeviceIds.Count; 'Covered Devices' = $entra.Count; 'Managed Devices Affected' = @($evidence | Where-Object 'Entra Hardware ID Conflict' -eq 'Yes').Count; 'Managed Devices Covered' = $entraCoveredManaged; 'Source Findings' = $hardwareConflicts.Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 13; 'Signal' = 'Device removal candidates'; 'Affected Devices' = $removalCandidateDeviceIds.Count; 'Covered Devices' = $entra.Count; 'Managed Devices Affected' = @($evidence | Where-Object 'Removal Candidate' -eq 'Yes').Count; 'Managed Devices Covered' = $entraCoveredManaged; 'Source Findings' = $removalCandidateDeviceIds.Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 14; 'Signal' = 'Older Windows 11 releases'; 'Affected Devices' = @($evidence | Where-Object 'Windows 11 Release State' -eq 'Older release').Count; 'Covered Devices' = $windows11Covered; 'Managed Devices Affected' = @($evidence | Where-Object 'Windows 11 Release State' -eq 'Older release').Count; 'Managed Devices Covered' = $windows11Covered; 'Source Findings' = @($evidence | Where-Object 'Windows 11 Release State' -eq 'Older release').Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 15; 'Signal' = 'AD-to-Intune unmatched devices'; 'Affected Devices' = @($evidence | Where-Object 'AD to Intune Match' -eq 'Unmatched').Count; 'Covered Devices' = $managedDevices; 'Managed Devices Affected' = @($evidence | Where-Object 'AD to Intune Match' -eq 'Unmatched').Count; 'Managed Devices Covered' = $managedDevices; 'Source Findings' = @($evidence | Where-Object 'AD to Intune Match' -eq 'Unmatched').Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 16; 'Signal' = 'Startup score below 50'; 'Affected Devices' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.StartupScore) -and [double]$_.StartupScore -lt 50 }).Count; 'Covered Devices' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.StartupScore) }).Count; 'Managed Devices Affected' = @($evidence | Where-Object 'Startup Score State' -eq 'Below 50').Count; 'Managed Devices Covered' = @($evidence | Where-Object 'Startup Score State' -ne 'Not observed').Count; 'Source Findings' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.StartupScore) -and [double]$_.StartupScore -lt 50 }).Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 17; 'Signal' = 'Boot time over 60 seconds'; 'Affected Devices' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CoreBootTime) -and [double]$_.CoreBootTime -gt 60 }).Count; 'Covered Devices' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CoreBootTime) }).Count; 'Managed Devices Affected' = @($evidence | Where-Object { $_.'Core Boot Time (s)' -and [double]$_.'Core Boot Time (s)' -gt 60 }).Count; 'Managed Devices Covered' = @($evidence | Where-Object { $_.'Core Boot Time (s)' }).Count; 'Source Findings' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CoreBootTime) -and [double]$_.CoreBootTime -gt 60 }).Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 18; 'Signal' = 'Sign-in time over 30 seconds'; 'Affected Devices' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CoreSignInTime) -and [double]$_.CoreSignInTime -gt 30 }).Count; 'Covered Devices' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CoreSignInTime) }).Count; 'Managed Devices Affected' = @($evidence | Where-Object { $_.'Core Sign-in Time (s)' -and [double]$_.'Core Sign-in Time (s)' -gt 30 }).Count; 'Managed Devices Covered' = @($evidence | Where-Object { $_.'Core Sign-in Time (s)' }).Count; 'Source Findings' = @($endpointStartupIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CoreSignInTime) -and [double]$_.CoreSignInTime -gt 30 }).Count }
    [pscustomobject][ordered]@{ 'Signal Sort' = 19; 'Signal' = 'App reliability score below 50'; 'Affected Devices' = @($endpointPerformanceIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.AppReliabilityScore) -and [double]$_.AppReliabilityScore -ge 0 -and [double]$_.AppReliabilityScore -lt 50 }).Count; 'Covered Devices' = @($endpointPerformanceIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.AppReliabilityScore) -and [double]$_.AppReliabilityScore -ge 0 }).Count; 'Managed Devices Affected' = @($evidence | Where-Object 'App Reliability State' -eq 'Below 50').Count; 'Managed Devices Covered' = @($evidence | Where-Object 'App Reliability State' -ne 'Not observed').Count; 'Source Findings' = @($endpointPerformanceIndex.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.AppReliabilityScore) -and [double]$_.AppReliabilityScore -ge 0 -and [double]$_.AppReliabilityScore -lt 50 }).Count }
)
$signalEvidence | Export-Csv -LiteralPath $SignalsOutputPath -NoTypeInformation -Encoding utf8NoBOM

$summary = [ordered]@{
    OutputPath = $OutputPath
    Rows = $evidence.Count
    EntraMatches = @($evidence | Where-Object 'Entra Match' -eq 'Matched').Count
    AdMatches = @($evidence | Where-Object 'AD to Intune Match' -eq 'Matched').Count
    EnabledAdWindowsPcs = $enabledAdWindowsPcs.Count
    EnabledAdWindowsPcsManagedByIntune = $enabledAdWindowsPcsManaged.Count
    EnabledAdWindowsPcsManagedByIntuneRate = if ($enabledAdWindowsPcs.Count -gt 0) { [math]::Round($enabledAdWindowsPcsManaged.Count / [double]$enabledAdWindowsPcs.Count, 6) } else { $null }
    DisabledAdComputerObjects = $disabledAdComputerObjects
    DisabledEntraDevices = $disabledEntraDevices
    UserLocations = @($evidence | Where-Object Country -ne 'Unknown').Count
    LowDisk = @($evidence | Where-Object 'Disk Space State' -eq 'Low').Count
    ApproachingLowDisk = @($evidence | Where-Object 'Disk Space State' -eq 'Approaching Low').Count
    NotSynced30Days = @($evidence | Where-Object { $lastSync = Convert-ToDateTimeOrNull $_.'Last Sync DateTime'; $null -ne $lastSync -and $lastSync -lt $now.Date.AddDays(-30) }).Count
    WindowsUpdateHardFailures = @($evidence | Where-Object 'Windows Update State' -in @('Error', 'Rollback')).Count
    Windows10Remaining = @($evidence | Where-Object 'Windows Generation' -eq 'Windows 10').Count
    Windows10NotCapable = @($evidence | Where-Object { $_.'Windows Generation' -eq 'Windows 10' -and $_.'Windows 11 Upgrade Eligibility' -eq 'Not capable' }).Count
    Windows10Capable = @($evidence | Where-Object { $_.'Windows Generation' -eq 'Windows 10' -and $_.'Windows 11 Upgrade Eligibility' -eq 'Capable' }).Count
    Windows10NotAssessedOrUnknown = @($evidence | Where-Object { $_.'Windows Generation' -eq 'Windows 10' -and $_.'Windows 11 Upgrade Eligibility' -in @('Not assessed', 'Unknown') }).Count
    EnabledAdWindows10WithoutIntune = $adOnlyWindows10
    AdWindows10AmbiguousIntuneMatches = $ambiguousWindows10Matches
    Windows10EligibilityUndeterminedAllSources = $intuneWindows10Undetermined + $adOnlyWindows10
    SecondWaveSignals = @($signalEvidence | ForEach-Object { [ordered]@{ Signal = $_.Signal; AffectedDevices = $_.'Affected Devices'; CoveredDevices = $_.'Covered Devices'; ManagedDevicesAffected = $_.'Managed Devices Affected'; ManagedDevicesCovered = $_.'Managed Devices Covered'; SourceFindings = $_.'Source Findings' } })
    SourceFiles = @($intunePath, $usersPath, $entraPath, $adPath, $upgradeEligibilityPath, $windowsUpdatePath, $endpointPerformancePath, $endpointStartupPath, $hardwareConflictPath, $removalCandidatePath)
}

[pscustomobject]$summary | ConvertTo-Json -Depth 3

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB5W9A0JC9MSN8n
# qMtYLzB8O0eco4Qy5fagXauf6oA2UqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC7aZUT0tqYF9pJLQpfXXz/qVhQ4fQBS46VGMQJP4qYxMA0GCSqG
# SIb3DQEBAQUABIIBgBF+/JfCJDPSZEYQQY2waa/UtFLNqOJycuunHirQg9KmVoUM
# g4GZirvLd1wod5LnYcNPZN84udDMmzLEe9Mh3YV6MaW9LIhM3hKC878UFPrKEkTq
# Jh/Efy9z/AQ9fcCWruFmmWTNbC0V0An4ixLfB0L3nm97puwh4zqVqK+xsC5+SjEL
# b2hpVZJu5sFrrqdfW1ELjn44/ysyrEp2+THoPywfAXQy0/tG5195Q0cRiNgtQ8kM
# nOvjg+wMRke+tcg4i0jfB7KBe4nNd7rCWyQeuOH1HPV9YqRN4R+4mlzsWrktNmV6
# laZ7KMRqhupV2D8Hty1RVyeVlyE2n7H4ult2bvNFmLoYe1EwrVY/9EYUR63mIFM2
# 4n1rqX2NEqJFZMgV02HJ3MJq6VKUlQyGb2syGF71lgNQIVpEctw8A8tQNdu/wK8M
# gQmjA2X56/Dqw4S4RR2MPEyyY5w66UEsKucZWcE37/OvzNC+s5WbfVPQwdlHE5Pk
# jlSyDqFVdXh8Fuv5JaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQyMTAx
# NTdaMC8GCSqGSIb3DQEJBDEiBCC3niD/m8KMv3n6SJOjvhXPIPe17XX9x0oT6ZJQ
# ZGclrjANBgkqhkiG9w0BAQEFAASCAgC1ipy7G7GB/4mTIMzpHwOf1d5OeFjNWCzP
# VRYkPCCa5LKrPtRGxWn1O5gxdka/mlqZsiWidE2LslDZEYcg+QG46mt4QfiXfmpJ
# CmK934IA4xt1FrFbL0W2s0fe4ut+ImojIKwHXDXABHdNtOKVe+tosjqeZgn0tYQp
# 36K/gmrxjS9fWC9IQSstHLzq+w7lP2NJM9aTXWzUvIfJfJCEw8mhpbRMpQKr8s5F
# 7USBgLFI557z1MzdEIYgB2JJg4Xzagdy/KpOQy4U9hGk4vH+5yL1mWhdUU2lJhDk
# lFKMuOB9fKvOQ7z/OfIcz3t6+MjQ6Hgvk6tvR8S+q7CYKqpb2XZGFBtCorh/0UYK
# uRVACy1Cnpb49fEizvWfNp84Dtqi84EIBiSc59UCyVphFwg33AOkfzetD6zvzMDD
# UvDXJ63vzW0DN/zBxCAcRwrSFTFfcyqHGHL6D0RtFm6sLe2bmOtGQ5HtjpHFZ2IB
# vfbVRr41S4CSni3SA+mAoz5qa5Txeov/d67xSDG/5BL4+PgFH2dAYaIokXxDZlaZ
# YqP0SizdW+yDr2fOCfiwEiVbkLtKeJ1RvpsUe7Psw1Be4qkKUSJ2hJ4rmCbtv0RJ
# y7pg9lVFkeQMKNdNZXTDKDrwj4DYgQgmXES30zAFlcJP+MmVkAu/PccJRGo69lnL
# LMFPp1vnHw==
# SIG # End signature block
