[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$DeviceEvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\DeviceInventoryEvidence.csv'),

    [string]$UserEvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\UserInventoryEvidence.csv'),

    [string]$ExtendedEvidenceRoot = $DataRoot,

    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\SecurityControlEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-PassValue {
    param([AllowNull()][object]$Value)
    return ([string]$Value).Trim() -match '^(?i:pass|passed|enabled|healthy|ok|compliant|true)$'
}

function Test-FailValue {
    param([AllowNull()][object]$Value)
    return ([string]$Value).Trim() -match '^(?i:fail|failed|disabled|unhealthy|critical|warning|noncompliant|false|error)$'
}

function Test-TrueValue {
    param([AllowNull()][object]$Value)
    return ([string]$Value).Trim() -match '^(?i:true|1|yes|enabled)$'
}

function ConvertTo-InvariantDouble {
    param([AllowNull()][object]$Value)
    $text = ([string]$Value).Trim().Replace(' ', '').Replace(',', '.')
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $number = 0.0
    if (-not [double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $null
    }
    return $number
}

function Resolve-OptionalEvidencePath {
    param([Parameter(Mandatory = $true)][string]$FileName)
    foreach ($root in @($ExtendedEvidenceRoot, $DataRoot) | Select-Object -Unique) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        $candidate = Join-Path $root $FileName
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $null
}

function New-ControlRow {
    param(
        [int]$Sort,
        [string]$Domain,
        [string]$Control,
        [string]$EvidenceStatus,
        [AllowNull()][Nullable[int64]]$Covered,
        [AllowNull()][Nullable[int64]]$Healthy,
        [AllowNull()][Nullable[int64]]$Affected,
        [string]$MetricUnit,
        [string]$Source,
        [string]$Action,
        [string]$SnapshotDate,
        [AllowNull()][Nullable[double]]$HealthRateOverride = $null
    )
    $coverageState = if ($EvidenceStatus -eq 'Observed' -and $Covered -gt 0) { 'Covered' } else { 'Not collected' }
    $gapState = if ($EvidenceStatus -ne 'Observed' -or $null -eq $Affected) { 'Not assessed' } elseif ($Affected -gt 0) { 'Gap observed' } else { 'No observed gap' }
    $severity = if ($EvidenceStatus -ne 'Observed') { 'Evidence gap' } elseif ($Affected -gt 0) { 'High' } else { 'None' }
    [pscustomobject][ordered]@{
        'Control Sort' = $Sort
        'Security Domain' = $Domain
        'Control Name' = $Control
        'Severity Label' = $severity
        'Evidence Status' = $EvidenceStatus
        'Coverage State' = $coverageState
        'Covered Entities' = $Covered
        'Healthy Entities' = $Healthy
        'Affected Entities' = $Affected
        'Metric Unit' = $MetricUnit
        'Health Rate' = if ($null -ne $HealthRateOverride) { ([double]$HealthRateOverride).ToString('0.####', [Globalization.CultureInfo]::InvariantCulture) } elseif ($null -ne $Covered -and $Covered -gt 0 -and $null -ne $Healthy) { ([double]$Healthy / [double]$Covered).ToString('0.####', [Globalization.CultureInfo]::InvariantCulture) } else { '' }
        'Gap State' = $gapState
        'Evidence Source' = $Source
        'Recommended Action' = $Action
        'Snapshot Date' = $SnapshotDate
    }
}

$policyPath = Join-Path $DataRoot 'Intune_Devices_Compliance_Policies.csv'
$adHealthPath = Join-Path $DataRoot 'AD_HealthCheck.csv'
$syncHealthPath = Join-Path $DataRoot 'M365_Entra_AzureADConnect_SyncHealth.csv'
foreach ($path in @($policyPath, $adHealthPath, $syncHealthPath, $DeviceEvidencePath, $UserEvidencePath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required private evidence file was not found: $path" }
}

$devices = @(Import-Csv -LiteralPath $DeviceEvidencePath)
$workforceUsers = @(Import-Csv -LiteralPath $UserEvidencePath)
$policies = @(Import-Csv -LiteralPath $policyPath)
$adHealth = @(Import-Csv -LiteralPath $adHealthPath)
$syncHealth = @(Import-Csv -LiteralPath $syncHealthPath)
$snapshotDate = (Get-Item -LiteralPath $policyPath).LastWriteTime.Date.ToString('yyyy-MM-dd')

$controls = [System.Collections.Generic.List[object]]::new()
$knownCompliance = @($devices | Where-Object { $_.'Compliance State' -ne 'Unknown' })
$compliant = @($knownCompliance | Where-Object 'Compliance State' -eq 'Compliant')
$noncompliant = @($knownCompliance | Where-Object { $_.'Compliance State' -in @('Noncompliant','In Grace Period') })
$controls.Add((New-ControlRow 1 'Device security' 'Device compliance' 'Observed' $knownCompliance.Count $compliant.Count $noncompliant.Count 'Devices' 'Intune device inventory' 'Remediate noncompliant devices and validate grace-period exceptions.' $snapshotDate))

$secureBootCovered = @($devices | Where-Object { $_.'Secure Boot State' -in @('Enabled','Disabled','Error') })
$secureBootHealthy = @($secureBootCovered | Where-Object 'Secure Boot State' -eq 'Enabled')
$secureBootAffected = @($secureBootCovered | Where-Object { $_.'Secure Boot State' -in @('Disabled','Error') })
$controls.Add((New-ControlRow 2 'Device security' 'Secure Boot' 'Observed' $secureBootCovered.Count $secureBootHealthy.Count $secureBootAffected.Count 'Devices' 'AD computer inventory reconciled to Intune' 'Enable Secure Boot on supported devices and investigate collection errors.' $snapshotDate))

function Get-DevicePolicyState {
    param([object[]]$Rows, [string]$Property)
    $byDevice = @{}
    foreach ($row in $Rows) {
        $key = ([string]$row.AzureADDeviceId).Trim().ToLowerInvariant()
        if (-not $key) { $key = ([string]$row.DeviceName).Trim().ToLowerInvariant() }
        if (-not $key) { continue }
        $value = $row.$Property
        if ([string]::IsNullOrWhiteSpace([string]$value)) { continue }
        if (-not $byDevice.ContainsKey($key)) { $byDevice[$key] = 'Pass' }
        if (Test-FailValue $value) { $byDevice[$key] = 'Fail' }
        elseif (-not (Test-PassValue $value) -and $byDevice[$key] -ne 'Fail') { $byDevice[$key] = 'Unknown' }
    }
    return $byDevice
}

$bitLocker = Get-DevicePolicyState $policies 'BitLocker/Encryption'
$bitLockerHealthy = @($bitLocker.Values | Where-Object { $_ -eq 'Pass' }).Count
$bitLockerAffected = @($bitLocker.Values | Where-Object { $_ -eq 'Fail' }).Count
$controls.Add((New-ControlRow 3 'Device security' 'Disk encryption' 'Observed' ([int64]$bitLocker.Count) ([int64]$bitLockerHealthy) ([int64]$bitLockerAffected) 'Devices' 'Intune compliance policy evidence' 'Remediate BitLocker or encryption policy failures and increase policy evidence coverage.' $snapshotDate))

$codeIntegrity = Get-DevicePolicyState $policies 'CodeIntegrity'
$codeIntegrityHealthy = @($codeIntegrity.Values | Where-Object { $_ -eq 'Pass' }).Count
$codeIntegrityAffected = @($codeIntegrity.Values | Where-Object { $_ -eq 'Fail' }).Count
$controls.Add((New-ControlRow 4 'Device security' 'Code integrity' 'Observed' ([int64]$codeIntegrity.Count) ([int64]$codeIntegrityHealthy) ([int64]$codeIntegrityAffected) 'Devices' 'Intune compliance policy evidence' 'Investigate code-integrity failures and close devices without observed policy evidence.' $snapshotDate))

$adHealthy = @($adHealth | Where-Object Status -eq 'OK').Count
$adAffected = @($adHealth | Where-Object { $_.Status -in @('Warning','Critical') }).Count
$controls.Add((New-ControlRow 5 'Identity infrastructure' 'Active Directory health checks' 'Observed' ([int64]$adHealth.Count) ([int64]$adHealthy) ([int64]$adAffected) 'Checks' 'AD health check evidence' 'Resolve critical and warning health checks before relying on directory services.' $snapshotDate))

$syncHealthy = @($syncHealth | Where-Object Status -eq 'OK').Count
$syncAffected = @($syncHealth | Where-Object { $_.Status -ne 'OK' }).Count
$controls.Add((New-ControlRow 6 'Identity infrastructure' 'Entra Connect synchronization health' 'Observed' ([int64]$syncHealth.Count) ([int64]$syncHealthy) ([int64]$syncAffected) 'Connectors' 'Entra Connect synchronization health' 'Restore synchronization health and confirm the next successful cycle.' $snapshotDate))

$conditionalAccessPath = Resolve-OptionalEvidencePath 'M365_Entra_ConditionalAccessPolicies.csv'
if ($conditionalAccessPath) {
    $conditionalAccess = @(Import-Csv -LiteralPath $conditionalAccessPath)
    $enabledPolicies = @($conditionalAccess | Where-Object { $_.State -eq 'enabled' })
    $nonEnforcedPolicies = @($conditionalAccess | Where-Object { $_.State -ne 'enabled' })
    $conditionalAccessDate = (Get-Item -LiteralPath $conditionalAccessPath).LastWriteTime.Date.ToString('yyyy-MM-dd')
    $controls.Add((New-ControlRow 7 'Access control' 'Conditional Access enforcement' 'Observed' ([int64]$conditionalAccess.Count) ([int64]$enabledPolicies.Count) ([int64]$nonEnforcedPolicies.Count) 'Policies' 'Microsoft Graph Conditional Access policies' 'Review disabled and report-only policies, confirm scope, and enforce approved controls.' $conditionalAccessDate))
} else {
    $controls.Add((New-ControlRow 7 'Access control' 'Conditional Access enforcement' 'Not collected' $null $null $null 'Policies' 'Not collected' 'Collect Conditional Access policy evidence before assessing enforcement.' $snapshotDate))
}

$authenticationMethodsPath = Resolve-OptionalEvidencePath 'M365_Entra_AuthenticationMethodsRegistration.csv'
if ($authenticationMethodsPath) {
    $mfaCovered = @($workforceUsers | Where-Object { $_.'MFA Registration State' -ne 'No evidence' })
    $mfaRegistered = @($mfaCovered | Where-Object 'MFA Registration State' -eq 'Registered')
    $mfaNotRegistered = @($mfaCovered | Where-Object 'MFA Registration State' -eq 'Not registered')
    $authenticationMethodsDate = (Get-Item -LiteralPath $authenticationMethodsPath).LastWriteTime.Date.ToString('yyyy-MM-dd')
    $controls.Add((New-ControlRow 8 'Identity protection' 'MFA registered among workforce' 'Observed' ([int64]$mfaCovered.Count) ([int64]$mfaRegistered.Count) ([int64]$mfaNotRegistered.Count) 'Users' 'Entra authentication-method registration matched by UserId to enabled AD-linked workforce accounts' 'Register an MFA method for workforce accounts without one; evaluate Conditional Access enforcement separately.' $authenticationMethodsDate))
} else {
    $controls.Add((New-ControlRow 8 'Identity protection' 'MFA registered among workforce' 'Not collected' $null $null $null 'Users' 'Not collected' 'Collect authentication-method registration evidence before assessing workforce MFA registration.' $snapshotDate))
}

$defenderPath = Resolve-OptionalEvidencePath 'Intune_EndpointSecurity_DefenderAgents.csv'
if ($defenderPath) {
    $defender = @(Import-Csv -LiteralPath $defenderPath)
    $defenderHealthy = @($defender | Where-Object { (Test-TrueValue $_.MalwareProtectionEnabled) -and (Test-TrueValue $_.RealTimeProtectionEnabled) -and -not (Test-TrueValue $_.SignatureUpdateOverdue) })
    $defenderAffected = @($defender | Where-Object { -not ((Test-TrueValue $_.MalwareProtectionEnabled) -and (Test-TrueValue $_.RealTimeProtectionEnabled) -and -not (Test-TrueValue $_.SignatureUpdateOverdue)) })
    $defenderDate = (Get-Item -LiteralPath $defenderPath).LastWriteTime.Date.ToString('yyyy-MM-dd')
    $controls.Add((New-ControlRow 9 'Endpoint protection' 'Defender protection active' 'Observed' ([int64]$defender.Count) ([int64]$defenderHealthy.Count) ([int64]$defenderAffected.Count) 'Devices' 'Intune endpoint security Defender agents' 'Restore malware and real-time protection, then update overdue signatures.' $defenderDate))
} else {
    $controls.Add((New-ControlRow 9 'Endpoint protection' 'Defender protection active' 'Not collected' $null $null $null 'Devices' 'Not collected' 'Collect endpoint protection health evidence before assessing Defender coverage.' $snapshotDate))
}

$firewallPath = Resolve-OptionalEvidencePath 'Intune_EndpointSecurity_FirewallStatus.csv'
if ($firewallPath) {
    $firewall = @(Import-Csv -LiteralPath $firewallPath)
    $firewallHealthy = @($firewall | Where-Object { $_.FirewallStatus -eq 'Enabled' })
    $firewallAffected = @($firewall | Where-Object { $_.FirewallStatus -ne 'Enabled' })
    $firewallDate = (Get-Item -LiteralPath $firewallPath).LastWriteTime.Date.ToString('yyyy-MM-dd')
    $controls.Add((New-ControlRow 10 'Endpoint protection' 'Firewall enabled' 'Observed' ([int64]$firewall.Count) ([int64]$firewallHealthy.Count) ([int64]$firewallAffected.Count) 'Devices' 'Intune endpoint security firewall status' 'Enable the firewall and investigate devices reporting disabled or limited status.' $firewallDate))
} else {
    $controls.Add((New-ControlRow 10 'Endpoint protection' 'Firewall enabled' 'Not collected' $null $null $null 'Devices' 'Not collected' 'Collect endpoint firewall health evidence before assessing protection coverage.' $snapshotDate))
}

$secureScorePath = Resolve-OptionalEvidencePath 'M365_Security_SecureScore.csv'
if ($secureScorePath) {
    $secureScores = @(Import-Csv -LiteralPath $secureScorePath)
    $latestSecureScore = $secureScores | Sort-Object { [datetime]$_.ScoreDate } -Descending | Select-Object -First 1
    $scorePercentage = ConvertTo-InvariantDouble $latestSecureScore.ScorePercentage
    $scoreRate = if ($null -ne $scorePercentage) { [Math]::Max(0.0, [Math]::Min(1.0, $scorePercentage / 100.0)) } else { $null }
    $scoreHealthy = if ($scoreRate -ge 1.0) { 1 } else { 0 }
    $scoreAffected = if ($scoreRate -lt 1.0) { 1 } else { 0 }
    $secureScoreDate = if ($latestSecureScore.ScoreDate) { ([datetime]$latestSecureScore.ScoreDate).ToString('yyyy-MM-dd') } else { (Get-Item -LiteralPath $secureScorePath).LastWriteTime.Date.ToString('yyyy-MM-dd') }
    $controls.Add((New-ControlRow 11 'Security posture' 'Microsoft Secure Score' 'Observed' 1 $scoreHealthy $scoreAffected 'Score' 'Microsoft Graph Secure Score' 'Prioritize the highest-impact Secure Score improvement actions and track the observed score trend.' $secureScoreDate $scoreRate))
} else {
    $controls.Add((New-ControlRow 11 'Security posture' 'Microsoft Secure Score' 'Not collected' $null $null $null 'Score' 'Not collected' 'Collect Microsoft Secure Score evidence before reporting a posture score or trend.' $snapshotDate))
}

$outputDirectory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$controls | Sort-Object 'Control Sort' | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8NoBOM

[pscustomobject]@{
    OutputPath = $OutputPath
    Controls = $controls.Count
    ObservedControls = @($controls | Where-Object 'Evidence Status' -eq 'Observed').Count
    ControlsWithoutEvidence = @($controls | Where-Object 'Evidence Status' -eq 'Not collected').Count
    ControlsWithGaps = @($controls | Where-Object 'Gap State' -eq 'Gap observed').Count
} | Format-List
