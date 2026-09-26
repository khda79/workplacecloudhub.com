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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC1/rNZ3Yw91MpM
# mzYnFLQWLutlfrWaWgx+K7vSLXFhu6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFAMwsf9CM3QZvwJyXbKJqt0JwkRMZkp9xYj7y+TqBYhMA0GCSqG
# SIb3DQEBAQUABIIBgEDLWdgNr8LegeJbWjT9h6PIHfsBaXafQ2ge2WuCcIMZtjuT
# OixZ3TPBvh0dlyX5rY6JqxUVt5YeyhTplz7bW//KgKoGEvf844ZCPuYtqcFrrjtd
# Erp2o+6wvw7Ze2rc7E24Qe/sF9inag5u4deHKDPYREq6uNa9oyRlxmXVgF7YVDKv
# ijVhOSSyjqtJ1PllFZg7I7L9uijcFJZzthMuDH6PE1dUwdHqjwWBIAXjXy+WakaV
# oIQqMUsSjXX9gNSz0VkdMK9CvRNSyyRJFhTbfUHUiNA/3Mb4PQvO59DnS1QH73sr
# HNPIA8G2YN2d4pd3OXmw1UCJ37JXq5cl69jB/L2htlYWzWxaD+sx+tEuoqYIFuh1
# /E/hr4z6cxYdgoTzs6ywPfpDjRiq5geWrZwCOvMh/lvtxJAmlA1CWOS2p9Q/7oLo
# oSWwZzNfJTj0ClbMzcH02iR419b9uXwD8rTadCLyH1zBAsJCBMoE5SK/PXA8pl0Y
# Q/LniEDCVMjNoTz2n6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjYyMzUw
# MDJaMC8GCSqGSIb3DQEJBDEiBCAGEfItAAZnn4ADVYjGWtwhKTvi5qHM35lbgS4f
# b+WyFzANBgkqhkiG9w0BAQEFAASCAgAyjVC2q8oPmYfH7x2XpQNNdQv+5cjHywXF
# DHFDueMxo99o5xuRHe8uIW8AABY5u52clwIOlpI5nCQ+WMAERPftC/HwiQdyeqPb
# O4EMqyOhw2PlkLKWT1Zw5Cy0DOk4B8UIle7AHOtxVOmCg4vNMm5VtV/7fMgtYz7f
# vJtC5Keoye6ZfcmSUw09QEmABwuzlwDEJeXNRakawhNca2qBTLRGx2CyEsF73Mr4
# cTiM3OWafWIiaF4SE/r4spaYKgKJcP2wH9BxAij5aCT+RnJC7rjKdsdvSk+reAca
# i2Y+2hKTzwSA2EGlZ3jB6rMfNWjtFc1NLmtZYJcAMU/3bX9MpUrSpiCWJvgHCv07
# 14qfVN0Cnt9dZn2uBrlnDxhLP005hLSJJZMOEkveVD3APPe3ib/41qP4NZvkpaH0
# rnuCsDY8w52E7Ei9THCDfE+QAH5QhCrxEolxB+dPzjE815saf2XDBTEWTKCApuS+
# KTETY43888bDRBUcT1Ce9N0sVYTIvKcafNKayQ8VaCn3NzqSj2NniV3yFvLnBDds
# XuTz7lyMigZvwj4qZmAL3yO6WXntFY2Ea2yhLpNueGZ5SBFLH/aecuYgcuAVmcZz
# 6Ou3BGfqJiTuhS1UXXIU4dIIJidDisFHCcY8rYvd7AUVmf5eyPGOSWAU/4rf1K3E
# sFo2nRh4QA==
# SIG # End signature block
