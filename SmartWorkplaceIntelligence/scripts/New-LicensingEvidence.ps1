[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$UserEvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\UserInventoryEvidence.csv'),

    [string]$MailboxEvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\MailboxEvidence.csv'),

    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\LicenseEvidence.csv'),

    [string]$OptimizationOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\LicenseOptimizationEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-NormalizedKey {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Convert-ToDouble {
    param([AllowNull()][object]$Value)
    $number = 0.0
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return 0.0 }
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [ref]$number)) { return $number }
    return 0.0
}

function Convert-ToDoubleOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $number = 0.0
    $styles = [Globalization.NumberStyles]::AllowLeadingWhite -bor [Globalization.NumberStyles]::AllowTrailingWhite -bor [Globalization.NumberStyles]::AllowLeadingSign -bor [Globalization.NumberStyles]::AllowDecimalPoint
    foreach ($culture in @([Globalization.CultureInfo]::InvariantCulture, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'))) {
        if ([double]::TryParse([string]$Value, $styles, $culture, [ref]$number) -and $number -ge 0 -and -not [double]::IsNaN($number) -and -not [double]::IsInfinity($number)) { return $number }
    }
    return $null
}

function Convert-ToInvariantDecimalText {
    param([double]$Value)
    return $Value.ToString('0.####', [Globalization.CultureInfo]::InvariantCulture)
}

function Convert-ToDateOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $parsed = [datetime]::MinValue
    foreach ($culture in @([Globalization.CultureInfo]::InvariantCulture, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'))) {
        if ([datetime]::TryParse([string]$Value, $culture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) { return $parsed }
    }
    return $null
}

function Test-OfficeDesktopExplicitNoUse {
    param([AllowNull()][object]$Row)
    if ($null -eq $Row) { return $false }
    $fields = @('Outlook (Windows)', 'Word (Windows)', 'Excel (Windows)', 'PowerPoint (Windows)', 'OneNote (Windows)')
    foreach ($field in $fields) {
        $property = $Row.PSObject.Properties[$field]
        if ($null -eq $property -or ([string]$property.Value).Trim().ToLowerInvariant() -notin @('false', 'no', '0')) { return $false }
    }
    return $true
}

function Get-Microsoft365Plan {
    param([AllowNull()][object]$SkuPartNumber)
    $sku = ([string]$SkuPartNumber).Trim().ToUpperInvariant()
    if ($sku -eq 'SPE_F1') { return 'F3' }
    if ($sku -eq 'SPE_E3') { return 'E3' }
    if ($sku -match '(^|_)SPE_E5($|_)|M365.*E5|MICROSOFT365.*E5') { return 'E5' }
    return 'Other'
}

function New-UniqueIndex {
    param([object[]]$Rows, [string]$PropertyName)
    $groups = @{}
    foreach ($row in $Rows) {
        $key = Get-NormalizedKey $row.$PropertyName
        if (-not $key) { continue }
        if (-not $groups.ContainsKey($key)) { $groups[$key] = [System.Collections.Generic.List[object]]::new() }
        $groups[$key].Add($row)
    }
    $index = @{}
    foreach ($entry in $groups.GetEnumerator()) { if ($entry.Value.Count -eq 1) { $index[$entry.Key] = $entry.Value[0] } }
    return $index
}

$tenantPath = Join-Path $DataRoot 'M365_Licenses_Tenant.csv'
$assignmentPath = Join-Path $DataRoot 'M365_Licenses_Users.csv'
$entraUsersPath = Join-Path $DataRoot 'M365_Users_Active.csv'
$activationPath = Join-Path $DataRoot 'M365_Apps_Activations.csv'
$officeUsage180Path = Join-Path $DataRoot 'M365_Apps_Usage_180D.csv'
foreach ($path in @($tenantPath, $assignmentPath, $entraUsersPath, $activationPath, $officeUsage180Path, $UserEvidencePath, $MailboxEvidencePath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required private evidence file was not found: $path" }
}

$capacity = @(Import-Csv -LiteralPath $tenantPath)
$assignments = @(Import-Csv -LiteralPath $assignmentPath)
$users = @(Import-Csv -LiteralPath $UserEvidencePath)
$entraUsers = @(Import-Csv -LiteralPath $entraUsersPath)
$activations = @(Import-Csv -LiteralPath $activationPath)
$officeUsage180ByUpn = @{}
foreach ($usage in @(Import-Csv -LiteralPath $officeUsage180Path)) {
    $periodProperty = $usage.PSObject.Properties['ReportPeriodRequested']
    if ($null -eq $periodProperty -or ([string]$periodProperty.Value).Trim() -ne 'D180') { throw 'M365_Apps_Usage_180D.csv does not contain the expected D180 report period.' }
    $usageUpn = Get-NormalizedKey $usage.'User Principal Name'
    $usageDate = Convert-ToDateOrNull $usage.'Report Refresh Date'
    if (-not $usageUpn -or $null -eq $usageDate) { continue }
    if (-not $officeUsage180ByUpn.ContainsKey($usageUpn) -or $usageDate -gt $officeUsage180ByUpn[$usageUpn].ReportDate) {
        $officeUsage180ByUpn[$usageUpn] = [pscustomobject]@{ ReportDate = $usageDate; Row = $usage }
    }
}
$mailboxes = @(Import-Csv -LiteralPath $MailboxEvidencePath)

$userByUpn = @{}
foreach ($user in $users) {
    $key = Get-NormalizedKey $user.'User Principal Name'
    if ($key -and -not $userByUpn.ContainsKey($key)) { $userByUpn[$key] = $user }
}
if (@($officeUsage180ByUpn.Keys | Where-Object { $userByUpn.ContainsKey($_) }).Count -eq 0) {
    throw 'The D180 Apps usage report has no usable, non-anonymized workforce matches. Check report refresh, UPN privacy settings and source schema before calculating E3-to-F3 candidates.'
}
$entraUserByUpn = New-UniqueIndex $entraUsers 'User principal name'
$mailboxByUpn = New-UniqueIndex $mailboxes 'Primary SMTP Address'
$lastDesktopActivationByUpn = @{}
foreach ($activation in $activations) {
    if (([string]$activation.'Product Type').Trim() -notmatch '(?i)MICROSOFT 365 APPS FOR ENTERPRISE') { continue }
    $upn = Get-NormalizedKey $activation.'User Principal Name'
    $date = Convert-ToDateOrNull $activation.'Last Activated Date'
    if (-not $upn -or $null -eq $date) { continue }
    if (-not $lastDesktopActivationByUpn.ContainsKey($upn) -or $date -gt $lastDesktopActivationByUpn[$upn]) { $lastDesktopActivationByUpn[$upn] = $date }
}

$assignmentsBySku = @{}
foreach ($assignment in $assignments) {
    $sku = Get-NormalizedKey $assignment.SkuPartNumber
    if (-not $sku) { continue }
    if (-not $assignmentsBySku.ContainsKey($sku)) { $assignmentsBySku[$sku] = [System.Collections.Generic.List[object]]::new() }
    $assignmentsBySku[$sku].Add($assignment)
}

$snapshotDateValue = (Get-Item -LiteralPath $tenantPath).LastWriteTime.Date
$snapshotDate = $snapshotDateValue.ToString('yyyy-MM-dd')
$optimizationRows = [System.Collections.Generic.List[object]]::new()
$evidence = foreach ($row in $capacity) {
    $skuKey = Get-NormalizedKey $row.TenantSkuPartNumber
    $skuAssignments = if ($assignmentsBySku.ContainsKey($skuKey)) { @($assignmentsBySku[$skuKey]) } else { @() }
    $enabled = Convert-ToDouble $row.TenantPrepaidEnabled
    $consumed = Convert-ToDouble $row.TenantConsumedUnits
    $available = [math]::Max(0, $enabled - $consumed)
    $name = [string]$row.TenantSkuDisplayName
    $microsoft365Plan = Get-Microsoft365Plan $row.TenantSkuPartNumber
    $excludedPattern = '(?i)trial|free|viral|adhoc|stream|windows store|business center|for iws|exploratory'
    $capacityClass = if ($enabled -gt 0 -and $enabled -lt 100000 -and $name -notmatch $excludedPattern) { 'Governed capacity' } else { 'Non-finite, free, or trial' }

    $distinctUsers = @($skuAssignments | ForEach-Object { Get-NormalizedKey $_.'User principal name' } | Where-Object { $_ } | Sort-Object -Unique)
    $reclaimable = 0
    $dormant = 0
    $disabledReclaimable = 0
    $e3ToF3ReviewCandidates = 0
    foreach ($upn in $distinctUsers) {
        $userEvidence = if ($userByUpn.ContainsKey($upn)) { $userByUpn[$upn] } else { $null }
        $entraUser = if ($entraUserByUpn.ContainsKey($upn)) { $entraUserByUpn[$upn] } else { $null }
        $mailbox = if ($mailboxByUpn.ContainsKey($upn)) { $mailboxByUpn[$upn] } else { $null }
        $accountState = if ($null -eq $entraUser) { 'Not observed' } elseif ((Get-NormalizedKey $entraUser.AccountEnabled) -eq 'true') { 'Enabled' } else { 'Disabled' }
        $activityState = if ($null -ne $userEvidence) { [string]$userEvidence.'Activity State' } else { 'Not observed' }
        $immediateReclaim = $accountState -eq 'Disabled' -or $activityState -in @('Stale >90D', 'Observed never used')
        if ($microsoft365Plan -eq 'E3' -and $accountState -eq 'Enabled' -and $null -ne $mailbox) {
            $mailboxSize = Convert-ToDoubleOrNull $mailbox.'Mailbox Size GB'
            $archiveDisabled = ([string]$mailbox.'Archive State').Trim() -eq 'Disabled'
            $officeUsage = if ($officeUsage180ByUpn.ContainsKey($upn)) { $officeUsage180ByUpn[$upn] } else { $null }
            $desktopAppsNotUsed = $null -ne $officeUsage -and $officeUsage.ReportDate.Date -ge (Get-Date).Date.AddDays(-10) -and (Test-OfficeDesktopExplicitNoUse -Row $officeUsage.Row)
            if ($null -ne $mailboxSize -and $mailboxSize -lt 2 -and $archiveDisabled -and $desktopAppsNotUsed) {
                $e3ToF3ReviewCandidates++
                $optimizationRows.Add([pscustomobject][ordered]@{
                    'User Principal Name' = $upn
                    'SKU Part Number' = [string]$row.TenantSkuPartNumber
                    'License Product' = if ($name) { $name.Trim() } else { [string]$row.TenantSkuPartNumber }
                    'Microsoft 365 Plan' = $microsoft365Plan
                    'Account State' = $accountState
                    'Activity State' = $activityState
                    'Last Activity Date' = if ($null -ne $userEvidence) { [string]$userEvidence.'Last Activity Date' } else { '' }
                    'Last Office Desktop Activation Date' = if ($lastDesktopActivationByUpn.ContainsKey($upn)) { $lastDesktopActivationByUpn[$upn].ToString('yyyy-MM-dd') } else { '' }
                    'Mailbox Size GB' = [string]$mailbox.'Mailbox Size GB'
                    'Archive State' = [string]$mailbox.'Archive State'
                    'Optimization Opportunity' = 'E3 to F3 review candidate'
                    'Optimization Reason' = 'Explicit no Office Windows app usage in the 180-day report, observed mailbox <2 GB, and disabled archive. Business review required; reclaim overlap is not additive.'
                    'Recommended Action' = if ($immediateReclaim) { 'Prioritize full license reclaim; consider F3 only if the account must remain licensed.' } else { 'Ask the business and licensing owner to validate desktop-app and workload requirements before any E3-to-F3 change.' }
                    'Evidence Status' = 'Observed; business validation required'
                    'Snapshot Date' = $snapshotDate
                })
            }
        }
        if ($immediateReclaim) {
            $reclaimable++
            if ($accountState -eq 'Disabled') { $disabledReclaimable++ }
            $reason = if ($accountState -eq 'Disabled') { 'Licensed account is disabled.' } else { 'No AD or Microsoft 365 activity was observed within the last 90 days.' }
            $optimizationRows.Add([pscustomobject][ordered]@{
                'User Principal Name' = $upn
                'SKU Part Number' = [string]$row.TenantSkuPartNumber
                'License Product' = if ($name) { $name.Trim() } else { [string]$row.TenantSkuPartNumber }
                'Microsoft 365 Plan' = $microsoft365Plan
                'Account State' = $accountState
                'Activity State' = $activityState
                'Last Activity Date' = if ($null -ne $userEvidence) { [string]$userEvidence.'Last Activity Date' } else { '' }
                'Last Office Desktop Activation Date' = if ($lastDesktopActivationByUpn.ContainsKey($upn)) { $lastDesktopActivationByUpn[$upn].ToString('yyyy-MM-dd') } else { '' }
                'Mailbox Size GB' = if ($null -ne $mailbox) { [string]$mailbox.'Mailbox Size GB' } else { '' }
                'Archive State' = if ($null -ne $mailbox) { [string]$mailbox.'Archive State' } else { 'Not observed' }
                'Optimization Opportunity' = 'Immediate reclaim review'
                'Optimization Reason' = $reason
                'Recommended Action' = 'Validate ownership and exception scope, then remove, reassign, or avoid renewing the license.'
                'Evidence Status' = 'Observed'
                'Snapshot Date' = $snapshotDate
            })
            continue
        }
        if ($activityState -eq 'Dormant 31-90D') { $dormant++ }
    }

    $utilization = if ($enabled -gt 0) { $consumed / $enabled } else { $null }
    if ($reclaimable -gt 0) {
        $priority = 'High'
        $signal = 'Inactive or never-used users retain license assignments'
        $action = 'Validate business need, then reclaim or reassign licenses from stale and never-used accounts.'
    }
    elseif ($capacityClass -eq 'Governed capacity' -and $null -ne $utilization -and $utilization -lt 0.5) {
        $priority = 'Medium'
        $signal = 'Governed license capacity is underused'
        $action = 'Review renewal quantity and assignment policy before the next purchasing cycle.'
    }
    else {
        $priority = 'None'
        $signal = 'No current license optimization signal'
        $action = 'Maintain assignment governance and refresh the evidence on schedule.'
    }

    [pscustomobject][ordered]@{
        'License Product' = if ($name) { $name.Trim() } else { [string]$row.TenantSkuPartNumber }
        'SKU Part Number' = [string]$row.TenantSkuPartNumber
        'Microsoft 365 Plan' = $microsoft365Plan
        'Capacity Class' = $capacityClass
        'Enabled Units' = [int64][math]::Round($enabled)
        'Consumed Units' = [int64][math]::Round($consumed)
        'Available Units' = [int64][math]::Round($available)
        'Capacity Utilization' = if ($null -ne $utilization) { Convert-ToInvariantDecimalText $utilization } else { '' }
        'Assigned Users' = $distinctUsers.Count
        'Direct Assignments' = @($skuAssignments | Where-Object Source -eq 'Direct').Count
        'Group Assignments' = @($skuAssignments | Where-Object Source -eq 'Group').Count
        'Direct and Group Assignments' = @($skuAssignments | Where-Object { ([string]$_.HasDirectAndGroup).Trim() -match '^(?i:true|1|yes)$' }).Count
        'Potentially Reclaimable Assignments' = $reclaimable
        'Disabled Reclaimable Assignments' = $disabledReclaimable
        'Dormant 31-90D Assignments' = $dormant
        'E3 to F3 Review Candidates' = $e3ToF3ReviewCandidates
        'Risk Priority' = $priority
        'Risk Signal' = $signal
        'Recommended Action' = $action
        'Snapshot Date' = $snapshotDate
        'Evidence Status' = 'Observed'
    }
}

$outputDirectory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$evidence | Sort-Object @{ Expression = 'Enabled Units'; Descending = $true }, 'License Product' | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8NoBOM
$optimizationRows | Sort-Object 'Optimization Opportunity', 'Microsoft 365 Plan', 'User Principal Name' | Export-Csv -LiteralPath $OptimizationOutputPath -NoTypeInformation -Encoding utf8NoBOM

[pscustomobject]@{
    OutputPath = $OutputPath
    Products = $evidence.Count
    GovernedProducts = @($evidence | Where-Object 'Capacity Class' -eq 'Governed capacity').Count
    GovernedEnabledUnits = ($evidence | Where-Object 'Capacity Class' -eq 'Governed capacity' | Measure-Object 'Enabled Units' -Sum).Sum
    GovernedConsumedUnits = ($evidence | Where-Object 'Capacity Class' -eq 'Governed capacity' | Measure-Object 'Consumed Units' -Sum).Sum
    PotentiallyReclaimableAssignments = ($evidence | Measure-Object 'Potentially Reclaimable Assignments' -Sum).Sum
    E3ToF3ReviewCandidates = ($evidence | Measure-Object 'E3 to F3 Review Candidates' -Sum).Sum
    OptimizationOutputPath = $OptimizationOutputPath
} | Format-List

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA5kNtbfC8t8jvS
# tOdUhrJJFOgmCuskfE/VruHBR+DRmKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILkfZ+yZzxdrsZ0rnutyFKCuasCeNmrXqVTHw3+/lcqqMA0GCSqG
# SIb3DQEBAQUABIIBgH4FklM2WRfG3EjHUAI6vwwQhuhlqLzGZosDIYRqG/mUEupQ
# Gml8O1GYHGrSeRFR2irHiAiFtEQ1HnzxgyfV3NHMn7TDrN3XktsCImhHgO6cYQgI
# BelZvnJ0Jc2+Sm7aMxuX0s8vgBc03tZYqVikVFq6ZewHVAS8U599pifO0TeO7Z4B
# T/BduMVnueA1SrsxY8lOep+Eg6EWC12TKjuLER7IePlAS98teMaS+XkYm8dZmX2n
# jU1vyoPSf2CZ57OVwepxALYHtdhlvTjRKaAcOfwcNjCM2cmiMFamkosQlMJy3XGx
# l6B1nHQuArc1gXdcfjUSvCQAGR5tdTny4bc02xsw7G6cAw9PTbrk1L4en5g2d0SA
# A86k6PnigkE9DHCQeLMT8VN4LZCdrmtkMM7uyHTwkdr4t25PdCKwQP9w3+d5NQb/
# MPsA6UCTQdfZ3P5fEUQH7A3OJroBiQjJAKnf/OXHnZ/AZFUKa2yMp5B32XWI4Ra4
# NGuMYigc2+oRq6L0JKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjkxOTEx
# MzBaMC8GCSqGSIb3DQEJBDEiBCDRxF/2DH2LaJW3grnkAtcSfGVQePSeKNL1IGLN
# dXeRYjANBgkqhkiG9w0BAQEFAASCAgBCfoCkG1tSzZ1goR44Px2JKW9tCxDcg0XJ
# JEyzTqi2s9Rj6p2olrPzpKtKK6FqzpQfKUxjfghLnee+0Ysn0lKtb+cZlvfxKeeW
# rrBOrF4UFu/VU2j+av6uBXlFY/NWFooaeacaFitGcmOg0uOLqUCyVA23uDaXd7VG
# djUjoGp9GICaXWCsZYuRUYacUsYzCBBOmh2+GjuLFyT3gdz6yLP5ZtG/DwU9pqH8
# DytTsIM7wNEXVIB29k4JHOqeiqAiGiPH+9NnNTtfpkImD79zsPe++qWVpNBJF1ci
# oeZVihSJG0lQj2XS9AnXunuZLbpzSPI8jPw5+claRQvN+t84mHeCAvv20/bFFNZx
# jV5/KB+4jGKMtKChe08Ofg7MRy5yy3ZbE6YRHBS+mM7XXCegeyruPbd74Z4g7n8A
# TR2qiO3UoGXkYN5bQsQhjQ5oI/PAnel3Sb/Yhc0Q002LH2gve9klZs3V+9J5YqxV
# uMMecE7FgRx6NCXbSv9wJssQDwDZQ0EbifEQzl191jOppI3FQkmHUXFwftu2+E5+
# Fu91Twq5d0O9PEBvTI57XfhXkT9k63zJIw5yE6ocGqiPAa9oBQLr1adhTckbjsiV
# St9Q5U3o1rCJjsBkerpcu/F3xY0G0Qy5jd1FHGXrjrRkU7tbhLtFUBdH6rgkK0lX
# o6ScX9QXTw==
# SIG # End signature block
