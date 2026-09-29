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
