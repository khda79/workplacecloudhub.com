[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [Parameter()]
    [string]$OutputRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private'),

    [Parameter()]
    [datetime]$AsOfDate = [datetime]::UtcNow.Date
)

$ErrorActionPreference = 'Stop'

function Get-NormalizedKey {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Test-TrueValue {
    param([AllowNull()][object]$Value)

    return ([string]$Value).Trim() -match '^(?i:true|1|yes)$'
}

function Get-DateValue {
    param([AllowNull()][object]$Value)

    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    return [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture)
}

function Get-AgeDays {
    param([AllowNull()][Nullable[datetime]]$SnapshotDate)

    if ($null -eq $SnapshotDate) { return $null }
    return [math]::Max(0, [int][math]::Floor(($AsOfDate.Date - $SnapshotDate.Date).TotalDays))
}

function Get-FreshnessState {
    param([AllowNull()][Nullable[int]]$AgeDays)

    if ($null -eq $AgeDays) { return 'Unknown' }
    if ($AgeDays -le 30) { return 'Fresh' }
    return 'Stale'
}

$paths = @{
    Protected = Join-Path $DataRoot 'M365_Backup_ProtectedMailboxes.csv'
    Coverage  = Join-Path $DataRoot 'M365_BackupPolicyScope_MailboxCoverage.csv'
    Members   = Join-Path $DataRoot 'M365_BackupPolicyScope_GroupMembers.csv'
}

foreach ($path in $paths.Values) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required private evidence file was not found: $path"
    }
}

$protected = @(Import-Csv -LiteralPath $paths.Protected)
$coverage = @(Import-Csv -LiteralPath $paths.Coverage)
$members = @(Import-Csv -LiteralPath $paths.Members)

$protectedByKey = @{}
foreach ($row in $protected) {
    foreach ($value in @($row.MailboxId, $row.Mail, $row.UserPrincipalName)) {
        $key = Get-NormalizedKey $value
        if ($key -and -not $protectedByKey.ContainsKey($key)) {
            $protectedByKey[$key] = $row
        }
    }
}

$memberByKey = @{}
foreach ($row in $members) {
    foreach ($value in @($row.MemberId, $row.MemberMail, $row.MemberUserPrincipalName)) {
        $key = Get-NormalizedKey $value
        if ($key -and -not $memberByKey.ContainsKey($key)) {
            $memberByKey[$key] = $row
        }
    }
}

$coverageKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$matchedProtectionUnits = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$evidence = [System.Collections.Generic.List[object]]::new()

foreach ($row in $coverage) {
    $rowKeys = @(
        $row.MemberId,
        $row.PrimarySmtpAddress,
        $row.MailboxUserPrincipalName,
        $row.MemberUserPrincipalName,
        $row.MemberMail
    ) | ForEach-Object { Get-NormalizedKey $_ } | Where-Object { $_ } | Select-Object -Unique

    foreach ($key in $rowKeys) { [void]$coverageKeys.Add($key) }

    $protectedRow = $null
    foreach ($key in $rowKeys) {
        if ($protectedByKey.ContainsKey($key)) {
            $protectedRow = $protectedByKey[$key]
            break
        }
    }

    $memberRow = $null
    foreach ($key in $rowKeys) {
        if ($memberByKey.ContainsKey($key)) {
            $memberRow = $memberByKey[$key]
            break
        }
    }

    $mailboxFound = Test-TrueValue $row.MailboxFoundInInventory
    if ($protectedRow) { [void]$matchedProtectionUnits.Add([string]$protectedRow.ProtectionUnitId) }

    if (-not $mailboxFound) {
        $protectionState = 'Not applicable'
        $reconciliationState = 'Scope member without mailbox'
        $riskPriority = 'Medium'
        $riskSignal = 'Policy scope member has no mailbox inventory record'
        $recommendedAction = 'Reconcile the group member with mailbox inventory or remove the invalid policy-scope member.'
    }
    elseif ($protectedRow) {
        $protectionState = 'Protected'
        $reconciliationState = 'Expected and protected'
        if ($memberRow -and -not (Test-TrueValue $memberRow.AccountEnabled)) {
            $riskPriority = 'Medium'
            $riskSignal = 'Disabled account remains in expected backup scope'
            $recommendedAction = 'Confirm retention and recovery requirements, then remove obsolete disabled accounts from policy scope.'
        }
        else {
            $riskPriority = 'None'
            $riskSignal = 'No current protection gap'
            $recommendedAction = 'Maintain policy coverage and refresh evidence on schedule.'
        }
    }
    else {
        $protectionState = 'Not protected'
        $reconciliationState = 'Expected not protected'
        $riskPriority = 'High'
        $riskSignal = 'Expected mailbox is not present in protected-mailbox evidence'
        $recommendedAction = 'Add the mailbox to a protection policy and verify a successful protection unit.'
    }

    $protectionSnapshot = if ($protectedRow) { Get-DateValue $protectedRow.ExportDateTime } else { $null }
    $policySnapshot = Get-DateValue $row.RunDateUtc
    $evidenceAgeDays = @(
        Get-AgeDays $protectionSnapshot
        Get-AgeDays $policySnapshot
    ) | Where-Object { $null -ne $_ } | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

    $accountState = if (-not $memberRow) { 'Unknown' } elseif (Test-TrueValue $memberRow.AccountEnabled) { 'Enabled' } else { 'Disabled' }
    $entityKey = if ($row.MemberId) { $row.MemberId } elseif ($protectedRow.MailboxId) { $protectedRow.MailboxId } else { $row.PrimarySmtpAddress }
    $primaryAddress = @($row.PrimarySmtpAddress, $row.MailboxUserPrincipalName, $row.MemberMail, $protectedRow.Mail, $protectedRow.UserPrincipalName) | Where-Object { $_ } | Select-Object -First 1
    $displayName = @($row.MailboxDisplayName, $row.MemberDisplayName, $protectedRow.DisplayName, $primaryAddress) | Where-Object { $_ } | Select-Object -First 1
    $mailboxType = @($protectedRow.MailboxType, $row.RecipientTypeDetails, 'Unknown') | Where-Object { $_ } | Select-Object -First 1

    $evidence.Add([pscustomobject][ordered]@{
        'Backup Entity Key' = [string]$entityKey
        'Display Name' = [string]$displayName
        'Primary Address' = [string]$primaryAddress
        'Entity Type' = 'Mailbox'
        'Mailbox Type' = ([string]$mailboxType).Trim()
        'Account State' = $accountState
        'Expected Scope State' = 'In expected scope'
        'Protection State' = $protectionState
        'Reconciliation State' = $reconciliationState
        'Protection Status' = if ($protectedRow) { [string]$protectedRow.Status } else { 'Not observed' }
        'Policy Scope Group' = [string]$row.GroupDisplayName
        'Protection Policy' = if ($protectedRow) { [string]$protectedRow.ProtectionPolicyName } else { '' }
        'Error State' = if ($protectedRow -and ($protectedRow.ErrorCode -or $protectedRow.ErrorMessage)) { 'Error' } else { 'None' }
        'Error Code' = if ($protectedRow) { [string]$protectedRow.ErrorCode } else { '' }
        'Risk Priority' = $riskPriority
        'Risk Signal' = $riskSignal
        'Recommended Action' = $recommendedAction
        'Protection Snapshot Date' = if ($protectionSnapshot) { $protectionSnapshot.ToString('yyyy-MM-dd') } else { '' }
        'Policy Snapshot Date' = if ($policySnapshot) { $policySnapshot.ToString('yyyy-MM-dd') } else { '' }
        'Protection Evidence Age Days' = Get-AgeDays $protectionSnapshot
        'Policy Evidence Age Days' = Get-AgeDays $policySnapshot
        'Evidence Age Days' = $evidenceAgeDays
        'Freshness State' = Get-FreshnessState $evidenceAgeDays
        'Evidence Status' = 'Observed'
    })
}

foreach ($row in $protected) {
    if ($matchedProtectionUnits.Contains([string]$row.ProtectionUnitId)) { continue }

    $rowKeys = @($row.MailboxId, $row.Mail, $row.UserPrincipalName) | ForEach-Object { Get-NormalizedKey $_ } | Where-Object { $_ }
    if ($rowKeys | Where-Object { $coverageKeys.Contains($_) }) { continue }

    $protectionSnapshot = Get-DateValue $row.ExportDateTime
    $evidenceAgeDays = Get-AgeDays $protectionSnapshot
    $primaryAddress = @($row.Mail, $row.UserPrincipalName) | Where-Object { $_ } | Select-Object -First 1

    $evidence.Add([pscustomobject][ordered]@{
        'Backup Entity Key' = [string]$row.MailboxId
        'Display Name' = [string]$row.DisplayName
        'Primary Address' = [string]$primaryAddress
        'Entity Type' = 'Mailbox'
        'Mailbox Type' = if ($row.MailboxType) { [string]$row.MailboxType } else { 'Unknown' }
        'Account State' = 'Unknown'
        'Expected Scope State' = 'Outside expected scope'
        'Protection State' = 'Protected'
        'Reconciliation State' = 'Protected outside expected scope'
        'Protection Status' = [string]$row.Status
        'Policy Scope Group' = ''
        'Protection Policy' = [string]$row.ProtectionPolicyName
        'Error State' = if ($row.ErrorCode -or $row.ErrorMessage) { 'Error' } else { 'None' }
        'Error Code' = [string]$row.ErrorCode
        'Risk Priority' = 'Medium'
        'Risk Signal' = 'Protected mailbox is outside the expected policy scope'
        'Recommended Action' = 'Confirm policy intent and align the protected mailbox with the governed scope.'
        'Protection Snapshot Date' = if ($protectionSnapshot) { $protectionSnapshot.ToString('yyyy-MM-dd') } else { '' }
        'Policy Snapshot Date' = ''
        'Protection Evidence Age Days' = Get-AgeDays $protectionSnapshot
        'Policy Evidence Age Days' = $null
        'Evidence Age Days' = $evidenceAgeDays
        'Freshness State' = Get-FreshnessState $evidenceAgeDays
        'Evidence Status' = 'Observed'
    })
}

New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
$outputPath = Join-Path $OutputRoot 'BackupMailboxEvidence.csv'
$evidence | Sort-Object 'Reconciliation State', 'Primary Address' | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding utf8

$summary = [pscustomobject]@{
    OutputPath = $outputPath
    Rows = $evidence.Count
    ExpectedMailboxes = @($evidence | Where-Object { $_.'Expected Scope State' -eq 'In expected scope' -and $_.'Reconciliation State' -ne 'Scope member without mailbox' }).Count
    ExpectedAndProtected = @($evidence | Where-Object { $_.'Reconciliation State' -eq 'Expected and protected' }).Count
    ExpectedNotProtected = @($evidence | Where-Object { $_.'Reconciliation State' -eq 'Expected not protected' }).Count
    ScopeMembersWithoutMailbox = @($evidence | Where-Object { $_.'Reconciliation State' -eq 'Scope member without mailbox' }).Count
    ProtectedOutsideExpectedScope = @($evidence | Where-Object { $_.'Reconciliation State' -eq 'Protected outside expected scope' }).Count
    DisabledScopeMembers = @($evidence | Where-Object { $_.'Expected Scope State' -eq 'In expected scope' -and $_.'Account State' -eq 'Disabled' }).Count
    ProtectionErrors = @($evidence | Where-Object { $_.'Error State' -eq 'Error' }).Count
    OldestEvidenceDays = ($evidence | Measure-Object 'Evidence Age Days' -Maximum).Maximum
}

$summary | Format-List
