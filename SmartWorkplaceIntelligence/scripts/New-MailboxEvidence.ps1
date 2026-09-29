[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$UserEvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\UserInventoryEvidence.csv'),

    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\MailboxEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-NormalizedKey {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Convert-ToDoubleOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $text = ([string]$Value).Trim().Replace([string][char]0x00A0, '').Replace(' ', '')
    $number = 0.0
    $cultures = if ($text.Contains(',') -and -not $text.Contains('.')) {
        @([Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [Globalization.CultureInfo]::InvariantCulture)
    } elseif ($text.Contains(',') -and $text.Contains('.') -and $text.LastIndexOf(',') -gt $text.LastIndexOf('.')) {
        @([Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [Globalization.CultureInfo]::InvariantCulture)
    } else {
        @([Globalization.CultureInfo]::InvariantCulture, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'))
    }
    foreach ($culture in $cultures) {
        if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, $culture, [ref]$number)) { return $number }
    }
    return $null
}

function Convert-ToDateOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$date)) { return $date }
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [Globalization.DateTimeStyles]::AssumeLocal, [ref]$date)) { return $date }
    return $null
}

function Convert-ToInvariantDecimalText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([double]$Value).ToString('0.0', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ListCount {
    param([AllowNull()][object[]]$Values)
    $items = foreach ($value in $Values) {
        if ([string]::IsNullOrWhiteSpace([string]$value)) { continue }
        [string]$value -split '[;|,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
    return @($items | Sort-Object -Unique).Count
}

function Get-DelegationTypes {
    param([AllowNull()][object]$SendAs, [AllowNull()][object]$FullAccess, [AllowNull()][object]$SendOnBehalf)
    $types = @()
    if ((Get-ListCount @($SendAs)) -gt 0) { $types += 'Send As' }
    if ((Get-ListCount @($FullAccess)) -gt 0) { $types += 'Full Access' }
    if ((Get-ListCount @($SendOnBehalf)) -gt 0) { $types += 'Send on Behalf' }
    if ($types.Count -eq 0) { return 'None' }
    return ($types -join '; ')
}

function Get-ExchangeVersion {
    param([AllowNull()][object]$AdminDisplayVersion)
    $value = [string]$AdminDisplayVersion
    if ($value -match 'Version\s+15\.1') { return 'Exchange Server 2016' }
    if ($value -match 'Version\s+15\.2\s+\(Build\s+(\d+)') {
        if ([int]$Matches[1] -ge 2562) { return 'Exchange Server SE' }
        return 'Exchange Server 2019'
    }
    return 'Exchange Server - version not observed'
}

function Get-RiskAssessment {
    param(
        [AllowNull()][Nullable[double]]$MailboxSize,
        [string]$ArchiveState,
        [int]$DelegationCount,
        [AllowNull()][Nullable[datetime]]$LastActivity,
        [string]$ForwardingState,
        [datetime]$AsOfDate
    )
    if ($null -ne $MailboxSize -and $MailboxSize -ge 45 -and $ArchiveState -ne 'Enabled') {
        return @('High', 'Mailbox >=45 GB without an enabled archive', 'Enable or validate the archive and review retention before quota pressure causes disruption.')
    }
    if ($DelegationCount -ge 10) {
        return @('High', 'Mailbox has extensive delegation', 'Review Full Access, Send As, and Send on Behalf assignments with the mailbox owner.')
    }
    if ($null -ne $LastActivity -and ($AsOfDate - $LastActivity).TotalDays -gt 180) {
        return @('Medium', 'Mailbox has no recent observed activity', 'Confirm ownership and retention requirements before deprovisioning or migration action.')
    }
    if ($ForwardingState -eq 'Configured') {
        return @('Medium', 'Mailbox forwarding is configured', 'Validate forwarding destination, business need, and data-loss prevention controls.')
    }
    return @('None', 'No current mailbox attention signal', 'Maintain lifecycle, archive, and delegation governance and refresh evidence on schedule.')
}

$paths = @{
    Exo = Join-Path $DataRoot 'Exchange_EXO_Mailboxes_AllDomains.csv'
    Stats = Join-Path $DataRoot 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv'
    Archive = Join-Path $DataRoot 'Exchange_EXO_Mailboxes_AllDomains_Archive.csv'
    Permissions = Join-Path $DataRoot 'Exchange_EXO_Mailboxes_AllDomains_Permissions.csv'
    OnPrem = Join-Path $DataRoot 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
    Servers = Join-Path $DataRoot 'Exchange_OnPrem_Servers_Inventory.csv'
}
foreach ($path in @($paths.Values) + @($UserEvidencePath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required private evidence file was not found: $path" }
}

$exo = @(Import-Csv -LiteralPath $paths.Exo)
$stats = @(Import-Csv -LiteralPath $paths.Stats)
$archives = @(Import-Csv -LiteralPath $paths.Archive)
$permissions = @(Import-Csv -LiteralPath $paths.Permissions)
$onPrem = @(Import-Csv -LiteralPath $paths.OnPrem)
$servers = @(Import-Csv -LiteralPath $paths.Servers -Delimiter ';')
$users = @(Import-Csv -LiteralPath $UserEvidencePath)
$asOfDate = (Get-Item -LiteralPath $paths.Exo).LastWriteTime.Date

function New-Index {
    param([object[]]$Rows, [string[]]$Properties)
    $index = @{}
    foreach ($row in $Rows) {
        foreach ($property in $Properties) {
            $key = Get-NormalizedKey $row.$property
            if ($key -and -not $index.ContainsKey($key)) { $index[$key] = $row }
        }
    }
    return $index
}

$statsIndex = New-Index $stats @('PrimarySmtpAddress','UserPrincipalName')
$archiveIndex = New-Index $archives @('PrimarySmtpAddress','UserPrincipalName')
$permissionIndex = New-Index $permissions @('PrimarySmtpAddress','UserPrincipalName')
$userIndex = New-Index $users @('User Principal Name')
$serverVersion = @{}
foreach ($server in $servers) {
    $key = Get-NormalizedKey $server.Name
    if ($key) { $serverVersion[$key] = Get-ExchangeVersion $server.AdminDisplayVersion }
}

$rows = [System.Collections.Generic.List[object]]::new()
$sequence = 0
foreach ($mailbox in $exo) {
    $sequence++
    $key = Get-NormalizedKey $mailbox.PrimarySmtpAddress
    if (-not $key) { $key = Get-NormalizedKey $mailbox.UserPrincipalName }
    $statsRow = if ($key -and $statsIndex.ContainsKey($key)) { $statsIndex[$key] } else { $null }
    $archiveRow = if ($key -and $archiveIndex.ContainsKey($key)) { $archiveIndex[$key] } else { $null }
    $permissionRow = if ($key -and $permissionIndex.ContainsKey($key)) { $permissionIndex[$key] } else { $null }
    $userRow = if ($key -and $userIndex.ContainsKey($key)) { $userIndex[$key] } else { $null }
    $size = Convert-ToDoubleOrNull $(if ($statsRow) { $statsRow.TotalItemSizeGB } else { $mailbox.TotalItemSizeGB })
    $archiveSize = Convert-ToDoubleOrNull $(if ($statsRow -and $statsRow.Archive_TotalItemSizeGB) { $statsRow.Archive_TotalItemSizeGB } elseif ($archiveRow) { $archiveRow.Archive_TotalItemSizeGB } else { $null })
    $archiveState = if (([string]$mailbox.ArchiveStatus -match '^(?i:active|enabled)$') -or ($null -ne $archiveSize -and $archiveSize -gt 0)) { 'Enabled' } else { 'Disabled' }
    $sendAs = if ($permissionRow) { $permissionRow.SendAs } else { $mailbox.SendAs }
    $fullAccess = if ($permissionRow) { $permissionRow.FullAccess } else { $mailbox.FullAccess }
    $sendOnBehalf = if ($permissionRow) { $permissionRow.GrantSendOnBehalfTo } else { $mailbox.GrantSendOnBehalfTo }
    $delegations = Get-ListCount @($sendAs, $fullAccess, $sendOnBehalf)
    $delegationTypes = Get-DelegationTypes $sendAs $fullAccess $sendOnBehalf
    $lastActivity = Convert-ToDateOrNull $(if ($statsRow -and $statsRow.LastUserActionTime) { $statsRow.LastUserActionTime } else { $mailbox.LastUserActionTime })
    $forwardingState = if ($mailbox.ForwardingSmtpAddress -or $mailbox.ForwardingAddress) { 'Configured' } else { 'None' }
    $risk = Get-RiskAssessment $size $archiveState $delegations $lastActivity $forwardingState $asOfDate
    $rows.Add([pscustomobject][ordered]@{
        'Mailbox Key' = ('Mailbox {0:D6}' -f $sequence)
        'Primary SMTP Address' = [string]$mailbox.PrimarySmtpAddress
        'Display Name' = [string]$mailbox.DisplayName
        'Country' = if ($userRow -and $userRow.Country) { [string]$userRow.Country } else { 'Unknown' }
        'Hosting Location' = 'Exchange Online'
        'Exchange Version' = 'Exchange Online'
        'Mailbox Type' = if ($mailbox.MailboxType) { [string]$mailbox.MailboxType } else { [string]$mailbox.RecipientTypeDetails }
        'Recipient Type' = [string]$mailbox.RecipientTypeDetails
        'Operational State' = if (([string]$mailbox.AccountEnabled).Trim() -match '^(?i:false|0|no)$') { 'Disabled' } else { 'Active' }
        'Mailbox Size GB' = Convert-ToInvariantDecimalText $size
        'Archive State' = $archiveState
        'Archive Size GB' = Convert-ToInvariantDecimalText $archiveSize
        'Total Storage GB' = Convert-ToInvariantDecimalText $(if ($null -eq $size -and $null -eq $archiveSize) { $null } else { [double]($size ?? 0) + [double]($archiveSize ?? 0) })
        'Delegation Count' = $delegations
        'Delegation Types' = $delegationTypes
        'Forwarding State' = $forwardingState
        'Last Activity Date' = if ($lastActivity) { $lastActivity.ToString('yyyy-MM-dd') } else { '' }
        'Hosting State' = 'Hosted online'
        'Risk Priority' = $risk[0]
        'Risk Signal' = $risk[1]
        'Recommended Action' = $risk[2]
        'Evidence Status' = 'Observed'
    })
}

foreach ($mailbox in $onPrem) {
    $sequence++
    $key = Get-NormalizedKey $mailbox.PrimarySMTPaddress
    if (-not $key) { $key = Get-NormalizedKey $mailbox.UserPrincipalName }
    $userRow = if ($key -and $userIndex.ContainsKey($key)) { $userIndex[$key] } else { $null }
    $serverKey = Get-NormalizedKey $mailbox.ServerName
    $exchangeVersion = if ($serverKey -and $serverVersion.ContainsKey($serverKey)) { $serverVersion[$serverKey] } else { 'Exchange Server - version not observed' }
    $sizeMb = Convert-ToDoubleOrNull $mailbox.'TotalItemSize-In-MB'
    $size = if ($null -ne $sizeMb) { $sizeMb / 1024 } else { $null }
    $archiveSizeMb = Convert-ToDoubleOrNull $mailbox.'ArchiveTotalItemSize-In-MB'
    $archiveSize = if ($null -ne $archiveSizeMb) { $archiveSizeMb / 1024 } else { $null }
    $archiveState = if (([string]$mailbox.ArchiveStatus -match '^(?i:active|enabled)$') -or ($null -ne $archiveSize -and $archiveSize -gt 0)) { 'Enabled' } else { 'Disabled' }
    $delegations = Get-ListCount @($mailbox.SendAs, $mailbox.FullAccess, $mailbox.GrantSendOnBehalfTo)
    $delegationTypes = Get-DelegationTypes $mailbox.SendAs $mailbox.FullAccess $mailbox.GrantSendOnBehalfTo
    $lastActivity = Convert-ToDateOrNull $mailbox.LastLogonTime
    $forwardingState = if ($mailbox.ForwardingSmtpAddress -or $mailbox.ForwardingAddress) { 'Configured' } else { 'None' }
    $risk = Get-RiskAssessment $size $archiveState $delegations $lastActivity $forwardingState $asOfDate
    $rows.Add([pscustomobject][ordered]@{
        'Mailbox Key' = ('Mailbox {0:D6}' -f $sequence)
        'Primary SMTP Address' = [string]$mailbox.PrimarySMTPaddress
        'Display Name' = [string]$mailbox.DisplayName
        'Country' = if ($userRow -and $userRow.Country) { [string]$userRow.Country } else { 'Unknown' }
        'Hosting Location' = 'On-premises'
        'Exchange Version' = $exchangeVersion
        'Mailbox Type' = if (([string]$mailbox.IsShared).Trim() -match '^(?i:true|1|yes)$') { 'SharedMailbox' } elseif (([string]$mailbox.IsResource).Trim() -match '^(?i:true|1|yes)$') { 'ResourceMailbox' } else { 'UserMailbox' }
        'Recipient Type' = [string]$mailbox.RecipientType
        'Operational State' = if (([string]$mailbox.AccountDisabled).Trim() -match '^(?i:true|1|yes)$') { 'Disabled' } else { 'Active' }
        'Mailbox Size GB' = Convert-ToInvariantDecimalText $size
        'Archive State' = $archiveState
        'Archive Size GB' = Convert-ToInvariantDecimalText $archiveSize
        'Total Storage GB' = Convert-ToInvariantDecimalText $(if ($null -eq $size -and $null -eq $archiveSize) { $null } else { [double]($size ?? 0) + [double]($archiveSize ?? 0) })
        'Delegation Count' = $delegations
        'Delegation Types' = $delegationTypes
        'Forwarding State' = $forwardingState
        'Last Activity Date' = if ($lastActivity) { $lastActivity.ToString('yyyy-MM-dd') } else { '' }
        'Hosting State' = 'Hosted on-premises'
        'Risk Priority' = $risk[0]
        'Risk Signal' = $risk[1]
        'Recommended Action' = $risk[2]
        'Evidence Status' = 'Observed'
    })
}

$outputDirectory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$rows | Sort-Object 'Hosting Location', 'Primary SMTP Address' | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8NoBOM

[pscustomobject]@{
    OutputPath = $OutputPath
    Mailboxes = $rows.Count
    ExchangeOnline = @($rows | Where-Object 'Hosting Location' -eq 'Exchange Online').Count
    OnPremises = @($rows | Where-Object 'Hosting Location' -eq 'On-premises').Count
    WithArchive = @($rows | Where-Object 'Archive State' -eq 'Enabled').Count
    WithDelegations = @($rows | Where-Object 'Delegation Count' -gt 0).Count
    HighRisk = @($rows | Where-Object 'Risk Priority' -eq 'High').Count
} | Format-List
