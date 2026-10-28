# Version 1.0.0
# Email client usage, OneDrive user activity, Microsoft 365 group activity and the daily
# Copilot user count trend. Every source is optional: an absent, empty or drifted CSV is
# reported as not collected and never becomes a zero or a failed preparation.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$EmailClientOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\EmailClientUsageEvidence.csv'),
    [string]$OneDriveActivityOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\OneDriveActivityEvidence.csv'),
    [string]$GroupActivityOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\GroupActivityEvidence.csv'),
    [string]$CopilotDailyOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\CopilotDailyAdoptionEvidence.csv'),
    # Optional reference clock for reproducible offline preparation; defaults to the current time.
    [Nullable[datetimeoffset]]$AsOfUtc
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ReferenceDate = if ($null -ne $AsOfUtc) { $AsOfUtc.UtcDateTime.Date } else { (Get-Date).Date }

function Convert-ToInvariantDecimalText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    return ([double]$Value).ToString('0.0########', [Globalization.CultureInfo]::InvariantCulture)
}

function Convert-ToDateOrNull {
    param([AllowNull()][object]$Value)
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) {
        return $parsed.Date
    }
    return $null
}

function Convert-ToCountOrNull {
    param([AllowNull()][object]$Value)
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $parsed = 0L
    if ([long]::TryParse([string]$Value, [Globalization.NumberStyles]::Integer, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) { return $parsed }
    return $null
}

function Get-Field {
    # Strict mode forbids reading an absent column; optional report columns resolve to null.
    param([Parameter(Mandatory = $true)][object]$Row, [Parameter(Mandatory = $true)][string]$Name)
    $property = $Row.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Test-Flag {
    param([AllowNull()][object]$Value)
    return ([string]$Value) -match '^(?i:yes|true|1)$'
}

function Read-OptionalSource {
    # Returns the rows of a usable source, or the reason it is not usable.
    param([Parameter(Mandatory = $true)][string]$FileName, [Parameter(Mandatory = $true)][string[]]$RequiredColumns)
    $path = Join-Path (Join-Path $DataRoot 'DATA-LAST') $FileName
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return [pscustomobject]@{ Rows = @(); Status = 'Not collected' } }
    $rows = @(Import-Csv -LiteralPath $path)
    if ($rows.Count -eq 0) { return [pscustomobject]@{ Rows = @(); Status = 'Not collected' } }
    $columns = @($rows[0].PSObject.Properties.Name)
    $missing = @($RequiredColumns | Where-Object { $columns -notcontains $_ })
    if ($missing.Count -gt 0) { return [pscustomobject]@{ Rows = @(); Status = 'Schema mismatch' } }
    return [pscustomobject]@{ Rows = $rows; Status = 'Observed' }
}

function Get-RefreshDate {
    param([Parameter(Mandatory = $true)][object[]]$Rows, [Parameter(Mandatory = $true)][string]$Column)
    $dates = @($Rows | ForEach-Object { Convert-ToDateOrNull (Get-Field $_ $Column) } | Where-Object { $null -ne $_ })
    if ($dates.Count -eq 0) { return $null }
    return ($dates | Sort-Object -Descending | Select-Object -First 1)
}

function Get-Freshness {
    # Same Current/Aging/Stale bands as the collaboration adoption evidence.
    param([AllowNull()][Nullable[datetime]]$RefreshDate)
    if ($null -eq $RefreshDate) { return [pscustomobject]@{ Age = $null; State = 'Unknown' } }
    $age = [math]::Max(0, ($script:ReferenceDate - $RefreshDate).Days)
    $state = if ($age -le 14) { 'Current' } elseif ($age -le 45) { 'Aging' } else { 'Stale' }
    return [pscustomobject]@{ Age = $age; State = $state }
}

function Get-ReportPeriodLength {
    param([Parameter(Mandatory = $true)][object[]]$Rows)
    foreach ($row in $Rows) {
        $days = Convert-ToCountOrNull (Get-Field $row 'Report Period')
        if ($null -ne $days) { return $days }
    }
    return $null
}

function Get-ActivityState {
    param([AllowNull()][Nullable[datetime]]$LastActivityDate, [Parameter(Mandatory = $true)][datetime]$ReferenceDate)
    if ($null -eq $LastActivityDate) { return 'No activity evidence' }
    $days = [math]::Max(0, ($ReferenceDate - $LastActivityDate).Days)
    if ($days -le 30) { return 'Active 30D' }
    if ($days -le 180) { return 'Dormant 31-180D' }
    return 'Inactive 180D+'
}

$lastRoot = Join-Path $DataRoot 'DATA-LAST'

# Email clients: one stable row per client; legacy protocols are flagged for review.
$emailClients = @(
    @{ Client = 'Outlook for Windows'; Column = 'Outlook For Windows'; Family = 'Outlook'; Legacy = 'No' }
    @{ Client = 'Outlook for Mac'; Column = 'Outlook For Mac'; Family = 'Outlook'; Legacy = 'No' }
    @{ Client = 'Outlook for mobile'; Column = 'Outlook For Mobile'; Family = 'Outlook'; Legacy = 'No' }
    @{ Client = 'Outlook on the web'; Column = 'Outlook For Web'; Family = 'Outlook'; Legacy = 'No' }
    @{ Client = 'Mail for Mac'; Column = 'Mail For Mac'; Family = 'Other client'; Legacy = 'No' }
    @{ Client = 'Other mobile apps'; Column = 'Other For Mobile'; Family = 'Other client'; Legacy = 'No' }
    @{ Client = 'POP3'; Column = 'POP3 App'; Family = 'Legacy protocol'; Legacy = 'Yes' }
    @{ Client = 'IMAP4'; Column = 'IMAP4 App'; Family = 'Legacy protocol'; Legacy = 'Yes' }
    @{ Client = 'SMTP'; Column = 'SMTP App'; Family = 'Legacy protocol'; Legacy = 'Yes' }
)
$email = Read-OptionalSource -FileName 'M365_Email_AppUsage.csv' -RequiredColumns (@('Report Refresh Date', 'User Principal Name', 'Is Deleted') + @($emailClients | ForEach-Object { $_.Column }))
$emailUsers = @($email.Rows | Where-Object { -not (Test-Flag $_.'Is Deleted') })
$emailClientUsers = @($emailUsers | Where-Object { $row = $_; @($emailClients | Where-Object { Test-Flag (Get-Field $row $_.Column) }).Count -gt 0 })
$emailRefresh = if ($email.Status -eq 'Observed') { Get-RefreshDate -Rows $email.Rows -Column 'Report Refresh Date' } else { $null }
$emailFreshness = Get-Freshness $emailRefresh
$emailPeriod = if ($email.Status -eq 'Observed') { Get-ReportPeriodLength -Rows $email.Rows } else { $null }
$emailRows = foreach ($client in $emailClients) {
    $users = if ($email.Status -eq 'Observed') { @($emailUsers | Where-Object { Test-Flag (Get-Field $_ $client.Column) }).Count } else { $null }
    [pscustomobject][ordered]@{
        Client = $client.Client; 'Client Family' = $client.Family; 'Legacy Protocol' = $client.Legacy
        Users = $users; 'Email Client Users' = if ($email.Status -eq 'Observed') { $emailClientUsers.Count } else { $null }
        'User Share (%)' = if ($email.Status -eq 'Observed' -and $emailClientUsers.Count -gt 0) { Convert-ToInvariantDecimalText ($users / $emailClientUsers.Count) } else { $null }
        'Report Period Days' = $emailPeriod; 'Usage Refresh Date' = if ($emailRefresh) { $emailRefresh.ToString('yyyy-MM-dd') } else { $null }
        'Evidence Age Days' = $emailFreshness.Age; 'Freshness State' = $emailFreshness.State; 'Evidence Status' = $email.Status
    }
}

# OneDrive user activity: tenant aggregate per activity over the report period.
$oneDriveMetrics = @(
    @{ Metric = 'Viewed or edited files'; Column = 'Viewed Or Edited File Count' }
    @{ Metric = 'Synced files'; Column = 'Synced File Count' }
    @{ Metric = 'Shared files internally'; Column = 'Shared Internally File Count' }
    @{ Metric = 'Shared files externally'; Column = 'Shared Externally File Count' }
)
$oneDrive = Read-OptionalSource -FileName 'M365_OneDrive_UserActivity.csv' -RequiredColumns (@('Report Refresh Date', 'User Principal Name', 'Is Deleted', 'Last Activity Date') + @($oneDriveMetrics | ForEach-Object { $_.Column }))
$oneDriveUsers = @($oneDrive.Rows | Where-Object { -not (Test-Flag $_.'Is Deleted') })
$oneDriveRefresh = if ($oneDrive.Status -eq 'Observed') { Get-RefreshDate -Rows $oneDrive.Rows -Column 'Report Refresh Date' } else { $null }
$oneDriveFreshness = Get-Freshness $oneDriveRefresh
$oneDrivePeriod = if ($oneDrive.Status -eq 'Observed') { Get-ReportPeriodLength -Rows $oneDrive.Rows } else { $null }
$oneDriveRows = [System.Collections.Generic.List[object]]::new()
foreach ($metric in @($oneDriveMetrics) + @(@{ Metric = 'Any OneDrive activity'; Column = $null })) {
    $users = $null; $files = $null
    if ($oneDrive.Status -eq 'Observed') {
        if ($metric.Column) {
            $counts = @($oneDriveUsers | ForEach-Object { Convert-ToCountOrNull (Get-Field $_ $metric.Column) })
            $users = @($counts | Where-Object { $null -ne $_ -and $_ -gt 0 }).Count
            $files = [long](($counts | Where-Object { $null -ne $_ } | Measure-Object -Sum).Sum)
        } else {
            $users = @($oneDriveUsers | Where-Object { $null -ne (Convert-ToDateOrNull $_.'Last Activity Date') }).Count
        }
    }
    $oneDriveRows.Add([pscustomobject][ordered]@{
        Metric = $metric.Metric; Users = $users; Files = $files; 'Population Users' = if ($oneDrive.Status -eq 'Observed') { $oneDriveUsers.Count } else { $null }
        'User Rate (%)' = if ($oneDrive.Status -eq 'Observed' -and $oneDriveUsers.Count -gt 0) { Convert-ToInvariantDecimalText ($users / $oneDriveUsers.Count) } else { $null }
        'Report Period Days' = $oneDrivePeriod; 'Usage Refresh Date' = if ($oneDriveRefresh) { $oneDriveRefresh.ToString('yyyy-MM-dd') } else { $null }
        'Evidence Age Days' = $oneDriveFreshness.Age; 'Freshness State' = $oneDriveFreshness.State; 'Evidence Status' = $oneDrive.Status
    })
}

# Microsoft 365 groups: one row per live group, with the collaboration object signal vocabulary.
$groups = Read-OptionalSource -FileName 'M365_Groups_Activity.csv' -RequiredColumns @('Report Refresh Date', 'Group Id', 'Group Display Name', 'Is Deleted', 'Owner Principal Name', 'Last Activity Date', 'Group Type', 'Member Count', 'External Member Count')
$teamIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$teamsPath = Join-Path $lastRoot 'M365_Teams_Teams.csv'
$teamsKnown = Test-Path -LiteralPath $teamsPath -PathType Leaf
if ($teamsKnown) { foreach ($team in Import-Csv -LiteralPath $teamsPath) { if (-not [string]::IsNullOrWhiteSpace([string]$team.TeamId)) { $null = $teamIds.Add([string]$team.TeamId) } } }
$groupRefresh = if ($groups.Status -eq 'Observed') { Get-RefreshDate -Rows $groups.Rows -Column 'Report Refresh Date' } else { $null }
# Activity is measured at the report date; the reference clock only stands in when Graph gives none.
$groupReference = if ($null -ne $groupRefresh) { $groupRefresh } else { $script:ReferenceDate }
$groupRows = [System.Collections.Generic.List[object]]::new()
$groupIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($row in @($groups.Rows | Where-Object { -not (Test-Flag $_.'Is Deleted') })) {
    $groupId = [string]$row.'Group Id'
    # Graph can repeat a group; the first live row wins so the table keeps its key.
    if ([string]::IsNullOrWhiteSpace($groupId) -or -not $groupIds.Add($groupId)) { continue }
    $last = Convert-ToDateOrNull $row.'Last Activity Date'
    $activity = Get-ActivityState -LastActivityDate $last -ReferenceDate $groupReference
    $ownerState = if ([string]::IsNullOrWhiteSpace([string]$row.'Owner Principal Name')) { 'No owner' } else { 'Owner observed' }
    $external = Convert-ToCountOrNull $row.'External Member Count'
    $externalState = if ($null -eq $external) { 'Not collected' } elseif ($external -gt 0) { 'External members present' } else { 'No external members observed' }
    $priority = if ($ownerState -eq 'No owner') { 'High' } elseif ($activity -in @('Inactive 180D+', 'No activity evidence')) { 'Medium' } elseif ($externalState -eq 'External members present') { 'Medium' } else { 'Normal' }
    $signal = if ($ownerState -eq 'No owner') { 'Group without owner' } elseif ($activity -in @('Inactive 180D+', 'No activity evidence')) { 'Inactive group' } elseif ($externalState -eq 'External members present') { 'Group with external members' } else { 'No current signal' }
    $action = switch ($signal) {
        'Group without owner' { 'Assign at least two accountable owners or retire the group.' }
        'Inactive group' { 'Confirm business need, then archive or delete the inactive group.' }
        'Group with external members' { 'Review external membership, ownership, and lifecycle policy.' }
        default { 'Maintain ownership and lifecycle monitoring.' }
    }
    $mailboxBytes = Convert-ToCountOrNull (Get-Field $row 'Exchange Mailbox Storage Used (Byte)')
    $siteBytes = Convert-ToCountOrNull (Get-Field $row 'SharePoint Site Storage Used (Byte)')
    $storage = if ($null -eq $mailboxBytes -and $null -eq $siteBytes) { $null } else { ([double]$mailboxBytes + [double]$siteBytes) / 1GB }
    $groupRows.Add([pscustomobject][ordered]@{
        'Group ID' = $groupId; 'Group Name' = [string]$row.'Group Display Name'; 'Group Type' = [string]$row.'Group Type'
        'Teams State' = if (-not $teamsKnown) { 'Unknown' } elseif ($teamIds.Contains($groupId)) { 'Teams-enabled' } else { 'Group only' }
        'Activity State' = $activity; 'Days Since Last Activity' = if ($null -eq $last) { $null } else { [math]::Max(0, ($groupReference - $last).Days) }; 'Last Activity Date' = if ($null -eq $last) { $null } else { $last.ToString('yyyy-MM-dd') }
        'Owner State' = $ownerState; 'Member Count' = Convert-ToCountOrNull $row.'Member Count'; 'External Member Count' = $external; 'External Access State' = $externalState
        'Received Email Count' = Convert-ToCountOrNull (Get-Field $row 'Exchange Received Email Count')
        'SharePoint Active File Count' = Convert-ToCountOrNull (Get-Field $row 'SharePoint Active File Count')
        'Viva Engage Posted Message Count' = Convert-ToCountOrNull (Get-Field $row 'Yammer Posted Message Count')
        'Storage Used GB' = Convert-ToInvariantDecimalText $storage; 'Risk Priority' = $priority; 'Risk Signal' = $signal; 'Recommended Action' = $action
        'Snapshot Date' = $groupReference.ToString('yyyy-MM-dd'); 'Evidence Status' = 'Observed'
    })
}

# Copilot daily trend: one row per report date and app; versions without an app leave no row.
$copilotApps = @(
    @{ App = 'Any app'; Prefix = 'AnyApp' }
    @{ App = 'Microsoft 365 Copilot'; Prefix = 'Microsoft365Copilot' }
    @{ App = 'Teams'; Prefix = 'Teams' }
    @{ App = 'Word'; Prefix = 'Word' }
    @{ App = 'Excel'; Prefix = 'Excel' }
    @{ App = 'PowerPoint'; Prefix = 'PowerPoint' }
    @{ App = 'Outlook'; Prefix = 'Outlook' }
    @{ App = 'OneNote'; Prefix = 'OneNote' }
    @{ App = 'Loop'; Prefix = 'Loop' }
    @{ App = 'Edge'; Prefix = 'Edge' }
    @{ App = 'Copilot Chat'; Prefix = 'CopilotChat' }
    @{ App = 'Copilot Chat (work)'; Prefix = 'CopilotChatWork' }
    @{ App = 'Copilot Chat (web)'; Prefix = 'CopilotChatWeb' }
)
$copilot = Read-OptionalSource -FileName 'M365_Copilot_UserCountTrend.csv' -RequiredColumns (@('ReportVersion', 'ReportRefreshDate', 'ReportDate', 'PromptsSubmitted') + @($copilotApps | ForEach-Object { $_.Prefix + 'EnabledUsers'; $_.Prefix + 'ActiveUsers' }))
$copilotRefresh = if ($copilot.Status -eq 'Observed') { Get-RefreshDate -Rows $copilot.Rows -Column 'ReportRefreshDate' } else { $null }
$copilotFreshness = Get-Freshness $copilotRefresh
$copilotRows = [System.Collections.Generic.List[object]]::new()
$copilotKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($row in $copilot.Rows) {
    $reportDate = Convert-ToDateOrNull $row.ReportDate
    if ($null -eq $reportDate) { continue }
    foreach ($app in $copilotApps) {
        $enabled = Convert-ToCountOrNull (Get-Field $row ($app.Prefix + 'EnabledUsers'))
        $active = Convert-ToCountOrNull (Get-Field $row ($app.Prefix + 'ActiveUsers'))
        if (($null -eq $enabled -and $null -eq $active) -or -not $copilotKeys.Add($reportDate.ToString('yyyyMMdd') + '|' + $app.App)) { continue }
        $copilotRows.Add([pscustomobject][ordered]@{
            'Report Date' = $reportDate.ToString('yyyy-MM-dd'); 'Date Key' = [int]$reportDate.ToString('yyyyMMdd'); App = $app.App
            'Enabled Users' = $enabled; 'Active Users' = $active
            'Active User Rate (%)' = if ($null -ne $enabled -and $null -ne $active -and $enabled -gt 0) { Convert-ToInvariantDecimalText ($active / $enabled) } else { $null }
            'Prompts Submitted' = if ($app.App -eq 'Any app') { Convert-ToCountOrNull $row.PromptsSubmitted } else { $null }
            'Report Version' = [string]$row.ReportVersion; 'Usage Refresh Date' = if ($copilotRefresh) { $copilotRefresh.ToString('yyyy-MM-dd') } else { $null }
            'Evidence Age Days' = $copilotFreshness.Age; 'Freshness State' = $copilotFreshness.State; 'Evidence Status' = 'Observed'
        })
    }
}

$outputs = @(
    @{ Path = $EmailClientOutputPath; Rows = @($emailRows); Columns = @('Client', 'Client Family', 'Legacy Protocol', 'Users', 'Email Client Users', 'User Share (%)', 'Report Period Days', 'Usage Refresh Date', 'Evidence Age Days', 'Freshness State', 'Evidence Status') }
    @{ Path = $OneDriveActivityOutputPath; Rows = @($oneDriveRows); Columns = @('Metric', 'Users', 'Files', 'Population Users', 'User Rate (%)', 'Report Period Days', 'Usage Refresh Date', 'Evidence Age Days', 'Freshness State', 'Evidence Status') }
    @{ Path = $GroupActivityOutputPath; Rows = @($groupRows); Columns = @('Group ID', 'Group Name', 'Group Type', 'Teams State', 'Activity State', 'Days Since Last Activity', 'Last Activity Date', 'Owner State', 'Member Count', 'External Member Count', 'External Access State', 'Received Email Count', 'SharePoint Active File Count', 'Viva Engage Posted Message Count', 'Storage Used GB', 'Risk Priority', 'Risk Signal', 'Recommended Action', 'Snapshot Date', 'Evidence Status') }
    @{ Path = $CopilotDailyOutputPath; Rows = @($copilotRows | Sort-Object 'Date Key', App); Columns = @('Report Date', 'Date Key', 'App', 'Enabled Users', 'Active Users', 'Active User Rate (%)', 'Prompts Submitted', 'Report Version', 'Usage Refresh Date', 'Evidence Age Days', 'Freshness State', 'Evidence Status') }
)
foreach ($output in $outputs) {
    $outputDirectory = Split-Path -Parent $output.Path
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) { New-Item -ItemType Directory -Path $outputDirectory | Out-Null }
    if ($output.Rows.Count -gt 0) {
        $output.Rows | Select-Object $output.Columns | Export-Csv -LiteralPath $output.Path -NoTypeInformation -Encoding utf8NoBOM
    } else {
        # A source that is not collected still publishes its schema.
        [IO.File]::WriteAllText($output.Path, (($output.Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
    }
}

[pscustomobject][ordered]@{
    EmailClientOutputPath = $EmailClientOutputPath; EmailClientStatus = $email.Status; EmailClientUsers = if ($email.Status -eq 'Observed') { $emailClientUsers.Count } else { $null }
    OneDriveActivityOutputPath = $OneDriveActivityOutputPath; OneDriveActivityStatus = $oneDrive.Status
    GroupActivityOutputPath = $GroupActivityOutputPath; GroupActivityStatus = $groups.Status; GroupRows = $groupRows.Count
    CopilotDailyOutputPath = $CopilotDailyOutputPath; CopilotDailyStatus = $copilot.Status; CopilotDailyRows = $copilotRows.Count
} | ConvertTo-Json -Depth 3
