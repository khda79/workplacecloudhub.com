# Version 1.0.1: an email client is used when Graph names its version (ProPlus, Outlook 2016, Undetermined), not only Yes.
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

function Test-ClientUsed {
    # Graph writes the client version (for example ProPlus, Outlook 2016 or Undetermined) and leaves unused clients empty.
    param([AllowNull()][object]$Value)
    return -not [string]::IsNullOrWhiteSpace([string]$Value)
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
$emailClientUsers = @($emailUsers | Where-Object { $row = $_; @($emailClients | Where-Object { Test-ClientUsed (Get-Field $row $_.Column) }).Count -gt 0 })
$emailRefresh = if ($email.Status -eq 'Observed') { Get-RefreshDate -Rows $email.Rows -Column 'Report Refresh Date' } else { $null }
$emailFreshness = Get-Freshness $emailRefresh
$emailPeriod = if ($email.Status -eq 'Observed') { Get-ReportPeriodLength -Rows $email.Rows } else { $null }
$emailRows = foreach ($client in $emailClients) {
    $users = if ($email.Status -eq 'Observed') { @($emailUsers | Where-Object { Test-ClientUsed (Get-Field $_ $client.Column) }).Count } else { $null }
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCncUhMslujW4ni
# cTAfT2vxXHU0EfeBORxSltH3LXAB5KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICuRnBtidmy8rEnCpXQ8MKg8wA/1kTwCQsD9LF2aNl+kMA0GCSqG
# SIb3DQEBAQUABIIBgEyL5Cb6NHXG4cpyhrSTkdfnEuxCo46vceD3wvSIhGbNIqx8
# 9ifz02IOyJe8p4sL/ye82rhiBYQQm4FDT1iQaHI7Cq7vHgkKp5ZwEd3UjI5/Tb1/
# EpSFkfSUJnnT+s9VdWVbu5FjZdndORdvAxXzv3h4LsNruQQjJMZXNpmEdM13KO3v
# +9fXnbV1HMzawVSo7mq+7z47zQnza7D3MuQ9BKXkHNeGpgLsAK7txwWX9oqIIt7U
# 3IAY07j2jvC2k59LOaYvAELOPOopdDa64WZ6aeTqMHOoxAubSCq+H3votl0YLd1y
# rpawbCBflqsGvInie8BE0y7vTlB789Fe9OMfvpEl+dTd1naoLIPXG6XvSnYw1H5s
# eQ+G5Wy2Q4yQYSyuXWw7PZYeCXE4RFRvxkCPNxcHpbqi+Jb38D60hnQ+lvQrbbPO
# zU/Sz5OM4uG1iCzcRVQfQY1SK2ShwpKIJi1FTB6VF5QH8S+Ze9sd9lsj+BYrRqQp
# riRx0IKxuBg1x5azLKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgwOTE4
# MDZaMC8GCSqGSIb3DQEJBDEiBCDTrrEDvD5N41DHSxnaVMZXqCKP4mMk9cfw/VaJ
# OPhQIjANBgkqhkiG9w0BAQEFAASCAgBv+mH0KN3kFup17otZDRu4x7rdkKZ7bBXa
# xOQd082Gp9hTsnQ1mMKUhvr2v+WAPJjJCWwrvZNIBhqCeFoJI0UxTOxbbsd8z0WE
# anZkwithDLfE2H3S5McRWjpPFE6QtiYShTQGzU+6QIlqcaCiI+/oKLA0baamq4Jc
# 0rWC14alO0LU2vOv58CVPRoOAe64R1LeiX11x1WHsjvFA1bHe0V3iN+dqdlVTgEV
# sd1HMmLSdV8JTEu8dA4S02BfZKNUpYquptEDU/Q8kLvFULHKRk/leSp20bKWYC4q
# ZEFMNHXfYtC83zAA4TQLrPyAbcD2hFSyLZKZwUDWRChSFdxfcm41g6GL8mke3CSy
# O4IV7SGqtOiD2L9XmSTk6lnr5WQ7uYMnuths0OvQ6Dr26mV0aPfclVfWFRuf8ERu
# uNJbwoxXhP8fSOvuO5M1jrANH2OGr8h4Os/y00TqYByw15TCcKQoJjEsRHHZIBoT
# Aom6xYicF69xbeTYr0zQbD4vhNXdT7hKcLNoItDhc4FhXYC+S6KS0FX1YgfT+ixp
# gYrL7hXYzvBRl4Qjvwf/fQXuabg1a1E/PFnffvP3wF6bwvq+bQOJtHQUkIXCXQ4c
# VwkFSRT5A+IitosMgQ4xOZDrMeroW+O9mml8jcsA6H2nNUSZQW0jl0oHn3HDYjpb
# PDRRxnL4cg==
# SIG # End signature block
