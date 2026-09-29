[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$AdoptionOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\CollaborationAdoptionEvidence.csv'),
    [string]$ObjectOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\CollaborationObjectEvidence.csv'),
    [string]$TrendOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\CollaborationTrendEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

function Get-SnapshotDate {
    param([Parameter(Mandatory = $true)][object[]]$Rows, [Parameter(Mandatory = $true)][string[]]$CandidateColumns)
    foreach ($column in $CandidateColumns) {
        foreach ($row in $Rows) {
            $value = $row.PSObject.Properties[$column].Value
            $date = Convert-ToDateOrNull $value
            if ($null -ne $date) { return $date }
        }
    }
    return (Get-Date).Date
}

function Get-ActivityState {
    param([AllowNull()][Nullable[datetime]]$LastActivityDate, [Parameter(Mandatory = $true)][datetime]$ReferenceDate)
    if ($null -eq $LastActivityDate) { return 'No activity evidence' }
    $days = [math]::Max(0, ($ReferenceDate - $LastActivityDate).Days)
    if ($days -le 30) { return 'Active 30D' }
    if ($days -le 180) { return 'Dormant 31-180D' }
    return 'Inactive 180D+'
}

function Get-UsageProfile {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$Service,
        [Parameter(Mandatory = $true)][string]$RefreshColumn,
        [Parameter(Mandatory = $true)][string]$ActivityColumn,
        [AllowNull()][string]$EligibilityColumn,
        [AllowNull()][string]$DeletedColumn
    )
    $refreshDate = Get-SnapshotDate -Rows $Rows -CandidateColumns @($RefreshColumn)
    $eligible = @($Rows | Where-Object {
        $isEligible = if ([string]::IsNullOrWhiteSpace($EligibilityColumn)) { $true } else { ([string]$_.PSObject.Properties[$EligibilityColumn].Value) -match '^(?i:yes|true|1)$' }
        $isDeleted = if ([string]::IsNullOrWhiteSpace($DeletedColumn)) { $false } else { ([string]$_.PSObject.Properties[$DeletedColumn].Value) -match '^(?i:yes|true|1)$' }
        $isEligible -and -not $isDeleted
    })
    $active = @($eligible | Where-Object {
        $last = Convert-ToDateOrNull $_.PSObject.Properties[$ActivityColumn].Value
        $null -ne $last -and $last -ge $refreshDate.AddDays(-29)
    })
    [pscustomobject]@{
        Service = $Service
        RefreshDate = $refreshDate
        EligibleUsers = $eligible.Count
        ActiveUsers = $active.Count
        AdoptionRate = if ($eligible.Count -gt 0) { $active.Count / $eligible.Count } else { $null }
    }
}

function Get-ObjectAggregate {
    param([Parameter(Mandatory = $true)][object[]]$Rows, [Parameter(Mandatory = $true)][string]$Service)
    $active = @($Rows | Where-Object { $_.'Activity State' -eq 'Active 30D' }).Count
    $inactive = @($Rows | Where-Object { $_.'Activity State' -in @('Inactive 180D+', 'No activity evidence') }).Count
    $orphaned = @($Rows | Where-Object { $_.'Owner State' -eq 'No owner' }).Count
    $external = @($Rows | Where-Object { $_.'External Access State' -in @('Guests present', 'External sharing enabled') }).Count
    $storage = [double](($Rows | Measure-Object 'Storage Used GB' -Sum).Sum)
    [pscustomobject]@{ Service = $Service; Total = $Rows.Count; Active = $active; Inactive = $inactive; Orphaned = $orphaned; External = $external; Storage = $storage }
}

$lastRoot = Join-Path $DataRoot 'DATA-LAST'
$teams = @(Import-Csv -LiteralPath (Join-Path $lastRoot 'M365_Teams_Teams.csv'))
$teamsUsage = @(Import-Csv -LiteralPath (Join-Path $lastRoot 'M365_Teams_UserActivity.csv'))
$sharePoint = @(Import-Csv -LiteralPath (Join-Path $lastRoot 'M365_SPO_Sites.csv'))
$sharePointUsage = @(Import-Csv -LiteralPath (Join-Path $lastRoot 'M365_SharePoint_UserActivity.csv'))
$oneDrive = @(Import-Csv -LiteralPath (Join-Path $lastRoot 'M365_OneDrive_Usage.csv'))
$copilot = @(Import-Csv -LiteralPath (Join-Path $lastRoot 'M365_Copilot_UserUsage.csv'))

$teamSnapshot = Get-SnapshotDate -Rows $teams -CandidateColumns @('RunDateUtc')
$sharePointSnapshot = Get-SnapshotDate -Rows $sharePoint -CandidateColumns @('RunDateUtc')
$oneDriveSnapshot = Get-SnapshotDate -Rows $oneDrive -CandidateColumns @('Report Refresh Date')

$objects = [System.Collections.Generic.List[object]]::new()
foreach ($row in $teams) {
    $last = Convert-ToDateOrNull $row.LastActivityDateUtc
    $activity = Get-ActivityState -LastActivityDate $last -ReferenceDate $teamSnapshot
    $ownerCount = [int]$row.OwnerCount
    $guestCount = [int]$row.GuestCount
    $ownerState = if ($ownerCount -gt 0) { 'Owner observed' } else { 'No owner' }
    $externalState = if ($guestCount -gt 0) { 'Guests present' } else { 'No guests observed' }
    $priority = if ($ownerCount -eq 0) { 'High' } elseif ($activity -in @('Inactive 180D+', 'No activity evidence')) { 'Medium' } elseif ($guestCount -gt 0) { 'Medium' } else { 'Normal' }
    $signal = if ($ownerCount -eq 0) { 'Team without owner' } elseif ($activity -in @('Inactive 180D+', 'No activity evidence')) { 'Inactive team' } elseif ($guestCount -gt 0) { 'Team with guests' } else { 'No current signal' }
    $action = switch ($signal) {
        'Team without owner' { 'Assign at least two accountable owners or retire the team.' }
        'Inactive team' { 'Confirm business need, then archive or retire the inactive team.' }
        'Team with guests' { 'Review guest access, ownership, and lifecycle policy.' }
        default { 'Maintain ownership and lifecycle monitoring.' }
    }
    $objects.Add([pscustomobject][ordered]@{
        'Object ID' = [string]$row.TeamId; 'Object Name' = [string]$row.TeamDisplayName; Service = 'Teams'; 'Object Type' = 'Team'; Visibility = [string]$row.Visibility
        'Activity State' = $activity; 'Days Since Last Activity' = if ($null -eq $last) { $null } else { [math]::Max(0, ($teamSnapshot - $last).Days) }; 'Last Activity Date' = if ($null -eq $last) { $null } else { $last.ToString('yyyy-MM-dd') }
        'Owner State' = $ownerState; 'Owner Count' = $ownerCount; 'Member Count' = [int]$row.MemberCount; 'Guest Count' = $guestCount; 'External Access State' = $externalState
        'Storage Used GB' = Convert-ToInvariantDecimalText ([double]$row.StorageUsedGB); 'Risk Priority' = $priority; 'Risk Signal' = $signal; 'Recommended Action' = $action
        'Snapshot Date' = $teamSnapshot.ToString('yyyy-MM-dd'); 'Evidence Status' = 'Observed'
    })
}

foreach ($row in $sharePoint) {
    $last = Convert-ToDateOrNull $row.LastActivityUtc
    $activity = Get-ActivityState -LastActivityDate $last -ReferenceDate $sharePointSnapshot
    $ownerState = if ([string]::IsNullOrWhiteSpace([string]$row.Owner)) { 'No owner' } else { 'Owner observed' }
    $externalState = if ($row.ExternalSharingEnabled -match '^(?i:true|1)$') { 'External sharing enabled' } elseif ([string]::IsNullOrWhiteSpace([string]$row.ExternalSharingEnabled)) { 'Not collected' } else { 'No external sharing observed' }
    $priority = if ($ownerState -eq 'No owner') { 'High' } elseif ($activity -in @('Inactive 180D+', 'No activity evidence')) { 'Medium' } elseif ($externalState -eq 'External sharing enabled') { 'Medium' } else { 'Normal' }
    $signal = if ($ownerState -eq 'No owner') { 'Site without owner' } elseif ($activity -in @('Inactive 180D+', 'No activity evidence')) { 'Inactive SharePoint site' } elseif ($externalState -eq 'External sharing enabled') { 'Externally shared site' } else { 'No current signal' }
    $action = switch ($signal) {
        'Site without owner' { 'Assign an accountable owner before lifecycle or sharing decisions.' }
        'Inactive SharePoint site' { 'Confirm business need, then archive or retire the inactive site.' }
        'Externally shared site' { 'Review external sharing, sensitivity, and owner accountability.' }
        default { 'Maintain ownership, activity, and sharing monitoring.' }
    }
    $objects.Add([pscustomobject][ordered]@{
        'Object ID' = [string]$row.SiteUrl; 'Object Name' = if ([string]::IsNullOrWhiteSpace([string]$row.Title)) { [string]$row.SiteUrl } else { [string]$row.Title }; Service = 'SharePoint'; 'Object Type' = [string]$row.Template; Visibility = [string]$row.SharingCapability
        'Activity State' = $activity; 'Days Since Last Activity' = if ($null -eq $last) { $null } else { [math]::Max(0, ($sharePointSnapshot - $last).Days) }; 'Last Activity Date' = if ($null -eq $last) { $null } else { $last.ToString('yyyy-MM-dd') }
        'Owner State' = $ownerState; 'Owner Count' = if ($ownerState -eq 'Owner observed') { 1 } else { 0 }; 'Member Count' = $null; 'Guest Count' = $null; 'External Access State' = $externalState
        'Storage Used GB' = Convert-ToInvariantDecimalText ([double]$row.StorageUsedGB); 'Risk Priority' = $priority; 'Risk Signal' = $signal; 'Recommended Action' = $action
        'Snapshot Date' = $sharePointSnapshot.ToString('yyyy-MM-dd'); 'Evidence Status' = 'Observed'
    })
}

foreach ($row in $oneDrive) {
    $last = Convert-ToDateOrNull $row.'Last Activity Date'
    $activity = Get-ActivityState -LastActivityDate $last -ReferenceDate $oneDriveSnapshot
    $ownerState = if ([string]::IsNullOrWhiteSpace([string]$row.'Owner Principal Name')) { 'No owner' } else { 'Owner observed' }
    $priority = if ($ownerState -eq 'No owner') { 'High' } elseif ($activity -eq 'No activity evidence') { 'Medium' } elseif ($activity -eq 'Inactive 180D+') { 'Medium' } else { 'Normal' }
    $signal = if ($ownerState -eq 'No owner') { 'OneDrive without owner evidence' } elseif ($activity -eq 'No activity evidence') { 'OneDrive without activity evidence' } elseif ($activity -eq 'Inactive 180D+') { 'Inactive OneDrive' } else { 'No current signal' }
    $action = switch ($signal) {
        'OneDrive without owner evidence' { 'Reconcile the account owner before retention or deletion decisions.' }
        'OneDrive without activity evidence' { 'Validate licensing and usage before lifecycle action.' }
        'Inactive OneDrive' { 'Review owner status, retention, and deprovisioning eligibility.' }
        default { 'Maintain activity and retention monitoring.' }
    }
    $storageGb = [double]$row.'Storage Used (Byte)' / 1GB
    $objects.Add([pscustomobject][ordered]@{
        'Object ID' = [string]$row.'Site Id'; 'Object Name' = if ([string]::IsNullOrWhiteSpace([string]$row.'Owner Display Name')) { [string]$row.'Site Id' } else { [string]$row.'Owner Display Name' }; Service = 'OneDrive'; 'Object Type' = 'OneDrive account'; Visibility = 'Private user content'
        'Activity State' = $activity; 'Days Since Last Activity' = if ($null -eq $last) { $null } else { [math]::Max(0, ($oneDriveSnapshot - $last).Days) }; 'Last Activity Date' = if ($null -eq $last) { $null } else { $last.ToString('yyyy-MM-dd') }
        'Owner State' = $ownerState; 'Owner Count' = if ($ownerState -eq 'Owner observed') { 1 } else { 0 }; 'Member Count' = $null; 'Guest Count' = $null; 'External Access State' = 'Not applicable'
        'Storage Used GB' = Convert-ToInvariantDecimalText $storageGb; 'Risk Priority' = $priority; 'Risk Signal' = $signal; 'Recommended Action' = $action
        'Snapshot Date' = $oneDriveSnapshot.ToString('yyyy-MM-dd'); 'Evidence Status' = 'Observed'
    })
}

$teamsProfile = Get-UsageProfile -Rows $teamsUsage -Service 'Teams' -RefreshColumn 'Report Refresh Date' -ActivityColumn 'Last Activity Date' -EligibilityColumn 'Is Licensed' -DeletedColumn 'Is Deleted'
$sharePointProfile = Get-UsageProfile -Rows $sharePointUsage -Service 'SharePoint' -RefreshColumn 'Report Refresh Date' -ActivityColumn 'Last Activity Date' -EligibilityColumn $null -DeletedColumn 'Is Deleted'
$oneDriveProfile = Get-UsageProfile -Rows $oneDrive -Service 'OneDrive' -RefreshColumn 'Report Refresh Date' -ActivityColumn 'Last Activity Date' -EligibilityColumn $null -DeletedColumn 'Is Deleted'
$copilotProfile = Get-UsageProfile -Rows $copilot -Service 'Copilot' -RefreshColumn 'ReportRefreshDate' -ActivityColumn 'LastActivityDate' -EligibilityColumn $null -DeletedColumn $null

$objectRows = @($objects)
$adoption = foreach ($profile in @($teamsProfile, $sharePointProfile, $oneDriveProfile, $copilotProfile)) {
    $aggregate = if ($profile.Service -eq 'Copilot') { [pscustomobject]@{ Total = $profile.EligibleUsers; Active = $profile.ActiveUsers; Inactive = $profile.EligibleUsers - $profile.ActiveUsers; Orphaned = 0; External = 0; Storage = 0 } } else { Get-ObjectAggregate -Rows @($objectRows | Where-Object Service -eq $profile.Service) -Service $profile.Service }
    $age = [math]::Max(0, ((Get-Date).Date - $profile.RefreshDate).Days)
    [pscustomobject][ordered]@{
        Service = $profile.Service; 'Eligible Users' = $profile.EligibleUsers; 'Active Users 30D' = $profile.ActiveUsers; 'Active User Rate (%)' = Convert-ToInvariantDecimalText $profile.AdoptionRate
        'Service Objects' = $aggregate.Total; 'Active Objects 30D' = $aggregate.Active; 'Inactive Objects' = $aggregate.Inactive; 'Orphaned Objects' = $aggregate.Orphaned; 'Objects with External Access' = $aggregate.External
        'Storage Used GB' = Convert-ToInvariantDecimalText $aggregate.Storage; 'Usage Refresh Date' = $profile.RefreshDate.ToString('yyyy-MM-dd'); 'Evidence Age Days' = $age
        'Freshness State' = if ($age -le 14) { 'Current' } elseif ($age -le 45) { 'Aging' } else { 'Stale' }; 'Evidence Status' = 'Observed'
    }
}

$trend = [System.Collections.Generic.List[object]]::new()
function Add-InventoryTrend {
    param([Parameter(Mandatory = $true)][string]$Root, [Parameter(Mandatory = $true)][string]$FileName, [Parameter(Mandatory = $true)][string]$Service)
    foreach ($folder in (Get-ChildItem -LiteralPath $Root -Directory | Sort-Object Name)) {
        $path = Join-Path $folder.FullName $FileName
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $rows = @(Import-Csv -LiteralPath $path)
        if ($rows.Count -eq 0) { continue }
        $snapshot = Get-SnapshotDate -Rows $rows -CandidateColumns @('RunDateUtc', 'Report Refresh Date', 'ReportRefreshDate')
        if ($Service -eq 'Teams') {
            $active = @($rows | Where-Object { [int]$_.InactiveDays -le 30 }).Count
            $inactive = @($rows | Where-Object { [int]$_.InactiveDays -gt 180 }).Count
            $orphaned = @($rows | Where-Object { [int]$_.OwnerCount -eq 0 }).Count
        } else {
            $active = @($rows | Where-Object { $_.IsInactive -notmatch '^(?i:true|1)$' -and [int]$_.DaysSinceLastActivity -le 30 }).Count
            $inactive = @($rows | Where-Object { $_.IsInactive -match '^(?i:true|1)$' }).Count
            $orphaned = @($rows | Where-Object { $_.IsOrphaned -match '^(?i:true|1)$' }).Count
        }
        $trend.Add([pscustomobject][ordered]@{
            'Snapshot Date' = $snapshot.ToString('yyyy-MM-dd'); 'Date Key' = [int]$snapshot.ToString('yyyyMMdd'); 'Snapshot Week' = $folder.Name; Service = $Service
            'Total Objects' = $rows.Count; 'Active Objects 30D' = $active; 'Inactive Objects' = $inactive; 'Orphaned Objects' = $orphaned
            'Eligible Users' = $null; 'Active Users 30D' = $null; 'Adoption Rate (%)' = $null; 'Evidence Status' = 'Observed'
        })
    }
}

Add-InventoryTrend -Root (Join-Path $DataRoot 'DATA-ALL\M365\Teams\Inventory\WeeklyHistory') -FileName 'M365_Teams_Teams.csv' -Service 'Teams'
Add-InventoryTrend -Root (Join-Path $DataRoot 'DATA-ALL\M365\SharePoint\Inventory\WeeklyHistory') -FileName 'M365_SPO_Sites.csv' -Service 'SharePoint'

$copilotHistoryRoot = Join-Path $DataRoot 'DATA-ALL\M365\Usage\Copilot\WeeklyHistory'
foreach ($folder in (Get-ChildItem -LiteralPath $copilotHistoryRoot -Directory | Sort-Object Name)) {
    $path = Join-Path $folder.FullName 'M365_Copilot_UserUsage.csv'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
    $rows = @(Import-Csv -LiteralPath $path)
    if ($rows.Count -eq 0) { continue }
    $profile = Get-UsageProfile -Rows $rows -Service 'Copilot' -RefreshColumn 'ReportRefreshDate' -ActivityColumn 'LastActivityDate' -EligibilityColumn $null -DeletedColumn $null
    $trend.Add([pscustomobject][ordered]@{
        'Snapshot Date' = $profile.RefreshDate.ToString('yyyy-MM-dd'); 'Date Key' = [int]$profile.RefreshDate.ToString('yyyyMMdd'); 'Snapshot Week' = $folder.Name; Service = 'Copilot'
        'Total Objects' = $profile.EligibleUsers; 'Active Objects 30D' = $profile.ActiveUsers; 'Inactive Objects' = $profile.EligibleUsers - $profile.ActiveUsers; 'Orphaned Objects' = 0
        'Eligible Users' = $profile.EligibleUsers; 'Active Users 30D' = $profile.ActiveUsers; 'Adoption Rate (%)' = Convert-ToInvariantDecimalText $profile.AdoptionRate; 'Evidence Status' = 'Observed'
    })
}

foreach ($outputPath in @($AdoptionOutputPath, $ObjectOutputPath, $TrendOutputPath)) {
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) { New-Item -ItemType Directory -Path $outputDirectory | Out-Null }
}

$adoption | Export-Csv -LiteralPath $AdoptionOutputPath -NoTypeInformation -Encoding utf8NoBOM
$objectRows | Export-Csv -LiteralPath $ObjectOutputPath -NoTypeInformation -Encoding utf8NoBOM
@($trend | Sort-Object 'Snapshot Date', Service) | Export-Csv -LiteralPath $TrendOutputPath -NoTypeInformation -Encoding utf8NoBOM

[pscustomobject][ordered]@{
    AdoptionOutputPath = $AdoptionOutputPath; AdoptionRows = @($adoption).Count
    ObjectOutputPath = $ObjectOutputPath; ObjectRows = $objectRows.Count
    TrendOutputPath = $TrendOutputPath; TrendRows = $trend.Count
    CurrentServices = @($adoption | Select-Object Service, 'Eligible Users', 'Active Users 30D', 'Active User Rate (%)', 'Service Objects', 'Inactive Objects', 'Orphaned Objects', 'Evidence Age Days', 'Freshness State')
} | ConvertTo-Json -Depth 4
