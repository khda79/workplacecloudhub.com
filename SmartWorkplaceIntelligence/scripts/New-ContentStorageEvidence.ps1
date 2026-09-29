[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$ObjectOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\ContentStorageEvidence.csv'),
    [string]$TrendOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\ContentStorageTrendEvidence.csv')
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
            $date = Convert-ToDateOrNull $row.PSObject.Properties[$column].Value
            if ($null -ne $date) { return $date }
        }
    }
    throw "No valid snapshot date was found in columns: $($CandidateColumns -join ', ')."
}

function Get-ActivityState {
    param([AllowNull()][Nullable[datetime]]$LastActivityDate, [Parameter(Mandatory = $true)][datetime]$ReferenceDate)
    if ($null -eq $LastActivityDate) { return 'No activity evidence' }
    $days = [math]::Max(0, ($ReferenceDate - $LastActivityDate).Days)
    if ($days -le 30) { return 'Active 30D' }
    if ($days -le 180) { return 'Dormant 31-180D' }
    return 'Inactive 180D+'
}

function Get-StoragePressure {
    param([Parameter(Mandatory = $true)][double]$UsedGb, [Parameter(Mandatory = $true)][double]$AllocatedGb)
    if ($AllocatedGb -le 0) { return 'Unknown' }
    $utilization = 100 * $UsedGb / $AllocatedGb
    if ($utilization -ge 90) { return 'Critical 90%+' }
    if ($utilization -ge 80) { return 'High 80-90%' }
    return 'Healthy <80%'
}

function Get-FreshnessState {
    param([Parameter(Mandatory = $true)][int]$EvidenceAgeDays)
    if ($EvidenceAgeDays -le 14) { return 'Current' }
    if ($EvidenceAgeDays -le 45) { return 'Aging' }
    return 'Stale'
}

$lastRoot = Join-Path $DataRoot 'DATA-LAST'
$sharePoint = @(Import-Csv -LiteralPath (Join-Path $lastRoot 'M365_SPO_Sites.csv'))
$oneDrive = @(Import-Csv -LiteralPath (Join-Path $lastRoot 'M365_OneDrive_Usage.csv'))
$today = (Get-Date).Date
$sharePointSnapshot = Get-SnapshotDate -Rows $sharePoint -CandidateColumns @('RunDateUtc')
$oneDriveSnapshot = Get-SnapshotDate -Rows $oneDrive -CandidateColumns @('Report Refresh Date')

$objects = [System.Collections.Generic.List[object]]::new()

foreach ($row in $sharePoint) {
    $lastActivity = Convert-ToDateOrNull $row.LastActivityUtc
    $activityState = Get-ActivityState -LastActivityDate $lastActivity -ReferenceDate $sharePointSnapshot
    $usedGb = [double]$row.StorageUsedGB
    $allocatedGb = [double]$row.StorageQuotaGB
    $utilization = if ($allocatedGb -gt 0) { 100 * $usedGb / $allocatedGb } else { $null }
    $pressure = Get-StoragePressure -UsedGb $usedGb -AllocatedGb $allocatedGb
    $ownerState = if ([string]::IsNullOrWhiteSpace([string]$row.Owner)) { 'Missing' } else { 'Observed' }
    $isLargeInactive = $activityState -eq 'Inactive 180D+' -and $usedGb -ge 10

    if ($pressure -eq 'Critical 90%+') {
        $priority = 'Critical'; $signal = 'SharePoint quota at 90% or more'; $action = 'Reclaim storage or extend quota after validating business need.'
    } elseif ($pressure -eq 'High 80-90%') {
        $priority = 'High'; $signal = 'SharePoint quota at 80% or more'; $action = 'Plan cleanup or capacity before the site reaches its quota.'
    } elseif ($ownerState -eq 'Missing') {
        $priority = 'High'; $signal = 'Orphaned SharePoint site'; $action = 'Assign an accountable owner before lifecycle or storage action.'
    } elseif ($isLargeInactive) {
        $priority = 'High'; $signal = 'Large inactive SharePoint site'; $action = 'Validate retention, then archive or remove obsolete content.'
    } elseif ($activityState -eq 'Inactive 180D+') {
        $priority = 'Medium'; $signal = 'Inactive SharePoint site'; $action = 'Confirm business need, retention, and archive or retirement eligibility.'
    } else {
        $priority = 'Normal'; $signal = 'No current signal'; $action = 'Maintain ownership, activity, and storage monitoring.'
    }

    $age = [math]::Max(0, ($today - $sharePointSnapshot).Days)
    $objects.Add([pscustomobject][ordered]@{
        'Content Object Key' = [string]$row.SiteUrl
        'Content Object Name' = if ([string]::IsNullOrWhiteSpace([string]$row.Title)) { [string]$row.SiteUrl } else { [string]$row.Title }
        Service = 'SharePoint'
        'Object Type' = if ([string]::IsNullOrWhiteSpace([string]$row.Template)) { 'Site' } else { [string]$row.Template }
        Owner = [string]$row.Owner
        'Owner State' = $ownerState
        'Activity State' = $activityState
        'Last Activity Date' = if ($null -eq $lastActivity) { $null } else { $lastActivity.ToString('yyyy-MM-dd') }
        'Days Since Last Activity' = if ($null -eq $lastActivity) { $null } else { [math]::Max(0, ($sharePointSnapshot - $lastActivity).Days) }
        'Storage Used GB' = Convert-ToInvariantDecimalText $usedGb
        'Storage Allocated GB' = Convert-ToInvariantDecimalText $allocatedGb
        'Storage Utilization (%)' = Convert-ToInvariantDecimalText $utilization
        'Storage Pressure' = $pressure
        'File Count' = $null
        'Active File Count' = $null
        'Deleted State' = 'Not deleted'
        'Risk Priority' = $priority
        'Risk Signal' = $signal
        'Recommended Action' = $action
        'Snapshot Date' = $sharePointSnapshot.ToString('yyyy-MM-dd')
        'Evidence Age Days' = $age
        'Freshness State' = Get-FreshnessState $age
        'Evidence Status' = 'Observed'
    })
}

foreach ($row in $oneDrive) {
    if ([string]$row.'Is Deleted' -match '^(?i:yes|true|1)$') { continue }
    $lastActivity = Convert-ToDateOrNull $row.'Last Activity Date'
    $activityState = Get-ActivityState -LastActivityDate $lastActivity -ReferenceDate $oneDriveSnapshot
    $usedGb = [double]$row.'Storage Used (Byte)' / 1GB
    $allocatedGb = [double]$row.'Storage Allocated (Byte)' / 1GB
    $utilization = if ($allocatedGb -gt 0) { 100 * $usedGb / $allocatedGb } else { $null }
    $pressure = Get-StoragePressure -UsedGb $usedGb -AllocatedGb $allocatedGb
    $ownerState = if ([string]::IsNullOrWhiteSpace([string]$row.'Owner Principal Name')) { 'Missing' } else { 'Observed' }
    $isLargeInactive = $activityState -eq 'Inactive 180D+' -and $usedGb -ge 10

    if ($pressure -eq 'Critical 90%+') {
        $priority = 'Critical'; $signal = 'OneDrive quota at 90% or more'; $action = 'Reclaim storage or extend quota after validating business need.'
    } elseif ($pressure -eq 'High 80-90%') {
        $priority = 'High'; $signal = 'OneDrive quota at 80% or more'; $action = 'Plan cleanup or capacity before the account reaches its quota.'
    } elseif ($isLargeInactive) {
        $priority = 'High'; $signal = 'Large inactive OneDrive account'; $action = 'Validate owner and retention, then archive or remove obsolete content.'
    } elseif ($activityState -eq 'No activity evidence') {
        $priority = 'Medium'; $signal = 'OneDrive without activity evidence'; $action = 'Refresh or reconcile usage evidence before lifecycle action.'
    } elseif ($activityState -eq 'Inactive 180D+') {
        $priority = 'Medium'; $signal = 'Inactive OneDrive account'; $action = 'Review owner status, retention, and deprovisioning eligibility.'
    } else {
        $priority = 'Normal'; $signal = 'No current signal'; $action = 'Maintain activity, retention, and storage monitoring.'
    }

    $age = [math]::Max(0, ($today - $oneDriveSnapshot).Days)
    $objects.Add([pscustomobject][ordered]@{
        'Content Object Key' = [string]$row.'Site Id'
        'Content Object Name' = if ([string]::IsNullOrWhiteSpace([string]$row.'Owner Display Name')) { [string]$row.'Site Id' } else { [string]$row.'Owner Display Name' }
        Service = 'OneDrive'
        'Object Type' = 'OneDrive account'
        Owner = [string]$row.'Owner Principal Name'
        'Owner State' = $ownerState
        'Activity State' = $activityState
        'Last Activity Date' = if ($null -eq $lastActivity) { $null } else { $lastActivity.ToString('yyyy-MM-dd') }
        'Days Since Last Activity' = if ($null -eq $lastActivity) { $null } else { [math]::Max(0, ($oneDriveSnapshot - $lastActivity).Days) }
        'Storage Used GB' = Convert-ToInvariantDecimalText $usedGb
        'Storage Allocated GB' = Convert-ToInvariantDecimalText $allocatedGb
        'Storage Utilization (%)' = Convert-ToInvariantDecimalText $utilization
        'Storage Pressure' = $pressure
        'File Count' = [long]$row.'File Count'
        'Active File Count' = [long]$row.'Active File Count'
        'Deleted State' = 'Not deleted'
        'Risk Priority' = $priority
        'Risk Signal' = $signal
        'Recommended Action' = $action
        'Snapshot Date' = $oneDriveSnapshot.ToString('yyyy-MM-dd')
        'Evidence Age Days' = $age
        'Freshness State' = Get-FreshnessState $age
        'Evidence Status' = 'Observed'
    })
}

$historyRoot = Join-Path $DataRoot 'DATA-ALL\M365\SharePoint\Inventory\WeeklyHistory'
$historyFolders = @(Get-ChildItem -LiteralPath $historyRoot -Directory | Sort-Object Name)
$historyRows = [System.Collections.Generic.List[object]]::new()
$historyCandidates = [System.Collections.Generic.List[object]]::new()

foreach ($folder in $historyFolders) {
    $path = Join-Path $folder.FullName 'M365_SPO_Sites.csv'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
    $rows = @(Import-Csv -LiteralPath $path)
    if ($rows.Count -eq 0) { continue }
    $historyCandidates.Add([pscustomobject]@{ Folder = $folder; Path = $path; Rows = $rows })
}

$maxHistoryRows = [int](($historyCandidates | Measure-Object { $_.Rows.Count } -Maximum).Maximum)
foreach ($candidate in $historyCandidates) {
    $rows = @($candidate.Rows)
    if ($maxHistoryRows -gt 0 -and $rows.Count -lt [math]::Floor(0.9 * $maxHistoryRows)) { continue }
    $snapshot = Get-SnapshotDate -Rows $rows -CandidateColumns @('RunDateUtc')
    $usedGb = [double](($rows | Measure-Object StorageUsedGB -Sum).Sum)
    $allocatedGb = [double](($rows | Measure-Object StorageQuotaGB -Sum).Sum)
    $inactive = @($rows | Where-Object { $_.IsInactive -match '^(?i:true|1)$' }).Count
    $orphaned = @($rows | Where-Object { $_.IsOrphaned -match '^(?i:true|1)$' }).Count
    $atPressure = @($rows | Where-Object { [double]$_.StorageQuotaGB -gt 0 -and (100 * [double]$_.StorageUsedGB / [double]$_.StorageQuotaGB) -ge 80 }).Count
    $historyRows.Add([pscustomobject][ordered]@{
        'Snapshot Date' = $snapshot.ToString('yyyy-MM-dd')
        'Date Key' = [int]$snapshot.ToString('yyyyMMdd')
        'Snapshot Week' = $candidate.Folder.Name
        'SharePoint Sites' = $rows.Count
        'SharePoint Storage Used GB' = Convert-ToInvariantDecimalText $usedGb
        'SharePoint Storage Allocated GB' = Convert-ToInvariantDecimalText $allocatedGb
        'SharePoint Storage Utilization (%)' = Convert-ToInvariantDecimalText $(if ($allocatedGb -gt 0) { 100 * $usedGb / $allocatedGb } else { $null })
        'Inactive SharePoint Sites' = $inactive
        'Orphaned SharePoint Sites' = $orphaned
        'Sites at 80%+ Quota' = $atPressure
        'Evidence Status' = 'Observed'
    })
}

foreach ($outputPath in @($ObjectOutputPath, $TrendOutputPath)) {
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }
}

$objectRows = @($objects)
$trendRows = @($historyRows | Sort-Object 'Snapshot Date')
$objectRows | Export-Csv -LiteralPath $ObjectOutputPath -NoTypeInformation -Encoding utf8NoBOM
$trendRows | Export-Csv -LiteralPath $TrendOutputPath -NoTypeInformation -Encoding utf8NoBOM

$inactiveStorageGb = [double](($objectRows | Where-Object { $_.'Activity State' -eq 'Inactive 180D+' } | Measure-Object 'Storage Used GB' -Sum).Sum)
$summary = [pscustomobject][ordered]@{
    ObjectOutputPath = $ObjectOutputPath
    ObjectRows = $objectRows.Count
    SharePointSites = @($objectRows | Where-Object Service -eq 'SharePoint').Count
    OneDriveAccounts = @($objectRows | Where-Object Service -eq 'OneDrive').Count
    StorageUsedTB = [math]::Round(([double](($objectRows | Measure-Object 'Storage Used GB' -Sum).Sum) / 1024), 2)
    InactiveStorageTB = [math]::Round(($inactiveStorageGb / 1024), 2)
    ContainersAt80PercentQuota = @($objectRows | Where-Object { $_.'Storage Pressure' -in @('Critical 90%+', 'High 80-90%') }).Count
    OrphanedSharePointSites = @($objectRows | Where-Object { $_.Service -eq 'SharePoint' -and $_.'Owner State' -eq 'Missing' }).Count
    OneDriveWithoutActivityEvidence = @($objectRows | Where-Object { $_.Service -eq 'OneDrive' -and $_.'Activity State' -eq 'No activity evidence' }).Count
    TrendOutputPath = $TrendOutputPath
    TrendRows = $trendRows.Count
    TrendFirstWeek = if ($trendRows.Count -gt 0) { $trendRows[0].'Snapshot Week' } else { $null }
    TrendLastWeek = if ($trendRows.Count -gt 0) { $trendRows[-1].'Snapshot Week' } else { $null }
}
$summary | ConvertTo-Json -Depth 3
