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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCErFF3z/PJgLlV
# u/gVpkkXv1l/dVuZtl99sviBVjbYwKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIE10F0eNF2gxX02jIawA/bNiGO5o5Aj6/Pa1EmBYKoMWMA0GCSqG
# SIb3DQEBAQUABIIBgIP1Q36cFTxWZFvK6FEd73x1pw6rdgcZhmt+JEyxh+YxFn3X
# Rkbdo2l9y9Ai/o548Jw6FVMFUU7mr540wuvof31Kg/5XX6KTd9WYtnov51d11iqi
# cyK1V75usE/x12AceP5feSIcud5fD7G9jqmAEMQIc411NsM2LCvcZGl/XcRGgBUv
# Z0yoteDOYblPcF4K3lcvv7zJXp+VAyHj1hQgs04yJLF+4ffesvLTcNu1wI5TZTSs
# E4VDet+wxBwNMjp2puiwCeJSxBnunNLwX+tdjdNGIIWUKDACbQuEo1jBpJLlw1Ah
# Ib8gV0309RCKBlmUtKABwp5v2tmSOyL4dh7ES6eOOYFwRJkLBwBX1XIN1gbhX1eZ
# mFef9aM644bxP84uX0Cll9lTU94ccnhqaaAcXtqU5QjPS8n6PjscXIKhMDGfxDGC
# YGqdDpCr1nyQz6m9Ey5oMwLpvgnKbk1CPfd6BI97hsb69wHOBTjoe/PNe6zRRLGl
# Otengerxr/ig8BcHA6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjYyMzUw
# MDBaMC8GCSqGSIb3DQEJBDEiBCC24YtvVXTVKvCa9mRzCZxpuz/a+W23ANJZPkuu
# 78whETANBgkqhkiG9w0BAQEFAASCAgBCkAeOt5LWprlF/RgSvvKOj+26W/HsfuP7
# qft5AmTvVbdZylq9JlFv+PPuxtq1DS9Jp7SN7q8MonZx2t+CqbuK7IDFW07utHju
# vUT/YrwbdOCJxW4hSj9mT6RlfDWmsoFGGEdJeS+rCN7u8avQ1eRvqsKDSvweU59t
# IvswJASxwek4+47vc/8+VWC1ab8uoEcs1KnmMbmmRu51tEb9D28WkvoWflKVP9QS
# 70/qYg13dszfc897KItB+SyFCrtQ8WGnbTCv/ZDjgaUz2Er2fgI2VBLWrbj8+thm
# H/lN6WXqyzlSnFrl0FV9uLgMdrRY0dIfqipSP2U6+5/lJzhktU10RY3z7ffTzfPp
# /3oeogp6ymSR0csLiUddOzgKMX6pQJD3KAYvfwXAeFs12BGAWSc1m1CajQZNTsta
# CUkAns+WItfufhIllQk5eWdSbz/AtW/qywEV3rPT5k25H4NUgGjQEOHWxkFdQ+ip
# UQMBbMd9lPxM4YULU+ypon3lIn8bqg/ks7nzdAoqCo3jS3rslAbcNTqy7W5KiIqE
# udRJECvT+Nu9NCXX3oRAydzx0YEPQ1pSh4IuTKh4AYbJBpPCu+HKoYOZ3whixJ/O
# xgJCG78ASe3GlQZ4iJ0FOXPFTg/0TRQGLzAA6qX/j8P9eEfFU4JU4EoW8V8wv6r2
# vi8a2/MAIw==
# SIG # End signature block
