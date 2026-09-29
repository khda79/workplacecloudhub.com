[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\DataTrustEvidence.csv'),

    [datetime]$AsOfDate = (Get-Date).Date
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function New-SourceDefinition {
    param(
        [string]$Domain,
        [string]$SourceName,
        [string]$FileName,
        [string[]]$ExpectedColumns,
        [string[]]$SnapshotColumns = @(),
        [bool]$EmptyAllowed = $false
    )

    [pscustomobject]@{
        Domain = $Domain
        SourceName = $SourceName
        FileName = $FileName
        ExpectedColumns = $ExpectedColumns
        SnapshotColumns = $SnapshotColumns
        EmptyAllowed = $EmptyAllowed
    }
}

function Convert-ToDateOrNull {
    param([AllowNull()][object]$Value)

    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $styles = [Globalization.DateTimeStyles]::AssumeLocal
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) { return $parsed }
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::CurrentCulture, $styles, [ref]$parsed)) { return $parsed }
    return $null
}

$sources = @(
    New-SourceDefinition 'Endpoint' 'Intune managed-device inventory' 'Intune_Devices_Inventory.csv' @('Device ID', 'Device name')
    New-SourceDefinition 'Identity' 'Entra active-user inventory' 'M365_Users_Active.csv' @('Object Id', 'User principal name')
    New-SourceDefinition 'Endpoint' 'Entra device inventory' 'M365_Entra_Devices.csv' @('ObjectId', 'DeviceId')
    New-SourceDefinition 'Endpoint' 'AD computer inventory' 'AD_Computers_AllDomains.csv' @('DistinguishedName', 'Name')
    New-SourceDefinition 'Endpoint' 'Windows 11 readiness assessment' 'Intune_Devices_UpgradeEligibility.csv' @('GraphId', 'UpgradeEligibility') @('ExportDateTime')
    New-SourceDefinition 'Endpoint' 'Windows Update status' 'Intune_WindowsUpdate_Status.csv' @('DeviceId', 'ExportDateTime') @('ExportDateTime')
    New-SourceDefinition 'Endpoint' 'Endpoint Analytics device performance' 'Intune_EndpointAnalytics_DevicePerformance.csv' @('DeviceId', 'EndpointAnalyticsScore') @('ReportRefreshDate')
    New-SourceDefinition 'Endpoint' 'Endpoint Analytics startup devices' 'Intune_EndpointAnalytics_StartupDevices.csv' @('DeviceId', 'StartupScore') @('ReportRefreshDate')
    New-SourceDefinition 'Endpoint' 'Entra hardware-ID conflicts' 'M365_Entra_Devices_HardwareIdConflicts.csv' @('HardwareId', 'DeviceCount') @() $true
    New-SourceDefinition 'Endpoint' 'Entra device removal candidates' 'M365_Entra_Devices_RemovalCandidates.csv' @('HardwareId', 'CandidateObjectId') @() $true
    New-SourceDefinition 'Applications' 'Intune discovered applications' 'Intune_DiscoveredApps_Summary.csv' @('AppName', 'AppVersion')
    New-SourceDefinition 'Collaboration' 'Teams inventory' 'M365_Teams_Teams.csv' @('TeamId', 'RunDateUtc') @('RunDateUtc')
    New-SourceDefinition 'Collaboration' 'Teams user activity' 'M365_Teams_UserActivity.csv' @('User Principal Name', 'Report Refresh Date') @('Report Refresh Date')
    New-SourceDefinition 'Content' 'SharePoint site inventory' 'M365_SPO_Sites.csv' @('SiteUrl', 'RunDateUtc') @('RunDateUtc')
    New-SourceDefinition 'Content' 'SharePoint user activity' 'M365_SharePoint_UserActivity.csv' @('User Principal Name', 'Report Refresh Date') @('Report Refresh Date')
    New-SourceDefinition 'Content' 'OneDrive usage' 'M365_OneDrive_Usage.csv' @('Site Id', 'Report Refresh Date') @('Report Refresh Date')
    New-SourceDefinition 'Collaboration' 'Copilot user usage' 'M365_Copilot_UserUsage.csv' @('UserPrincipalName', 'ReportRefreshDate') @('ReportRefreshDate')
    New-SourceDefinition 'Identity' 'M365 user activity' 'M365_Users_Activity.csv' @('UserPrincipalName', 'ReportRefreshDate') @('ReportRefreshDate')
    New-SourceDefinition 'Identity' 'AD user inventory' 'AD_Users_AllDomains.csv' @('DistinguishedName', 'SamAccountName')
    New-SourceDefinition 'Identity' 'AD duplicate UPN findings' 'AD_Users_DuplicateUPN.csv' @('UserPrincipalName') @() $true
    New-SourceDefinition 'Messaging' 'Hybrid identity issues' 'Exchange_HybridIdentity_Issues.csv' @('IssueNumber', 'Potential_Issue') @() $true
    New-SourceDefinition 'Licensing' 'M365 license assignments' 'M365_Licenses_Users.csv' @('User principal name', 'SkuPartNumber')
    New-SourceDefinition 'Licensing' 'Governed Microsoft 365 license prices' 'M365_License_Prices.csv' @('SkuPartNumber', 'Currency', 'UnitPriceMonthly', 'PriceSource') @('CollectedAtUtc')
    New-SourceDefinition 'Security' 'Microsoft Secure Score' 'M365_Security_SecureScore.csv' @('ScoreDate', 'CurrentScore', 'MaxScore', 'ScorePercentage') @('CollectedAtUtc')
    New-SourceDefinition 'Security' 'Authentication methods registration' 'M365_Entra_AuthenticationMethodsRegistration.csv' @('UserId', 'UserPrincipalName', 'IsMfaCapable', 'IsMfaRegistered') @('CollectedAtUtc')
    New-SourceDefinition 'Security' 'Conditional Access policies' 'M365_Entra_ConditionalAccessPolicies.csv' @('PolicyId', 'DisplayName', 'State') @('CollectedAtUtc')
    New-SourceDefinition 'Security' 'Conditional Access policy targets' 'M365_Entra_ConditionalAccessPolicyTargets.csv' @('PolicyId', 'PolicyDisplayName', 'PolicyState', 'TargetType', 'TargetId') @('CollectedAtUtc')
    New-SourceDefinition 'Security' 'AD health checks' 'AD_HealthCheck.csv' @('Forest', 'Domain', 'Category', 'Check', 'Status') @('RunDateUtc')
    New-SourceDefinition 'Security' 'Intune compliance policy state' 'Intune_Devices_Compliance_Policies.csv' @('DeviceName', 'AzureADDeviceId', 'state', 'platformType') @('lastReportedDateTime')
    New-SourceDefinition 'Security' 'Defender agent health' 'Intune_EndpointSecurity_DefenderAgents.csv' @('DeviceId', 'DeviceName', 'LastReportedDateTime', 'RealTimeProtectionEnabled') @('CollectedAtUtc')
    New-SourceDefinition 'Security' 'Firewall health' 'Intune_EndpointSecurity_FirewallStatus.csv' @('DeviceId', 'DeviceName', 'LastReportedDateTime', 'FirewallStatus') @('CollectedAtUtc')
    New-SourceDefinition 'Security' 'Entra Connect synchronization health' 'M365_Entra_AzureADConnect_SyncHealth.csv' @('CheckName', 'Status', 'LastSyncDateTimeUtc') @('ExportDateTime')
    New-SourceDefinition 'Backup' 'Protected mailbox units' 'M365_Backup_ProtectedMailboxes.csv' @('ProtectionUnitId', 'MailboxId') @('ExportDateTime')
    New-SourceDefinition 'Backup' 'Protected SharePoint site units' 'M365_Backup_ProtectedSharePointSites.csv' @('Workload', 'ProtectionUnitId', 'DirectoryObjectId', 'Status') @('CollectedAtUtc') $true
    New-SourceDefinition 'Backup' 'Protected OneDrive account units' 'M365_Backup_ProtectedOneDriveAccounts.csv' @('Workload', 'ProtectionUnitId', 'DirectoryObjectId', 'Status') @('CollectedAtUtc') $true
    New-SourceDefinition 'Backup' 'Mailbox policy-scope coverage' 'M365_BackupPolicyScope_MailboxCoverage.csv' @('MemberId', 'MailboxFoundInInventory') @('RunDateUtc')
    New-SourceDefinition 'Backup' 'Mailbox policy-scope members' 'M365_BackupPolicyScope_GroupMembers.csv' @('MemberId', 'AccountEnabled') @('RunDateUtc')
)

$lastRoot = Join-Path $DataRoot 'DATA-LAST'
$evidence = [System.Collections.Generic.List[object]]::new()

foreach ($source in $sources) {
    $path = Join-Path $lastRoot $source.FileName
    $fileExists = Test-Path -LiteralPath $path -PathType Leaf
    $rows = @()
    $columns = @()
    $snapshotDate = $null
    $snapshotBasis = 'Not available'
    $readError = $null

    if ($fileExists) {
        try {
            $rows = @(Import-Csv -LiteralPath $path)
            if ($rows.Count -gt 0) {
                $columns = @($rows[0].PSObject.Properties.Name)
                foreach ($snapshotColumn in $source.SnapshotColumns) {
                    if ($columns -contains $snapshotColumn) {
                        $snapshotDate = Convert-ToDateOrNull $rows[0].$snapshotColumn
                        if ($snapshotDate) {
                            $snapshotBasis = "Business snapshot: $snapshotColumn"
                            break
                        }
                    }
                }
            }
            else {
                $headerLine = Get-Content -LiteralPath $path -TotalCount 1
                if ($headerLine) {
                    $columns = @((ConvertFrom-Csv -InputObject @($headerLine, $headerLine))[0].PSObject.Properties.Name)
                }
            }
            if (-not $snapshotDate) {
                $snapshotDate = (Get-Item -LiteralPath $path).LastWriteTime
                $snapshotBasis = 'File modified time fallback'
            }
        }
        catch {
            $readError = $_.Exception.Message
        }
    }

    $missingColumns = @($source.ExpectedColumns | Where-Object { $columns -notcontains $_ })
    $schemaState = if (-not $fileExists) { 'Not tested' } elseif ($readError) { 'Unreadable' } elseif ($missingColumns.Count -gt 0) { 'Schema mismatch' } else { 'Qualified' }
    $fileState = if (-not $fileExists) { 'Missing' } elseif ($readError) { 'Unreadable' } elseif ($rows.Count -eq 0 -and $source.EmptyAllowed) { 'Valid empty findings output' } elseif ($rows.Count -eq 0) { 'Unexpected empty' } else { 'Populated' }
    $ageDays = if ($snapshotDate) { [math]::Max(0, [int][math]::Floor(($AsOfDate.Date - $snapshotDate.Date).TotalDays)) } else { $null }
    $freshnessState = if ($null -eq $ageDays) { 'Unknown' } elseif ($ageDays -le 7) { 'Fresh' } elseif ($ageDays -le 30) { 'Aging' } else { 'Stale' }
    $decisionReadiness = if (-not $fileExists -or $readError -or $missingColumns.Count -gt 0 -or ($rows.Count -eq 0 -and -not $source.EmptyAllowed)) { 'Blocked' } elseif ($freshnessState -eq 'Stale') { 'Refresh required' } else { 'Ready' }
    $evidenceStatus = if ($decisionReadiness -eq 'Blocked') { 'Not qualified' } elseif ($freshnessState -eq 'Stale') { 'Observed stale' } else { 'Observed' }
    $recommendedAction = switch ($decisionReadiness) {
        'Blocked' {
            if (-not $fileExists) { 'Restore the missing source export before using dependent KPIs.' }
            elseif ($readError) { 'Repair or recollect the unreadable source export.' }
            elseif ($missingColumns.Count -gt 0) { 'Reconcile schema drift before refreshing dependent evidence.' }
            else { 'Validate the unexpectedly empty source before using dependent KPIs.' }
        }
        'Refresh required' { 'Refresh the source before using it for current operational decisions.' }
        default { 'Maintain the current collection and schema contract.' }
    }

    $evidence.Add([pscustomobject][ordered]@{
        'Source Key' = "$($source.Domain)|$($source.SourceName)"
        'Evidence Domain' = $source.Domain
        'Source Name' = $source.SourceName
        'File Name' = $source.FileName
        'File State' = $fileState
        'Schema State' = $schemaState
        'Expected Columns' = ($source.ExpectedColumns -join ', ')
        'Missing Columns' = ($missingColumns -join ', ')
        'Row Count' = $rows.Count
        'Column Count' = $columns.Count
        'Snapshot Date' = if ($snapshotDate) { $snapshotDate.ToString('yyyy-MM-dd') } else { '' }
        'Snapshot Basis' = $snapshotBasis
        'Evidence Age Days' = $ageDays
        'Freshness State' = $freshnessState
        'Decision Readiness' = $decisionReadiness
        'Evidence Status' = $evidenceStatus
        'Recommended Action' = $recommendedAction
        'Read Error' = if ($readError) { $readError } else { '' }
    })

    $rows = $null
    [gc]::Collect()
}

New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
$evidence | Sort-Object 'Evidence Domain', 'Source Name' | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8

[pscustomobject]@{
    OutputPath = $OutputPath
    Sources = $evidence.Count
    Ready = @($evidence | Where-Object 'Decision Readiness' -eq 'Ready').Count
    RefreshRequired = @($evidence | Where-Object 'Decision Readiness' -eq 'Refresh required').Count
    Blocked = @($evidence | Where-Object 'Decision Readiness' -eq 'Blocked').Count
    Fresh = @($evidence | Where-Object 'Freshness State' -eq 'Fresh').Count
    Aging = @($evidence | Where-Object 'Freshness State' -eq 'Aging').Count
    Stale = @($evidence | Where-Object 'Freshness State' -eq 'Stale').Count
    Missing = @($evidence | Where-Object 'File State' -eq 'Missing').Count
    SchemaIssues = @($evidence | Where-Object 'Schema State' -ne 'Qualified').Count
} | Format-List
