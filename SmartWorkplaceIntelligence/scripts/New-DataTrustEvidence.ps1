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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD92FE9KPKLPwAd
# pgE/YhLPpFBp398+qsAO7t8Qz1e4R6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGyF0plkkqvM/KnCTXaXRP3SRI8B0Bm6s+hGe2gQmqsHMA0GCSqG
# SIb3DQEBAQUABIIBgHvpIEsqte/4Qk1q/I40XXStVykWxNCiKlbqznou7wHx6a6u
# 0xhQqe0/WcBSS4bhExU68te+j+tT/IgOQyZTb+xUyv3NZ8va0zN+ZqiJ+lGCpIXE
# S2/iZTsvfNGipgJRJT0wEyCMlSX07OFxNT5eXdmV/nj51Y7cMtCFgjGK/N6IlCf8
# GSwCAA/FLWL5M8t75ddiNKQ6LwJ4qoNjuCM5+PgLbKvH36sxrK04yi1je6u2RYm4
# MPD0460IOZr3mWdYjmqzpBgBjTeYdHMQkrlzX6xjJLVmcuBgnPRnZ4oSVwFhB2Qr
# KoTCcznPAuDwl6wWlc8mb10syQxj3PBFg77qCwt1exSiEBvTktUwYO2929zNLXI+
# XjqGptp2BOHz86Go5PNE+6OUYmNbxKgYBtRbzENO1zyZCI84J7b4b4MOaPeZkP04
# 7YjiBoOfxJenuniHkjmF6+gQM3oih3wG3m/9w4a37/+An6M/He9kcBWKoeMyP53M
# kIViJu1138UYYJnKr6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjkxOTEx
# MjlaMC8GCSqGSIb3DQEJBDEiBCDCL6tBA5uS2cbvg94I/RdHocbyMsKSklKCYrIq
# h+X/ADANBgkqhkiG9w0BAQEFAASCAgBn1ae6/bq7p+hmpv30HrA4TH+dnfA3qf7c
# 69IPYHNlqumRsF09Az47VCIOuFAI4j4u7kzCHp6ZSIW7ZTprpMcQOaWL7WvFkc4E
# 273l1+4RfEGlpDE8rl/odZz/ijPah6Ec3mLfnhktqbqhPpIH4eM7bZNDdpmiPnrQ
# hklqET1lUA26Expqg1bCubRUexmtmq0M2OWHUY0e8iewPhTeRY1HZj85thmYB30/
# hUhZcbxJhyGt2ZGOaKi23ZXD3NGkoMi+kZgdANdkFZ731j9knNQ+S3mU2pSTJ6g6
# ynujAWdPNYeM+hvxHKBAPIt0z4Ez6TBuwtjK+loYtNsnrzaok66/ahudKvngFYt9
# eCvRMY8SAOtNQfO2EeJAwL3KnfZhQ1rJb6sw2gHC+hGpWueyFC5Kec5Z20UVIR65
# ZK8LrNIE+FGRi6S1jMjtfmb4CtaHgsJ9X/EKvfhjdHxBXdkOpJS2teT8oI7gAtFH
# KlG6Chfk8UoaynaL/SGyA4T1OijjeBfri0mkrp4JelEaUnRrrk4gUlA+vxoEcmC7
# 2gskfqXoEiGuvLCa2ahWdIRCJDnl4s18cdcvFrwcDlPy/OrYCe9QU76RDbbsKLUj
# Lh4zzzfQObSa+nIQKSCrj7PZhPqtPvGkNboLmkF9kT6z4xmJPbKRMBf3dtEW3nty
# atsGpsWmtA==
# SIG # End signature block
