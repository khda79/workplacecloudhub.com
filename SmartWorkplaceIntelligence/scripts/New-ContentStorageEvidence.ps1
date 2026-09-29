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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBdlj/o5VN/eGni
# lC/Gsz9kmKC7EYVczg3mX5fFLFYVjqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDGV08T2B6CECe2VqNPkw9wBykT1UgngsVgnobpsHlfhMA0GCSqG
# SIb3DQEBAQUABIIBgEnsdnw/7hZ6DUEG3B5qEp3VhJZ0Ar4qj8gx5622xTaQe9Ag
# oJBgAg9tMPmcLY+xmjQIgrcQymKGB4kXK4IdfDQDLv/boMumKqYM086HuAVlzp6f
# fk2vuuP2vk8w55nFlOm5IZmfxjnRngaRSo8Kl3xF74s24JEriOZHudmQkYKlu1q8
# HJmWj4RmhN0pyqamLELs7KDf1OqA9sHJvT4vaC3F5By4uUtuXp6wB2PLjIleFS/i
# +zLauN+vuQfurMELp97UdBqhIXHloCLBn6Dpy0ktUmJbyxZ8bTqUI/TmBDWzIzpI
# +eMR0x+H9fuJo9ak1v1g1/PRTjVMmwcK0dEvHm62GNBCZ3IxS4Kf7eRh+a4Yzuw5
# ZheDw617z5V5qQVzgPZNuDgR1tBRWlsnvzGgOHK438SG6FJ1s9/Nc9ppPtm8Xo9T
# 1DOUjDIpiAVxJo0crfm+2ybi/f4I0SHK4g1XRXnHdEBfBqR93hRvzV725OrpdmfS
# c9PlUsZjdiqbo4kPt6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjkxOTEx
# MjlaMC8GCSqGSIb3DQEJBDEiBCDEHJ+s9GdhhILZfvH5Emh7OcngpCnJeWfRjgNI
# PBpoSDANBgkqhkiG9w0BAQEFAASCAgB236aR5SgFSN8v3l/tsHRMOeLqR+1bihwy
# 3bQeR+A+GKZjV6VdqUnM5rNCSDDiSudGWyf5O3xfemM9U8CgfRexkJgQFtuOx6tR
# 9Y0MuPCMR56lsXRzjRMFfSgvbOPZrjoSXlekNMcC7ehGOdR5vj6pbbdxPVFqjvzC
# upMF/S4DP/aInNIpdUsmwiXF1BCCbB7KDLirm6MaiOJxwNHydwmptO2dHT7wu2Bg
# 1D5agiEVt4Kb/NSVoqfAwRl0G/C1fS7AyHpdhbXb1yyI9vau4bY4j9azkzdrCKGl
# VTIpxpg2JdE4ljDp2dBfqXJbgJhTjvkgbRofytKBQRB6DZ558sD+Xwmu5llqrMI3
# OR/NuQE7i2sK4GX+oOG3h6ICllZrkuo1F3ofDSjn0R3XUBE1cRyI5LF9LlUEsmbW
# edQQ0FlttMTc/lHoeIxDIPFLNKwqWmL7XTd3RTU/gGxKVSAGAgN1sUt7teXX8gZI
# diETFXotsVovyyNJiHi7pWG1o+HBcybuip/nLXLFVKw1FrqHN5W7orPX7assm9up
# 6maHC4Jisbyb2JQtNpTCjwSB5OP8hufMAGDShlr7l+L17JvFia+SVpwCnHp3NZQc
# HWzHh6QItjsxNaQnqy9/1bACERU2OndOgXe92nLsLHigHaSuy9qbUBRjJsZ4FuPU
# zbFpRSBOpA==
# SIG # End signature block
