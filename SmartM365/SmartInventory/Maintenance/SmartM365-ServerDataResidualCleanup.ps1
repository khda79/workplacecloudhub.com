<#
.SYNOPSIS
Retires one approved repair backup, the former CMDB lock, and named legacy DATA-LAST files.

.DESCRIPTION
Preview creates a hash-pinned file plan outside DATA. Execute requires that plan
and repeats all checks before removing files. The repair audit and all repaired
outputs must remain verifiable; one retired Discovered Apps weekly detail may be
covered by its SHA256-identical SummaryHistory copy. No other DATA is removed.
#>
#requires -Version 7.2
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DataRootPath,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$EvidenceDirectory,
    [Parameter(Mandatory)][ValidatePattern('^\d{8}T\d{9}Z-[a-fA-F0-9]{8}$')][string]$BackupFolderName,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ExpectedAuditTenantKey,
    [Parameter(Mandatory)][ValidateRange(1,10000)][int]$ExpectedRepairFiles,
    [Parameter(Mandatory)][ValidateRange(0,10000)][int]$ExpectedSummaryFallbacks,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string[]]$ApprovedDataLastNames,
    [string]$ReportDirectory = '',
    [string]$PlanPath = '',
    [switch]$Execute
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$startedAt = Get-Date
$tenantContextPath = Join-Path $PSScriptRoot '..\..\Config\SmartM365-TenantContext.ps1'
if (Test-Path -LiteralPath $tenantContextPath -PathType Leaf) {
    . $tenantContextPath
    Write-SmartM365StartupBanner
}

$root = [IO.Path]::GetFullPath($DataRootPath).TrimEnd('\', '/')
if (-not (Test-Path -LiteralPath $root -PathType Container) -or
    -not [string]::Equals((Split-Path $root -Leaf), $Tenant, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'DataRootPath must be an existing tenant directory whose leaf matches -Tenant.'
}
if ((Get-Item -LiteralPath $root -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
    throw 'A reparse-point DATA root is not supported.'
}
$rootPrefix = $root + [IO.Path]::DirectorySeparatorChar
function Assert-OutsideData([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    if ($full.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or
        [string]::Equals($full, $root, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path must be outside the DATA root: $full"
    }
    return $full
}
function ConvertTo-SafeRelative([string]$Relative) {
    $value = $Relative.Replace('/', '\')
    if ($value -notmatch '^DATA-ALL\\' -or $value -match '(^|\\)\.{1,2}(\\|$)' -or
        $value -match '[\x00-\x1f:*?"<>|]' -or $value.EndsWith('\')) {
        throw "Invalid repair source path: $Relative"
    }
    return $value
}
function Get-PlainFile([string]$Relative) {
    if ($Relative -match '(^|\\)\.{1,2}(\\|$)' -or $Relative -match '[\x00-\x1f:*?"<>|]') {
        throw "Invalid DATA path: $Relative"
    }
    $path = [IO.Path]::GetFullPath((Join-Path $root $Relative))
    if (-not $path.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path escapes DATA root: $Relative"
    }
    $cursor = $path
    while (-not [string]::Equals($cursor, $root, [StringComparison]::OrdinalIgnoreCase)) {
        if (-not (Test-Path -LiteralPath $cursor)) { throw "Missing DATA path: $Relative" }
        if ((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Linked DATA path: $Relative"
        }
        $cursor = Split-Path -Path $cursor -Parent
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Not a DATA file: $Relative" }
    return Get-Item -LiteralPath $path -Force
}
function Get-Hash([object]$File) { (Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256).Hash }
function Add-PlanRow([string]$Category, [string]$Relative, [object]$File) {
    $plan.Add([pscustomobject]@{
        Category = $Category
        RelativePath = $Relative
        Bytes = [int64]$File.Length
        LastWriteTimeUtc = $File.LastWriteTimeUtc.ToString('o')
        SHA256 = Get-Hash $File
    })
}

$evidenceRoot = Assert-OutsideData $EvidenceDirectory
if (-not (Test-Path -LiteralPath $evidenceRoot -PathType Container)) { throw "Evidence directory not found: $evidenceRoot" }
if (-not $ReportDirectory) {
    $ReportDirectory = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ServerDataResidualCleanup_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$reportRoot = Assert-OutsideData $ReportDirectory
New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null
if ($Execute) {
    if (-not $PlanPath) { throw '-Execute requires the plan.csv produced by a previous Preview.' }
    $PlanPath = Assert-OutsideData $PlanPath
    if (-not (Test-Path -LiteralPath $PlanPath -PathType Leaf)) { throw "Plan not found: $PlanPath" }
} elseif ($PlanPath) { throw '-PlanPath is used only with -Execute.' }

$plan = [Collections.Generic.List[object]]::new()
$evidencePath = Join-Path $evidenceRoot 'data-last-old.csv'
if (-not (Test-Path -LiteralPath $evidencePath -PathType Leaf)) { throw "Missing evidence CSV: $evidencePath" }
$lastRows = @(Import-Csv -LiteralPath $evidencePath)
$names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($name in $ApprovedDataLastNames) {
    if ($name -notmatch '^[^\\/:*?"<>|\x00-\x1f]+\.csv$' -or -not $names.Add($name)) {
        throw "Invalid or repeated approved DATA-LAST name: $name"
    }
    $matching = @($lastRows | Where-Object { $_.Name -ceq $name -and $_.Decision -eq 'ReviewOldOutput' })
    if ($matching.Count -ne 1) { throw "Name is not a unique reviewed old output: $name" }
    $row = $matching[0]
    $relative = 'DATA-LAST\' + $name
    $file = Get-PlainFile $relative
    if ($file.Length -ne [int64]$row.Bytes -or
        $file.LastWriteTimeUtc.ToString('o') -cne [string]$row.LastWriteTimeUtc) {
        throw "DATA-LAST file changed since review: $name"
    }
    Add-PlanRow 'OldDataLast' $relative $file
}

$jobsRelative = 'DATA-ALL\Orchestrator\Config\Orchestrator-Jobs.json.txt'
$jobsFile = Get-PlainFile $jobsRelative
$jobs = Get-Content -LiteralPath $jobsFile.FullName -Raw | ConvertFrom-Json -ErrorAction Stop
if (-not $jobs.Jobs) { throw 'The effective orchestrator manifest contains no jobs.' }
$legacy = @($jobs.Jobs | Where-Object {
    $_.Name -eq 'CmdbEvidence-Prepare' -or
    $_.ScriptPath -match 'SmartM365-CmdbEvidence-Prepare\.ps1' -or
    'CmdbEvidence-Prepare' -in @($_.DependsOn)
})
if ($legacy.Count) { throw 'The effective orchestrator manifest still references CMDB preparation.' }
$lockRelative = '.cmdb-preparation.lock'
$lockFile = Get-PlainFile $lockRelative
$handle = [IO.File]::Open($lockFile.FullName, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
$handle.Dispose()
Add-PlanRow 'RetiredCmdbLock' $lockRelative $lockFile

$backupRelative = 'DATA-REPAIR-BACKUPS\' + $BackupFolderName
$backupRoot = Join-Path $root $backupRelative
if (-not (Test-Path -LiteralPath $backupRoot -PathType Container) -or
    (Get-Item -LiteralPath $backupRoot -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
    throw "Repair backup directory is missing or linked: $backupRelative"
}
$auditName = if (Test-Path -LiteralPath (Join-Path $backupRoot 'repair.json.txt') -PathType Leaf) { 'repair.json.txt' }
    elseif (Test-Path -LiteralPath (Join-Path $backupRoot 'repair.json') -PathType Leaf) { 'repair.json' }
    else { throw "Repair audit missing: $backupRelative" }
$auditRelative = Join-Path $backupRelative $auditName
$auditFile = Get-PlainFile $auditRelative
$audit = Get-Content -LiteralPath $auditFile.FullName -Raw | ConvertFrom-Json -ErrorAction Stop
if ($audit.SchemaVersion -ne 1 -or $audit.TenantKey -cne $ExpectedAuditTenantKey -or
    $audit.Status -ne 'Completed' -or @($audit.Files).Count -ne $ExpectedRepairFiles) {
    throw 'Repair audit schema, tenant key, status, or receipt count differs from approval.'
}
$expectedBackupFiles = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
[void]$expectedBackupFiles.Add($auditFile.FullName)
$receipts = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$retainedRows = [Collections.Generic.List[object]]::new()
$fallbackCount = 0
foreach ($entry in @($audit.Files)) {
    $sourceRelative = ConvertTo-SafeRelative ([string]$entry.Relative)
    if (-not $receipts.Add($sourceRelative) -or $entry.Status -ne 'Repaired' -or
        [string]$entry.OriginalSHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or
        [string]$entry.RepairedSHA256 -notmatch '^[A-Fa-f0-9]{64}$') {
        throw "Invalid or repeated repair receipt: $sourceRelative"
    }
    $currentPath = Join-Path $root $sourceRelative
    $currentMatches = $false
    $retainedRelative = $sourceRelative
    if (Test-Path -LiteralPath $currentPath -PathType Leaf) {
        $current = Get-PlainFile $sourceRelative
        $currentMatches = (Get-Hash $current) -eq [string]$entry.RepairedSHA256
    }
    if (-not $currentMatches) {
        if ($sourceRelative -notmatch '^DATA-ALL\\Intune\\Applications\\DiscoveredApps\\WeeklyHistory\\(?<Week>\d{4}-W\d{2})\\Intune_DiscoveredApps_Summary\.csv$') {
            throw "Repaired source is missing or changed: $sourceRelative"
        }
        $summaryRelative = 'DATA-ALL\Intune\Applications\DiscoveredApps\SummaryHistory\' + $Matches.Week + '\Intune_DiscoveredApps_Summary.csv'
        $summary = Get-PlainFile $summaryRelative
        if ((Get-Hash $summary) -ne [string]$entry.RepairedSHA256) {
            throw "Preserved summary hash mismatch: $summaryRelative"
        }
        $retainedRelative = $summaryRelative
        $fallbackCount++
    }
    $retainedRows.Add([pscustomobject]@{RepairSource=$sourceRelative
        RetainedCopy=$retainedRelative;SHA256=[string]$entry.RepairedSHA256})
    foreach ($family in @('originals','replaced')) {
        $relative = Join-Path (Join-Path $backupRelative $family) $sourceRelative
        $file = Get-PlainFile $relative
        if ((Get-Hash $file) -ne [string]$entry.OriginalSHA256) {
            throw "Repair rollback copy hash mismatch: $relative"
        }
        [void]$expectedBackupFiles.Add($file.FullName)
        Add-PlanRow 'RepairBackupFile' $relative $file
    }
}
if ($fallbackCount -ne $ExpectedSummaryFallbacks) {
    throw "SummaryHistory fallback count $fallbackCount differs from expected $ExpectedSummaryFallbacks."
}
$actualDirs = @(Get-ChildItem -LiteralPath $backupRoot -Directory -Recurse -Force)
if (@($actualDirs | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }).Count) {
    throw 'Repair backup contains a linked directory.'
}
$actualBackupFiles = @(Get-ChildItem -LiteralPath $backupRoot -File -Recurse -Force)
if ($actualBackupFiles.Count -ne $expectedBackupFiles.Count -or
    @($actualBackupFiles | Where-Object { -not $expectedBackupFiles.Contains($_.FullName) }).Count) {
    throw 'Repair backup contains unexpected or missing files.'
}
Add-PlanRow 'RepairAudit' $auditRelative $auditFile

$auditReportPath = Join-Path $reportRoot 'repair-audit.json.txt'
Copy-Item -LiteralPath $auditFile.FullName -Destination $auditReportPath -Force
if ((Get-FileHash -LiteralPath $auditReportPath -Algorithm SHA256).Hash -cne (Get-Hash $auditFile)) {
    throw 'The copied repair audit differs from the source audit.'
}
$retainedRows.ToArray() | Export-Csv -LiteralPath (Join-Path $reportRoot 'retained-copies.csv') -NoTypeInformation -Encoding utf8
$actualDirs | ForEach-Object {
    [pscustomobject]@{RelativePath=$_.FullName.Substring($rootPrefix.Length)}
} | Export-Csv -LiteralPath (Join-Path $reportRoot 'backup-directories.csv') -NoTypeInformation -Encoding utf8

$planned = @($plan.ToArray() | Sort-Object Category, RelativePath)
$uniquePaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($item in $planned) {
    if (-not $uniquePaths.Add($item.RelativePath)) { throw "Repeated planned path: $($item.RelativePath)" }
}
if (-not $Execute) {
    $outputPlan = Join-Path $reportRoot 'plan.csv'
    $planned | Export-Csv -LiteralPath $outputPlan -NoTypeInformation -Encoding utf8
    $result = [pscustomobject]@{Mode='Preview';PlannedFiles=$planned.Count;RemovedFiles=0
        RepairReceipts=@($audit.Files).Count;SummaryFallbacks=$fallbackCount
        PlanPath=$outputPlan;ReportDirectory=$reportRoot}
} else {
    $approved = @(Import-Csv -LiteralPath $PlanPath)
    if ($approved.Count -ne $planned.Count) { throw 'Preview plan count differs from current evidence.' }
    $columns = @('Category','RelativePath','Bytes','LastWriteTimeUtc','SHA256')
    for ($i = 0; $i -lt $planned.Count; $i++) {
        foreach ($column in $columns) {
            if ([string]$approved[$i].$column -cne [string]$planned[$i].$column) {
                throw "Preview plan differs from current evidence at row $($i + 1): $column"
            }
        }
    }
    $results = [Collections.Generic.List[object]]::new()
    $deleteOrder = @($planned | Where-Object Category -eq 'OldDataLast') +
        @($planned | Where-Object Category -eq 'RetiredCmdbLock') +
        @($planned | Where-Object Category -in @('RepairBackupFile','RepairAudit'))
    try {
        foreach ($item in $deleteOrder) {
            $file = Get-PlainFile $item.RelativePath
            if ($file.Length -ne [int64]$item.Bytes -or
                $file.LastWriteTimeUtc.ToString('o') -cne $item.LastWriteTimeUtc -or
                (Get-Hash $file) -cne $item.SHA256) { throw "Changed before removal: $($item.RelativePath)" }
            if ($item.Category -eq 'RetiredCmdbLock') {
                $handle = [IO.File]::Open($file.FullName, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
                $handle.Dispose()
            }
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            if (Test-Path -LiteralPath $file.FullName) { throw "File remains after removal: $($item.RelativePath)" }
            $results.Add([pscustomobject]@{Category=$item.Category;RelativePath=$item.RelativePath;Result='Removed'})
        }
        foreach ($directory in @($actualDirs | Sort-Object { $_.FullName.Length } -Descending)) {
            Remove-Item -LiteralPath $directory.FullName -ErrorAction Stop
        }
        Remove-Item -LiteralPath $backupRoot -ErrorAction Stop
        $backupParent = Split-Path -Path $backupRoot -Parent
        if (-not @(Get-ChildItem -LiteralPath $backupParent -Force).Count) {
            Remove-Item -LiteralPath $backupParent -ErrorAction Stop
        }
    } finally {
        $results.ToArray() | Export-Csv -LiteralPath (Join-Path $reportRoot 'results.csv') -NoTypeInformation -Encoding utf8
    }
    $result = [pscustomobject]@{Mode='Execute';PlannedFiles=$planned.Count;RemovedFiles=$results.Count
        RepairReceipts=@($audit.Files).Count;SummaryFallbacks=$fallbackCount
        BackupDirectoryRemoved=(-not (Test-Path -LiteralPath $backupRoot));PlanPath=$PlanPath;ReportDirectory=$reportRoot}
}
$result
if (Get-Command Write-SmartM365CompletionBanner -ErrorAction SilentlyContinue) {
    Write-SmartM365CompletionBanner -Status Success -ScriptName 'SmartM365-ServerDataResidualCleanup' -StartedAt $startedAt -EndedAt (Get-Date)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD+GCXjXfs8eKjf
# TmvsH8ncjsrO4em2yGLBCLf4z/70x6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIM5/NugJGhMkgzV29SUZisnyNqgjxb46WVOotTG+ZX7KMA0GCSqG
# SIb3DQEBAQUABIIBgAA4nL58N8p30lhKEbzHR97G2AONHkXU0ZqIHOy6t+zrvX3i
# ICa9ISWlR/B7TWZspA8E6DJv88j1jcnZ1bduwfCITeNaa8NADiEtvWL8rXHHUrB/
# NJ8UYVcLPqGyOr215ZloLkK/jpvydB3EEhnetiIMciqJZ2c9lmFjXeBXybsnqse8
# qrEfSpQIjochUnOMgx/hb3UvQbWsZRGQf9dALAgPjUHdTNImrcx+VHWZOVtPq9Wm
# 23u+gIUQ5QwloH7TKk7d4NP1ck/TyhF95Szy3aLvVbFusgtgsVb0vEEDvRYbnX3e
# apb8Rnqg2mhgL0sBFL/IlxNLfsXN7+WcQpKU7e1HYk2XskJw826vF2zbSp9LIYpf
# xXtugZYUl/3Qcqn9VyNLn8oInIPWGN/pDFYVWtBgjCn8lqLqGsZLNWYICOt7BlGp
# Y0Jzb8L563ijooj9oz3L/Rw6oTW2/pu/JyMlKE2JkAESeXU7Xc5g2KSKlwRtAw/P
# iN5RpB0+pXJxy0d186GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxOTMz
# NDRaMC8GCSqGSIb3DQEJBDEiBCBVdQPDyOGdHDzMBCxPGn1jrYyb+xp7qYa7AJ4d
# KimiljANBgkqhkiG9w0BAQEFAASCAgCIP7RtLe7u6jh97fbKJabn23YO7CbRsxJz
# EaAoXngycGE3erOhWvFPZpPBq0W+3uf9r2oD5lfcqJ37//Gz7PFOBs9p2AVbiKFq
# 3NJazJ8hg/Vw2kVhVQy689pTT1l8zTB8EpIF4rJ2QtDmYHNZHKUXMs1Q3gbMrmKg
# v4/ZYW3x0EYqNp/hnfRChHV89iuBafP+CfJngK7xFS08fCMrVnSWFoeYcnS6fWN2
# JFqce7sn2VSRcrZfzdrHSy3Zba1cGOPxCqvqyS9AZF0u72V5Lh+aBq7cwp4o5pLn
# lC4fAcYzAB8kTVah0+IOzrfARp2YOW4u7I+ZjtzuG8W+OBGiFlUmhyru2Vw1cLY4
# HtJxh6VDeLBtV+CXeDhtxJ5r2h5Q0kDqqAAPFdUBlpbXYFJ2i8Qlg7f3pPzlHkgN
# ApN8bRxhi9AcoDa6PXaZ3vYrrojWmAuMvLoFiC8i9Ct9vfDVS9pfRi4Cjyp4sWX4
# ErBUFe5fLjlpzsWKU2b5bqDRklAlciBSWjJKRiqqHqwTrqAo69dDXAjRzpJgXkLU
# xq0mF+aW0a+DCa5JjZ+6+GsNJmI4wTtTyuyLugBThf2yDhM9p2Nj/7icTHG3uqVI
# odjlRptJZkShm9jwIjzGTW8t+kHkHTH34Js/oWOIaU0XIz25kh1H7v/M6D05jLpO
# 5aY90Zsg6g==
# SIG # End signature block
