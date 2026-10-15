<#
.SYNOPSIS
Reviews additional SmartM365 server DATA cleanup candidates without deleting files.

.DESCRIPTION
Produces evidence for Usage run duplicates, WorkplaceScope snapshots, generated
Entra removal scripts, old DATA-LAST outputs, prepared repair backups, failed and
retired prepared batches, and the retired CMDB lock. This script is read-only.

.EXAMPLE
./SmartM365-ServerDataSecondPass.ps1 -Tenant prod -DataRootPath '<server DATA root>'
#>
#requires -Version 7.2
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DataRootPath,
    [string]$ReportDirectory = '',
    [string]$EffectiveJobsPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
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
if (-not $ReportDirectory) {
    $ReportDirectory = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ServerDataSecondPass_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$reportRoot = [IO.Path]::GetFullPath($ReportDirectory).TrimEnd('\', '/')
if ($reportRoot.StartsWith(($root + [IO.Path]::DirectorySeparatorChar), [StringComparison]::OrdinalIgnoreCase) -or
    [string]::Equals($reportRoot, $root, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'ReportDirectory must be outside the DATA root.'
}
New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null

function Get-FileHashValue([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }
function Get-TotalBytes([object[]]$Items, [string]$Property) {
    [int64]$total = 0
    foreach ($item in $Items) { $total += [int64]$item.$Property }
    return $total
}
function Get-Week([datetime]$Date) {
    '{0}-W{1:00}' -f [Globalization.ISOWeek]::GetYear($Date), [Globalization.ISOWeek]::GetWeekOfYear($Date)
}
function Get-Relative([string]$Path) { [IO.Path]::GetRelativePath($root, $Path) }
function Test-OwnedPlainPath([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    $prefix = $root + [IO.Path]::DirectorySeparatorChar
    if (-not $full.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return $false }
    $cursor = $root
    foreach ($part in ([IO.Path]::GetRelativePath($root, $full) -split '[\\/]')) {
        $cursor = Join-Path $cursor $part
        $item = Get-Item -LiteralPath $cursor -Force -ErrorAction SilentlyContinue
        if ($item -and ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { return $false }
    }
    return $true
}
function Write-Report([string]$Name, [object[]]$Rows) {
    $path = Join-Path $reportRoot $Name
    if ($Rows.Count) { $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8 }
    else { [IO.File]::WriteAllText($path, '', [Text.UTF8Encoding]::new($false)) }
}

# A run file is redundant only when its exact bytes remain in the same week's
# consolidated history or in a later run of the same report and week.
$usageRows = [Collections.Generic.List[object]]::new()
$usageFolderRows = [Collections.Generic.List[object]]::new()
$usage = Join-Path $root 'DATA-ALL\M365\Usage'
$keepers = @{}
$weeklyHashes = @{}
if (Test-Path -LiteralPath $usage -PathType Container) {
    $runs = @(Get-ChildItem -LiteralPath $usage -Directory -Force |
        Where-Object Name -Match '^\d{8}_\d{6}$' | Sort-Object Name -Descending)
    foreach ($run in $runs) {
        $runDate = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($run.Name, 'yyyyMMdd_HHmmss',
            [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$runDate)) { continue }
        $children = @(Get-ChildItem -LiteralPath $run.FullName -Force)
        $plain = (Test-OwnedPlainPath $run.FullName) -and
            -not @($children | Where-Object { $_.PSIsContainer -or ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) }).Count
        $runRows = [Collections.Generic.List[object]]::new()
        foreach ($file in @($children | Where-Object { -not $_.PSIsContainer })) {
            $match = [regex]::Match($file.Name, '^(?<Base>.+)_(?<Stamp>\d{8}[-_]\d{6})\.csv$',
                [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $decision = 'KeepUnqualified'; $copy = ''; $hash = ''; $week = ''; $base = ''
            if ($plain -and $match.Success) {
                $base = $match.Groups['Base'].Value
                $week = Get-Week $runDate
                $hash = Get-FileHashValue $file.FullName
                $weekly = Join-Path (Join-Path (Join-Path $usage 'WeeklyHistory') $week) ($base + '.csv')
                if ((Test-Path -LiteralPath $weekly -PathType Leaf) -and
                    (Test-OwnedPlainPath $weekly) -and (Get-Item -LiteralPath $weekly).Length -eq $file.Length) {
                    if (-not $weeklyHashes.ContainsKey($weekly)) { $weeklyHashes[$weekly] = Get-FileHashValue $weekly }
                    if ($weeklyHashes[$weekly] -eq $hash) { $decision = 'DuplicateOfWeekly'; $copy = Get-Relative $weekly }
                }
                if ($decision -ne 'DuplicateOfWeekly') {
                    $key = ($week + '|' + $base + '|' + $hash).ToLowerInvariant()
                    if ($keepers.ContainsKey($key)) { $decision = 'DuplicateOfLaterRun'; $copy = $keepers[$key] }
                    else { $decision = 'KeepDistinct'; $keepers[$key] = Get-Relative $file.FullName }
                }
            }
            $row = [pscustomobject]@{Run=$run.Name;Path=(Get-Relative $file.FullName);Bytes=$file.Length
                Week=$week;Base=$base;SHA256=$hash;Decision=$decision;RetainedCopy=$copy}
            $usageRows.Add($row); $runRows.Add($row)
        }
        $eligible = $plain -and $runDate -lt (Get-Date).AddDays(-7) -and
            ($runRows.Count -eq 0 -or -not @($runRows | Where-Object Decision -NotIn @('DuplicateOfWeekly','DuplicateOfLaterRun')).Count)
        $usageFolderRows.Add([pscustomobject]@{Run=$run.Name;Files=$runRows.Count
            Bytes=(Get-TotalBytes -Items $runRows.ToArray() -Property 'Bytes')
            Decision=if ($eligible) { 'ReviewExactDuplicate' } else { 'Keep' }
            Reason=if (-not $plain) { 'Nested or linked content' } elseif ($runDate -ge (Get-Date).AddDays(-7)) { 'Recent run' }
                elseif (-not $eligible) { 'Distinct or unqualified file remains' } else { 'Every file has a retained SHA256-identical copy' }} )
    }
}
Write-Report 'usage-files.csv' $usageRows.ToArray()
Write-Report 'usage-folders.csv' $usageFolderRows.ToArray()

# These exports have no general weekly-history contract. Report their canonical
# DATA-LAST comparison without treating a different historical snapshot as waste.
$scopeRows = [Collections.Generic.List[object]]::new()
foreach ($relativeRoot in @('DATA-ALL\M365\WorkplaceScope','DATA-ALL\Intune\WorkplaceScope')) {
    $scopeRoot = Join-Path $root $relativeRoot
    if (-not (Test-Path -LiteralPath $scopeRoot -PathType Container)) { continue }
    foreach ($file in @(Get-ChildItem -LiteralPath $scopeRoot -File -Recurse -Force | Sort-Object FullName)) {
        $base = [regex]::Replace($file.Name, '_\d{8}[-_]\d{6}(?=\.csv$)', '')
        $latest = Join-Path (Join-Path $root 'DATA-LAST') $base
        $copyStatus = 'NoCanonicalCopy'
        if ((Test-Path -LiteralPath $latest -PathType Leaf) -and (Test-OwnedPlainPath $file.FullName) -and
            (Test-OwnedPlainPath $latest)) {
            $copyStatus = if ((Get-Item -LiteralPath $latest).Length -eq $file.Length -and
                (Get-FileHashValue $latest) -eq (Get-FileHashValue $file.FullName)) { 'ExactCanonicalCopy' }
                else { 'DistinctSnapshot' }
        }
        $scopeRows.Add([pscustomobject]@{Path=(Get-Relative $file.FullName);Bytes=$file.Length
            LastWriteTimeUtc=$file.LastWriteTimeUtc.ToString('o');CanonicalName=$base;Comparison=$copyStatus})
    }
}
Write-Report 'workplace-scope.csv' $scopeRows.ToArray()

$entraRows = [Collections.Generic.List[object]]::new()
$entraRoot = Join-Path $root 'DATA-ALL\M365\Devices\EntraDevices'
if (Test-Path -LiteralPath $entraRoot -PathType Container) {
    foreach ($file in @(Get-ChildItem -LiteralPath $entraRoot -File -Recurse -Force -Filter 'SmartM365-EntraDevices-RemoveCandidates_*.ps1')) {
        $entraRows.Add([pscustomobject]@{Path=(Get-Relative $file.FullName);Bytes=$file.Length
            LastWriteTimeUtc=$file.LastWriteTimeUtc.ToString('o');Decision='ManualOperatorUseReview'})
    }
}
Write-Report 'entra-generated-scripts.csv' $entraRows.ToArray()

$lastRows = [Collections.Generic.List[object]]::new()
$last = Join-Path $root 'DATA-LAST'
if (Test-Path -LiteralPath $last -PathType Container) {
    foreach ($file in @(Get-ChildItem -LiteralPath $last -File -Force | Sort-Object Name)) {
        $old = $file.LastWriteTimeUtc -lt [datetime]'2026-08-01'
        if (-not $old -and $file.Extension -ne '.lock' -and
            $file.Name -notlike '*.migration.log') { continue }
        $decision = if ($file.Name -eq 'Intune_Devices_Win11Readiness.csv') { 'KeepRequested' }
            elseif ($file.Extension -eq '.lock') { 'KeepCoordinationLock' }
            elseif ($file.Name -like '*.migration.log') { 'ReviewOldTransportLog' }
            elseif ($file.Name -match '_MAXITEMS[-_]\d+') { 'ReviewOldTestOutput' }
            else { 'ReviewOldOutput' }
        $lastRows.Add([pscustomobject]@{Name=$file.Name;Bytes=$file.Length
            LastWriteTimeUtc=$file.LastWriteTimeUtc.ToString('o');Decision=$decision})
    }
}
Write-Report 'data-last-old.csv' $lastRows.ToArray()

$repairRows = [Collections.Generic.List[object]]::new()
$repairRoot = Join-Path $root 'DATA-REPAIR-BACKUPS'
if (Test-Path -LiteralPath $repairRoot -PathType Container) {
    foreach ($folder in @(Get-ChildItem -LiteralPath $repairRoot -Directory -Force)) {
        $issues = [Collections.Generic.List[string]]::new()
        $auditPath = @('repair.json.txt','repair.json' | ForEach-Object { Join-Path $folder.FullName $_ } |
            Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1)
        $count = 0; $status = 'MissingAudit'
        if ($auditPath.Count) {
            try {
                $audit = Get-Content -LiteralPath $auditPath[0] -Raw | ConvertFrom-Json
                $status = [string]$audit.Status
                if ($audit.SchemaVersion -ne 1 -or $audit.TenantKey -ne $Tenant -or $status -ne 'Completed') {
                    $issues.Add('Audit is not a completed repair for this tenant.')
                }
                foreach ($entry in @($audit.Files)) {
                    $count++
                    if ($entry.Status -ne 'Repaired' -or $entry.Relative -notmatch '^DATA-ALL[/\\]' -or
                        $entry.OriginalSHA256 -notmatch '^[0-9A-Fa-f]{64}$' -or
                        $entry.RepairedSHA256 -notmatch '^[0-9A-Fa-f]{64}$') { $issues.Add('Invalid file receipt.'); continue }
                    $relative = $entry.Relative.Replace('/', [IO.Path]::DirectorySeparatorChar)
                    foreach ($check in @(
                        @{Path=(Join-Path $root $relative);Hash=$entry.RepairedSHA256},
                        @{Path=(Join-Path $folder.FullName (Join-Path 'originals' $relative));Hash=$entry.OriginalSHA256},
                        @{Path=(Join-Path $folder.FullName (Join-Path 'replaced' $relative));Hash=$entry.OriginalSHA256}
                    )) {
                        if (-not (Test-Path -LiteralPath $check.Path -PathType Leaf) -or
                            -not (Test-OwnedPlainPath $check.Path) -or (Get-FileHashValue $check.Path) -ne $check.Hash) {
                            $issues.Add('Missing or changed repaired source/backup: ' + $entry.Relative)
                        }
                    }
                }
                if ($count -eq 0) { $issues.Add('Audit has no file receipts.') }
            } catch { $issues.Add('Unreadable audit: ' + $_.Exception.Message) }
        } else { $issues.Add('Missing repair audit.') }
        $files = @(Get-ChildItem -LiteralPath $folder.FullName -File -Recurse -Force)
        if ($count -gt 0 -and $files.Count -ne (2 * $count + 1)) {
            $issues.Add('Backup contains unexpected or missing files.')
        }
        $repairRows.Add([pscustomobject]@{Folder=(Get-Relative $folder.FullName);Status=$status
            ReceiptFiles=$count;StoredFiles=$files.Count;Bytes=(Get-TotalBytes -Items $files -Property 'Length')
            Verification=if ($issues.Count) { 'Review' } else { 'CompletedAndHashVerified' }
            Issues=($issues -join ' | ')})
    }
}
Write-Report 'repair-backups.csv' $repairRows.ToArray()

$preparedRows = [Collections.Generic.List[object]]::new()
$prepared = Join-Path $root 'DATA-POWERBI'
foreach ($family in 'failed','retired') {
    $familyRoot = Join-Path $prepared $family
    if (-not (Test-Path -LiteralPath $familyRoot -PathType Container)) { continue }
    foreach ($folder in @(Get-ChildItem -LiteralPath $familyRoot -Directory -Force)) {
        $files = @(Get-ChildItem -LiteralPath $folder.FullName -File -Recurse -Force)
        $preparedRows.Add([pscustomobject]@{Family=$family;Folder=(Get-Relative $folder.FullName)
            Files=$files.Count;Bytes=(Get-TotalBytes -Items $files -Property 'Length')
            NewestUtc=if ($files.Count) { (@($files | Sort-Object LastWriteTimeUtc -Descending)[0].LastWriteTimeUtc.ToString('o')) } else { '' }
            Decision=if ($family -eq 'retired') { 'KeepAuditReceipts' } else { 'ReviewFailureDiagnostics' }})
    }
}
Write-Report 'prepared-diagnostics.csv' $preparedRows.ToArray()

if (-not $EffectiveJobsPath) { $EffectiveJobsPath = Join-Path $PSScriptRoot '..\Orchestrator\Orchestrator-Jobs.json' }
$lockPath = Join-Path $root '.cmdb-preparation.lock'
$jobStatus = 'MissingEffectiveJobs'; $lockStatus = 'Absent'
if (Test-Path -LiteralPath $EffectiveJobsPath -PathType Leaf) {
    try {
        $jobs = Get-Content -LiteralPath $EffectiveJobsPath -Raw | ConvertFrom-Json
        $legacy = @($jobs.Jobs | Where-Object { $_.Name -eq 'CmdbEvidence-Prepare' -or
            $_.ScriptPath -match 'SmartM365-CmdbEvidence-Prepare\.ps1' -or
            'CmdbEvidence-Prepare' -in @($_.DependsOn) })
        $jobStatus = if ($legacy.Count) { 'LegacyJobOrDependencyPresent' } else { 'NoLegacyJobOrDependency' }
    } catch { $jobStatus = 'InvalidEffectiveJobs: ' + $_.Exception.Message }
}
if (Test-Path -LiteralPath $lockPath -PathType Leaf) {
    $lockStatus = 'Unavailable'
    try {
        $handle = [IO.File]::Open($lockPath, 'Open', 'ReadWrite', 'None')
        $handle.Dispose()
        $lockStatus = if (Test-OwnedPlainPath $lockPath) { 'Free' } else { 'LinkedPath' }
    } catch { $lockStatus = 'HeldOrInaccessible' }
}
Write-Report 'cmdb-lock.csv' @([pscustomobject]@{Path=(Get-Relative $lockPath)
    EffectiveJobsPath=$EffectiveJobsPath;JobStatus=$jobStatus;LockStatus=$lockStatus
    Decision=if ($jobStatus -eq 'NoLegacyJobOrDependency' -and $lockStatus -eq 'Free') { 'EligibleForCleanup' } else { 'KeepPendingReview' }})

[pscustomobject]@{
    Mode = 'Preview'
    DataRootPath = $root
    UsageDuplicateFolders = @($usageFolderRows | Where-Object Decision -eq 'ReviewExactDuplicate').Count
    OldDataLastFiles = @($lastRows | Where-Object Decision -NotIn @('KeepRequested','KeepCoordinationLock')).Count
    RepairBackupsVerified = @($repairRows | Where-Object Verification -eq 'CompletedAndHashVerified').Count
    CmdbLock = $lockStatus
    ReportDirectory = $reportRoot
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAclJqPAvwr9ifH
# 0qQP9MNYfApvbxd7dP53WN3p5/5Uf6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEiM5ebojwDtCY3X+G0TGtjMxm4HQJ3+5f9bxrJRWzlOMA0GCSqG
# SIb3DQEBAQUABIIBgEGZRK6X4mfgsT1iOG/qQVUIMf9xggN+9lWQHuoc/ONgtX/V
# xkdPTt4cBbiGW2QE8Qxz/YViLqQk91SklAqY8l9X4OFUcWmWN0TvxgYmvmC0TdWE
# hcMEre35O8QoIeJBeL5RUJoY0jy84iTQs5usPw07R+UzG060wY3JWUOiECKToFLu
# w+59Le+YA8rOqq9wC1lZuFbOjYxV0VkyzkfOd/AGUPNMEJrsXeRl3zBJwAq9ENTp
# a5VTwrDx5sv/8vPqlwz7mgzqwu5euWQIgJkjA/ZDd4To+8CmUm1PcsaI+kuNEaH9
# DrKmRLAgRV12BhUyJtNyjknN9HvxnGSDv0MwW19Yrrj4KpkABc3avXbZ6PD/uYTH
# WgxVmhRgzEiNwJClWYh05U9dGVcK43GyVLEFKZ/x67BmNWmhGlf01m39/fLUyj2z
# 8p/V+AtPEhRzl4vGrkphldy+R4htjDi4/CJapN4Ser92kSBXgWU8G8FTL1dmGJM9
# Rj3PNWGNvIAqcqNqMqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxODI0
# MzJaMC8GCSqGSIb3DQEJBDEiBCBLiRdVBkyt9Se+UK/mg82bm+lTwW8LCLv8Ha1c
# WvAW6zANBgkqhkiG9w0BAQEFAASCAgCr+PUBcIrK5eptRjmpMLsiGBnnafPhMV/u
# zZJjOPhzZqVPjqAzTwHf6xpbP9IqF2tb+jTO9ky5wvMmoojuUYFiazQgf6Ua25F5
# d4zxyR8pWjl+43W2ACEQ7DmEdAYvo7eQg9o8Tu70oXnCvjtDZUjWE8OUorW+Loa7
# 1yOh8m21QocGnSdS46BX0YJdkvYmFUeLbC/mKDgn5SjlgSwhKGRMIsFX//WLLTgd
# E2gaL5xQu8V7wLolZdeMPopsHYlhfc3xzaZPtT0He3Rt8t7bKovElzfGSMIuO4rs
# 4IiI6hdFuEVFidkUVhECdGnZ5y4qPcfk7WI0mm/hD6Hj0fQSsp1tVwIOrFnec7vE
# BFJeqddp0nqJGaCH4/jTKBSXlFjvwIaCmtwHwt9IeG9AG+qIbSUzQjV9jFcQOkDg
# aaznigNvTlfzn6q4p7bgyXgHTjCDx9sWbxKDdqwpBY8OyabIX+pJIlASlG2UYH+n
# ejidh5QnaRL4sGRY0w66uKV/6fXVH8Jo3ZvtZuEd0R32DzCffhGJA80VmR2sgsYk
# vVHWGaj6mmECQJvg1cl9Sjv9UACvbe4a1GoF1jX7opzg1umj5lR4YwXmTXqLjQ+P
# 3hqB7SaSMlshE3CQV+j1hUyeKcdby1pCQa+u+rc50t3T39xFNynUK0l+hhfVhz72
# coXXiS3NhA==
# SIG # End signature block
