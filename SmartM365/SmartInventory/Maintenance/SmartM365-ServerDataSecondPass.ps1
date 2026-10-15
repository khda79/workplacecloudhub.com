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
            Bytes=[int64](($runRows | Measure-Object Bytes -Sum).Sum)
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
            ReceiptFiles=$count;StoredFiles=$files.Count;Bytes=[int64](($files | Measure-Object Length -Sum).Sum)
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
            Files=$files.Count;Bytes=[int64](($files | Measure-Object Length -Sum).Sum)
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKHaXPxuefjhNL
# fX/DlyO5kuNhTzq/nTT5T0YThLm1F6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC/ZCOJD/uj8jJZR5Nf3Pbe
# ZSOtcNUUBVLD+h20zwZgvTANBgkqhkiG9w0BAQEFAASCAYCFR9ZtlY0bhGn4ustg
# tBdnkvhSkqIaxELwfoocnMNmyc7yLqjXtf2pm0fcjRT5001e2IgRfjMZwuRTCCqh
# 3WKnCIML+8VwEP2yqLzq9tutU7BprLWQRK93cOSKFQsObTt805mmSDmY2K/BCxvl
# K35NhmZTcKSfbwqt5vGqG4gFkCriTvcNsQuevy7lZUX1771Gckpg6nx3O/H3HEG4
# I3cY9YjO+oTGuFKLzlRzryWMOKshibcYZJwpv6gj+vjgRXugYbR5Enwmj7exhIr9
# 8kooTXAVO1EqU8gud20GnlcOpUb25rNOVb7tpAe9ZZRc1PAHG+4EMS/o0Z7wkWRk
# yrKBflPMWsC4GsgAc0sV4cL6HkcztkkiTP8+QnbPCvvWGlP7Sp4FwRv3YKijWBNX
# pIz5EE5GzglAqiE39g/mYmrsi0TDj78e7aV22gHQkfTFwdYj5JlWanjM9gih4FmF
# oS+tPRWmKMGrluw+AkTF4s7llfZOSLR2jDP4akd6Mgk/jrE=
# SIG # End signature block
