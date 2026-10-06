<#
.SYNOPSIS
    Verify portfolio summary states without connecting to SharePoint.
.VERSION
    1.0.9
#>

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
. (Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-Summary.ps1')
$guiPath = Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-GUI.ps1'
$tokens = $null; $parseErrors = $null
$guiAst = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | ForEach-Object Message) }
$badgeFunction = $guiAst.Find({
    param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'Get-ComparisonBadgeText'
}, $true)
if (-not $badgeFunction) { throw 'Comparison badge formatter is missing from the GUI.' }
. ([scriptblock]::Create($badgeFunction.Extent.Text))
foreach ($name in @('Get-LatestCsvFile', 'Get-CsvFileItems', 'Format-RunAgeText')) {
    $node = $guiAst.Find({
        param($item) $item -is [Management.Automation.Language.FunctionDefinitionAst] -and $item.Name -eq $name
    }, $true)
    if (-not $node) { throw "GUI scan selector is missing: $name" }
    . ([scriptblock]::Create($node.Extent.Text))
}

$safeRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$testRoot = [IO.Path]::GetFullPath((Join-Path $safeRoot ('SharePointMigration-summary-tests-' + [guid]::NewGuid().ToString('N'))))
if (-not $testRoot.StartsWith($safeRoot, [StringComparison]::OrdinalIgnoreCase) -or
    (Split-Path $testRoot -Leaf) -notlike 'SharePointMigration-summary-tests-*') {
    throw 'Test directory is outside the temporary folder.'
}

try {
    $sourceDir = Join-Path $testRoot 'scans\source\files'
    $targetDir = Join-Path $testRoot 'scans\target\files'
    $comparisonDir = Join-Path $testRoot 'comparisons\files'
    $sourcePermissionDir = Join-Path $testRoot 'scans\source\permissions'
    $targetPermissionDir = Join-Path $testRoot 'scans\target\permissions'
    $permissionComparisonDir = Join-Path $testRoot 'comparisons\permissions'
    foreach ($dir in @($sourceDir,$targetDir,$comparisonDir,
            $sourcePermissionDir,$targetPermissionDir,$permissionComparisonDir)) {
        [void](New-Item -ItemType Directory -Path $dir -Force)
    }
    $migration = [pscustomobject]@{
        Name = 'Fixture'; Root = $testRoot
        Config = @{
            Name = 'Fixture'
            Output = @{
                SourceFileScans = 'scans\source\files'
                TargetFileScans = 'scans\target\files'
                FileComparisons = 'comparisons\files'
                SourcePermissionScans = 'scans\source\permissions'
                TargetPermissionScans = 'scans\target\permissions'
                PermissionComparisons = 'comparisons\permissions'
            }
            Comparison = @{ MaxScanAgeHours = 24; MaxScanAgeDifferenceHours = 12 }
        }
    }
    $params = @{
        Migration = $migration; SourceScope = 'source'; SourceTooltip = 'source'
        TargetScope = 'target'; TargetTooltip = 'target'
        SourceType = 'SP2019'; TargetType = 'SPO'
    }
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Scan needed' -or $row.ScanGapText -ne '—' -or
        $null -ne $row.ScanGapDays -or $row.ComparisonPercent -ne '—' -or
        $row.PermissionComparisonPercent -ne '—' -or
        $row.ComparisonDisplay -ne '—' -or $row.PermissionComparisonDisplay -ne '—' -or
        $null -ne $row.GlobalComparisonRate -or $row.GlobalComparisonVisual.Percent -ne '—' -or
        $row.SourceScansDisplay -ne "Files —`nPerms —" -or
        $row.TargetScansDisplay -ne "Files —`nPerms —" -or
        $row.StatusTooltip -notmatch 'Files: Scan needed' -or
        $row.StatusTooltip -notmatch 'Permissions: Scan needed') { throw 'Missing scans were not identified.' }

    $sourceStamp = (Get-Date).AddMinutes(-2).ToString('yyyyMMdd-HHmmss')
    $targetStamp = (Get-Date).AddMinutes(-1).ToString('yyyyMMdd-HHmmss')
    $comparisonStamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $sourceCsv = Join-Path $sourceDir "SP2019-FileInventory-Fixture-$sourceStamp.csv"
    $targetCsv = Join-Path $targetDir "SPO-FileInventory-Fixture-$targetStamp.csv"
    'File' | Set-Content -LiteralPath $sourceCsv -Encoding utf8
    'File' | Set-Content -LiteralPath $targetCsv -Encoding utf8
    $targetErrorPath = Join-Path $targetDir ("{0}-Errors.csv" -f [IO.Path]::GetFileNameWithoutExtension($targetCsv))
    'Scope;Message' | Set-Content -LiteralPath $targetErrorPath -Encoding utf8
    if ($null -ne (Get-LatestCsvFile -Directory $targetDir -Filter 'SPO-FileInventory-Fixture-*.csv') -or
        @(Get-CsvFileItems -Directory $targetDir -Filter 'SPO-FileInventory-Fixture-*.csv').Count -ne 0 -or
        $null -ne (Get-SmartM365LatestPortfolioScan -Directory $targetDir -Filter 'SPO-FileInventory-Fixture-*.csv')) {
        throw 'A scan with an error sidecar must be hidden from the GUI and overview.'
    }
    Remove-Item -LiteralPath $targetErrorPath
    $row = Get-SmartM365PortfolioRow @params
    if ($row.SourceFileScanFile.FullName -ne $sourceCsv -or
        $row.TargetFileScanFile.FullName -ne $targetCsv -or
        $row.SourceFileScanDate -isnot [datetime] -or
        $row.TargetFileScanDate -isnot [datetime]) {
        throw 'The overview did not expose the latest file inventories for metric calculation.'
    }
    if ($row.Status -ne 'Scan needed' -or $row.StatusTooltip -notmatch 'Files: Compare needed') {
        throw 'The combined status must include missing permission scans.'
    }
    if ($row.SourceScansDisplay -ne "Files $($row.SourceScan)`nPerms —" -or
        $row.TargetScansDisplay -ne "Files $($row.TargetScan)`nPerms —") {
        throw 'The overview must distinguish available file scans from missing permission scans.'
    }
    if ($row.ScanGapDays -le 0 -or $row.ScanGapDays -ge 0.01 -or
        $row.ScanGapTooltip -notmatch 'Target is newer') {
        throw 'Scan day difference or direction is incorrect.'
    }
    $targetReceipt = "$targetCsv.manifest.json.txt"
    @{ SchemaVersion = 1; InventoryFile = [IO.Path]::GetFileName($targetCsv)
        CompletedAtUtc = [datetime]::UtcNow.ToString('o'); Sha256 = 'synthetic'; Rows = 0 } |
        ConvertTo-Json | Set-Content -LiteralPath $targetReceipt -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Scan needed' -or $row.StatusTooltip -notmatch 'Target file inventory is empty; compare') {
        throw 'A pre-migration empty target must invite file comparison.'
    }

    $folder = Join-Path $comparisonDir "Fixture-files-$comparisonStamp-test"
    [void](New-Item -ItemType Directory -Path $folder)
    $summaryPath = Join-Path $folder 'Summary.csv'
    $summary = [pscustomobject]@{
        SourceCsv = $sourceCsv; TargetCsv = $targetCsv
        MatchedKeys = 8; SourceUniqueKeys = 10; TargetUniqueKeys = 10
        ValidationStatus = 'ReviewNeeded'; MissingInTarget = 2; ExtraInTarget = 0
        SourceFilteredRows = 0; TargetFilteredRows = 0
    }
    $summary | Export-Csv -LiteralPath $summaryPath -Delimiter ';' -NoTypeInformation -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Review needed' -or $row.ComparisonRate -ne 80 -or
        $row.ComparisonDisplay -ne "$($row.ComparisonDate) · $($row.ComparisonPercent)" -or
        $row.PermissionComparisonDisplay -ne '—') { throw 'Comparison findings or overview display were not identified.' }
    $fileBadge = Get-ComparisonBadgeText -Folder (Get-Item -LiteralPath $folder) -MigrationName 'Fixture' -Kind Files
    if ($fileBadge -notlike "* today* · $($row.ComparisonPercent)") {
        throw 'File comparison badge did not show the date and rate from the displayed report.'
    }
    if ($null -ne $row.GlobalComparisonRate -or $row.GlobalComparisonVisual.Percent -ne '—' -or
        $row.ComparisonVisual.Caption -notlike '* today' -or $row.ComparisonVisual.Color -ne '#9A6700') {
        throw 'A missing permission rate must leave the global average unavailable; file date and partial-rate styling must be visible.'
    }
    $summary.MatchedKeys = 0
    $summary.TargetUniqueKeys = 0
    $summary.MissingInTarget = 10
    $summary.ValidationStatus = 'ReviewNeeded'
    $summary | Add-Member -NotePropertyName TargetEmptyVerified -NotePropertyValue 'True'
    $summary | Export-Csv -LiteralPath $summaryPath -Delimiter ';' -NoTypeInformation -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.ComparisonRate -ne 0 -or $row.ComparisonPercent -notmatch '^0[,.]00 %$' -or
        $row.StatusTooltip -notmatch 'Target scan verified empty') {
        throw 'Verified empty target comparison must display a 0 percent file rate.'
    }
    $fileBadge = Get-ComparisonBadgeText -Folder (Get-Item -LiteralPath $folder) -MigrationName 'Fixture' -Kind Files
    if ($fileBadge -notlike "* today* · $($row.ComparisonPercent)") {
        throw 'Verified empty target comparison badge must show 0 percent.'
    }
    $summary.PSObject.Properties.Remove('TargetEmptyVerified')
    $summary | Export-Csv -LiteralPath $summaryPath -Delimiter ';' -NoTypeInformation -Encoding utf8
    $unverifiedBadge = Get-ComparisonBadgeText -Folder (Get-Item -LiteralPath $folder) -MigrationName 'Fixture' -Kind Files
    if ($unverifiedBadge -notmatch 'Rate unavailable$') {
        throw 'Unverified empty target comparison badge must not show a rate.'
    }
    Remove-Item -LiteralPath $targetReceipt -Force
    $summary.MatchedKeys = 8
    $summary.TargetUniqueKeys = 10
    $summary.MissingInTarget = 2
    $summary | Export-Csv -LiteralPath $summaryPath -Delimiter ';' -NoTypeInformation -Encoding utf8
    [void]$migration.Config.Comparison.Remove('MaxScanAgeHours')
    [void]$migration.Config.Comparison.Remove('MaxScanAgeDifferenceHours')
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -eq 'Refresh error' -or $row.ComparisonRate -ne 80) {
        throw 'Optional scan age settings must use defaults when absent.'
    }
    $migration.Config.Comparison.MaxScanAgeHours = 24
    $migration.Config.Comparison.MaxScanAgeDifferenceHours = 12

    $summary.MatchedKeys = 10
    $summary.MissingInTarget = 0
    $summary.ValidationStatus = 'NoRelevantDifference'
    $summary | Export-Csv -LiteralPath $summaryPath -Delimiter ';' -NoTypeInformation -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Scan needed' -or $row.ComparisonRate -ne 100 -or
        $row.StatusTooltip -notmatch 'Files: Up to date') {
        throw 'A clean file comparison must not hide missing permission scans.'
    }

    $sourcePermissionStamp = (Get-Date).AddMinutes(-4).ToString('yyyyMMdd-HHmmss')
    $targetPermissionStamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $sourcePermissionCsv = Join-Path $sourcePermissionDir "SP2019-PermissionInventory-Fixture-$sourcePermissionStamp.csv"
    $targetPermissionCsv = Join-Path $targetPermissionDir "SPO-PermissionInventory-Fixture-$targetPermissionStamp.csv"
    'Principal' | Set-Content -LiteralPath $sourcePermissionCsv -Encoding utf8
    'Principal' | Set-Content -LiteralPath $targetPermissionCsv -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Compare needed' -or $row.StatusTooltip -notmatch 'Permissions: Compare needed') {
        throw 'A missing permission comparison was not identified.'
    }
    if ($row.SourcePermissionScan -eq $row.SourceScan -or
        $row.TargetPermissionScan -eq $row.TargetScan -or
        $row.SourceScansDisplay -ne "Files $($row.SourceScan)`nPerms $($row.SourcePermissionScan)" -or
        $row.TargetScansDisplay -ne "Files $($row.TargetScan)`nPerms $($row.TargetPermissionScan)" -or
        $row.TargetScansSortDate.ToString('yyyy-MM-dd HH:mm') -ne $row.TargetPermissionScan -or
        $row.TargetScansTooltip -notmatch 'PermissionInventory') {
        throw 'The overview must show distinct file and permission scan dates on both sides.'
    }
    $permissionFolder = Join-Path $permissionComparisonDir "Fixture-permissions-$comparisonStamp-test"
    [void](New-Item -ItemType Directory -Path $permissionFolder)
    $permissionSummary = [pscustomobject]@{
        SourceCsv = $sourcePermissionCsv; TargetCsv = $targetPermissionCsv
        MatchedPermissions = 7; SourceUniqueKeys = 10; TargetUniqueKeys = 10
        ValidationStatus = 'ReviewNeeded'; MissingInSPO = 3
    }
    $permissionSummaryPath = Join-Path $permissionFolder 'Summary.csv'
    $permissionSummary | Export-Csv -LiteralPath $permissionSummaryPath -Delimiter ',' -NoTypeInformation -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.PermissionComparisonRate -ne 70 -or $row.PermissionComparisonDate -eq '—' -or
        $row.PermissionComparisonDisplay -ne "$($row.PermissionComparisonDate) · $($row.PermissionComparisonPercent)" -or
        $row.ComparisonRate -ne 100 -or $row.Status -ne 'Review needed' -or
        $row.StatusTooltip -notmatch 'Permissions: Review needed') {
        throw 'Permission findings were not included in the combined status.'
    }
    $permissionBadge = Get-ComparisonBadgeText -Folder (Get-Item -LiteralPath $permissionFolder) -MigrationName 'Fixture' -Kind Permissions
    if ($permissionBadge -notlike "* today* · $($row.PermissionComparisonPercent)") {
        throw 'Permission comparison badge did not show the date and rate from the displayed report.'
    }
    # Different source counts must not change the user-approved 50/50 weighting.
    $permissionSummary.MatchedPermissions = 700
    $permissionSummary.SourceUniqueKeys = 1000
    $permissionSummary.TargetUniqueKeys = 1000
    $permissionSummary | Export-Csv -LiteralPath $permissionSummaryPath -Delimiter ',' -NoTypeInformation -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.GlobalComparisonRate -ne 85 -or $row.GlobalComparisonVisual.Percent -ne ('{0:N2} %' -f 85) -or
        $row.GlobalComparisonVisual.Progress -ne 85 -or $row.GlobalComparisonTooltip -notmatch 'Equal-weight' -or
        $row.PermissionComparisonVisual.Caption -notlike '* today') {
        throw 'Global comparison must average the two rates equally rather than pool their denominators.'
    }
    $permissionSummary.SourceUniqueKeys = 10
    $permissionSummary.TargetUniqueKeys = 10

    $permissionSummary.MatchedPermissions = 10
    $permissionSummary.MissingInSPO = 0
    $permissionSummary.ValidationStatus = 'NoRelevantDifference'
    $permissionSummary | Export-Csv -LiteralPath $permissionSummaryPath -Delimiter ',' -NoTypeInformation -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Up to date' -or $row.StatusTooltip -notmatch 'Permissions: Up to date') {
        throw 'Both clean comparisons must produce an up-to-date status.'
    }
    if ($row.GlobalComparisonRate -ne 100 -or $row.GlobalComparisonVisual.Color -ne '#087F5B' -or
        $row.ComparisonVisual.Color -ne '#087F5B' -or $row.PermissionComparisonVisual.Color -ne '#087F5B') {
        throw 'Three complete, current rates must appear green.'
    }

    $permissionSummary.PSObject.Properties.Remove('ValidationStatus')
    $permissionSummary | Export-Csv -LiteralPath $permissionSummaryPath -Delimiter ',' -NoTypeInformation -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Review needed') {
        throw 'A legacy permission report without ValidationStatus must remain readable and require review.'
    }
    $permissionSummary | Add-Member -NotePropertyName ValidationStatus -NotePropertyValue 'NoRelevantDifference'
    $permissionSummary | Export-Csv -LiteralPath $permissionSummaryPath -Delimiter ',' -NoTypeInformation -Encoding utf8

    $newerSource = Join-Path $sourceDir ("SP2019-FileInventory-Fixture-{0}.csv" -f (Get-Date).AddSeconds(1).ToString('yyyyMMdd-HHmmss'))
    'File' | Set-Content -LiteralPath $newerSource -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Compare needed' -or $row.ComparisonVisual.Notice -ne 'Recalculate' -or
        $row.GlobalComparisonVisual.Notice -ne 'Recalculate' -or $row.GlobalComparisonVisual.Color -ne '#9A6700') {
        throw 'A newer file scan must mark the file rate and global average for recalculation.'
    }
    Remove-Item -LiteralPath $newerSource

    $newerPermission = Join-Path $sourcePermissionDir ("SP2019-PermissionInventory-Fixture-{0}.csv" -f (Get-Date).AddSeconds(2).ToString('yyyyMMdd-HHmmss'))
    'Principal' | Set-Content -LiteralPath $newerPermission -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Compare needed' -or $row.StatusTooltip -notmatch 'Permissions: Compare needed' -or
        $row.PermissionComparisonVisual.Notice -ne 'Recalculate' -or $row.GlobalComparisonVisual.Notice -ne 'Recalculate') {
        throw 'A newer permission scan must require a new comparison.'
    }

    $visual = Get-SmartM365PortfolioRateVisual -Rate 0 -Date (Get-Date).Date.AddDays(-1).AddHours(11)
    if ($visual.Percent -ne ('{0:N2} %' -f 0) -or $visual.ProgressVisibility -ne 'Visible' -or
        $visual.Caption -ne '11:00 yesterday' -or $visual.Color -ne '#9A6700') {
        throw 'A zero rate is available and must not be confused with missing data.'
    }
    $visual = Get-SmartM365PortfolioRateVisual -Rate $null -Date $null
    if ($visual.ProgressVisibility -ne 'Hidden' -or $visual.Color -ne '#64748B' -or $visual.Caption -ne 'No comparison') {
        throw 'Missing rates must remain visibly unavailable.'
    }

    'SharePointMigration summary offline tests passed.'
}
finally {
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    if ($resolvedRoot.StartsWith($safeRoot, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path $resolvedRoot -Leaf) -like 'SharePointMigration-summary-tests-*' -and
        (Test-Path -LiteralPath $resolvedRoot -PathType Container)) {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}


# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBchCcC+rRnYS7d
# k+JsgUjMkhYNS5Ar5VpQz1/lbF0OwaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC+EHoNd1g2nI7cXzoxCHfC
# p1Vew/7w6cox8eX+3DeL7TANBgkqhkiG9w0BAQEFAASCAYBwrUbZffaEyX+D9p+d
# kO6fDi9y4XVS1BRmd6+XylhfjasxunYV44J9+2KAYKq6gBfXYVMWjkO4HXQXc41C
# FNK5MVaOTRvvHHYUm2EYrAVXaXDn14XX5EY3ZgJzq+u+03WkWfIzSH0+ykoMQ3gy
# OhxjLUeMlv03j8sk/nFJGNTMxuynl3b9fMFOlzc/c//SMkydyUhe4oWMRJVd8bXB
# OwwrOdKV6VnBzPTR7Z7baWkLAynH8GhfLpGYHZ/YPExqoPlLgMKVObrKocqKlG7U
# DjhBI04J12RhmA6gS/JmTV24jN21ks2sqVoAZqpoLjtWPs7TBUbEu5sDTb3AXGAo
# jeJJutUfmoAfuwIM/zhQCL/6nUsIMnCdaSrlxWIgzxRjFqQH5IcjNfmfHY9QeH+P
# UHaZrhhclx5l+M2U4iaCEZ9/hz0GPvDZK+CXJ4JVhSiLq7lECpmCf5BLgoXxbkGh
# LyTV6gCuv2vslTgNosFmti3KNr3RiKwFx4XrtzLnpfz7jD0=
# SIG # End signature block
