<#
.SYNOPSIS
    Verify portfolio summary states without connecting to SharePoint.
.VERSION
    1.0.7
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
    if ($fileBadge -ne "$($row.ComparisonDate) · $($row.ComparisonPercent)") {
        throw 'File comparison badge did not show the date and rate from the displayed report.'
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
    if ($fileBadge -ne "$($row.ComparisonDate) · $($row.ComparisonPercent)") {
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
    if ($permissionBadge -ne "$($row.PermissionComparisonDate) · $($row.PermissionComparisonPercent)") {
        throw 'Permission comparison badge did not show the date and rate from the displayed report.'
    }

    $permissionSummary.MatchedPermissions = 10
    $permissionSummary.MissingInSPO = 0
    $permissionSummary.ValidationStatus = 'NoRelevantDifference'
    $permissionSummary | Export-Csv -LiteralPath $permissionSummaryPath -Delimiter ',' -NoTypeInformation -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Up to date' -or $row.StatusTooltip -notmatch 'Permissions: Up to date') {
        throw 'Both clean comparisons must produce an up-to-date status.'
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
    if ($row.Status -ne 'Compare needed') { throw 'A newer scan must require a new comparison.' }
    Remove-Item -LiteralPath $newerSource

    $newerPermission = Join-Path $sourcePermissionDir ("SP2019-PermissionInventory-Fixture-{0}.csv" -f (Get-Date).AddSeconds(2).ToString('yyyyMMdd-HHmmss'))
    'Principal' | Set-Content -LiteralPath $newerPermission -Encoding utf8
    $row = Get-SmartM365PortfolioRow @params
    if ($row.Status -ne 'Compare needed' -or $row.StatusTooltip -notmatch 'Permissions: Compare needed') {
        throw 'A newer permission scan must require a new comparison.'
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAIGkwtiZhjnpX7
# C/t57Eit62nDy1RFWCNw5qEqycmTSKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIonaxJtyQChD9H3n87peOYS9TrHnCAbO+j7uRcyvK2XMA0GCSqG
# SIb3DQEBAQUABIIBgKfPoomTukavfowhbXRCyr4/oul/Qid5EfoeKleQ9tRNRPyJ
# HsrraQva0J0bkOLGZcCtIYXmp9pP55+PobUw9i1btgyDX+B4J8NMPpsvbfrn/dfj
# BclJk/KZYNarHrboGqTcF/rZtD0ifWqUzA1tGoUcOvWKAjZnJIks2xSScDIGWiT4
# MjZejG6dhBgMQvGswMwlScSX9y1rHVEolQQWA5hDxJknxi9eRy+UF5OFVle+Nbt+
# T1rLqA2xnqZrIUljWbUpFZM6OmAk7PnX1AqPsLe/0I1k39kAgmX2BfTYKAgL02aS
# 7ap0T4sxNVvz0RqOvk25ZLcdVGtD5TgdI475gzdWBFkoEplxPxhaFophAIWActdK
# PzG0SoIaimM21WnyvHfkv7QnFFgi/ZkupxagcuhNc8Ayc5RiL3EuNjAw3ZtcwQ1r
# JwHF/uJyAzXil4DWd15XNwH5z6GMcqVmrzdPg895juL2l6B71cmB3RpdkFu/taR4
# kTv1R6t+08UWluetsqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQxOTAz
# NTVaMC8GCSqGSIb3DQEJBDEiBCCXcLp9VjaSSIYnse/WCAkKNoHLntRsxrM6Stxc
# imxP1jANBgkqhkiG9w0BAQEFAASCAgBkpJxkocSr9nfd/COgPZYtEv4TPhKyln49
# nCk5FdexfccdTYn4a7+vhpbPZXVMtaLnubVLFsTiIG5aF0a8XFLRNQQIotohWRbQ
# Z3b0VtHeDpwZexpZZn40sT/5L5zeTrJSfdrUzIz+JzTXGKBJpRlqGdQ425dcIkVP
# zbAqcgqYM78oJsy6zY7tatCcjRKiDr5DsYlUOv97nWxlim70O6wpX0xHlLPqSVR1
# hQ+C808qDqrdfqliwvByksaxIRLFci+RpWxGZNX+5pl7cW0WqBwCFMnFao7RQxoY
# DImKplr9AxbHnmuJRmdOUTjTo8IBqxfCpfqtMeErWBU+ksHJYduV8x/LRHkxroyD
# MX4DyZZskZO5cOA7kz1FsfLT5jLnf402ROHKaNXyYC6fA//xbNIzxQNosu8rbdDv
# /HgmajdukAaFdfhsliUXMn5s8c0+U+VJ0RFm5KE34kb75rVxnpXbFeVJvI/vEuq/
# n5mAtXR3NeJAJgQ9+AqDoVFcAU1pX6fLMCvqBrq0swtP80KAnnVTHYYp0d9Vsq6l
# Zfbd4Z7L1hBRCTeu0EgLwFmg2DaBkzMZ8t9lLu8T0QExEh+Xqh1t+FLgEP588FC3
# Ufd+ZPyVZkphVLTapVZvftCt30hKLHcYxTKWwS7yfVbBd4Jrm8FT3ToYK+pydU8M
# fQ+7ADqgzg==
# SIG # End signature block
