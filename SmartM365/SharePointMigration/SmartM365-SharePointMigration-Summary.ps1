<#
.SYNOPSIS
    Read-only portfolio summary for the SharePoint migration GUI.
.VERSION
    1.0.9
#>

function Get-SmartM365PortfolioTimestamp {
    param([string]$Name)
    if ($Name -match '-(?<date>\d{8})-(?<time>\d{6})(?:-[^.]+)?(?:\.csv)?$') {
        try {
            return [datetime]::ParseExact(
                ($Matches.date + $Matches.time), 'yyyyMMddHHmmss',
                [Globalization.CultureInfo]::InvariantCulture)
        }
        catch { return $null }
    }
    return $null
}

function Get-SmartM365LatestPortfolioScan {
    param([string]$Directory, [string]$Filter)
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return $null }
    $files = @(Get-ChildItem -LiteralPath $Directory -Filter $Filter -File -Recurse -ErrorAction Stop |
        Where-Object { $_.Name -notlike '*-Errors.csv' } |
        ForEach-Object {
            $stamp = Get-SmartM365PortfolioTimestamp $_.Name
            $date = if ($null -ne $stamp) { $stamp } else { $_.LastWriteTime }
            $provenance = if ($null -ne $stamp) { 'CSV filename (legacy scan)' } else { 'file timestamp (unverified)' }
            $rows = $null
            $receiptPath = "$($_.FullName).manifest.json.txt"
            if (Test-Path -LiteralPath $receiptPath -PathType Leaf) {
                try {
                    $receipt = Get-Content -LiteralPath $receiptPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    if ($receipt.SchemaVersion -ne 1 -or $receipt.InventoryFile -ne $_.Name -or
                        -not $receipt.CompletedAtUtc -or -not $receipt.Sha256) {
                        throw 'Incomplete scan receipt.'
                    }
                    $date = [datetimeoffset]::Parse([string]$receipt.CompletedAtUtc,
                        [Globalization.CultureInfo]::InvariantCulture).LocalDateTime
                    $rows = [int]$receipt.Rows
                    $provenance = 'scan receipt (hash not recalculated)'
                }
                catch { $provenance = 'invalid receipt; ' + $provenance }
            }
            [pscustomobject]@{
                File = $_
                Date = $date
                Provenance = $provenance
                Rows = $rows
            }
        } | Sort-Object Date -Descending)
    if ($files.Count -eq 0) { return $null }
    return $files[0]
}

function Get-SmartM365LatestPortfolioComparison {
    param(
        [string]$Directory,
        [string]$MigrationName,
        [ValidateSet('Files','Permissions')][string]$Kind = 'Files',
        [System.IO.DirectoryInfo]$SelectedFolder
    )
    if ($SelectedFolder) {
        $stamp = Get-SmartM365PortfolioTimestamp $SelectedFolder.Name
        $folders = @([pscustomobject]@{
            Folder = $SelectedFolder
            Date = if ($null -ne $stamp) { $stamp } else { $SelectedFolder.LastWriteTime }
        })
    }
    else {
        if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return $null }
        $folders = @(Get-ChildItem -LiteralPath $Directory -Directory -Filter "$MigrationName-*" -ErrorAction Stop |
            ForEach-Object {
                $stamp = Get-SmartM365PortfolioTimestamp $_.Name
                [pscustomobject]@{
                    Folder = $_
                    Date = if ($null -ne $stamp) { $stamp } else { $_.LastWriteTime }
                }
            } | Sort-Object Date -Descending)
    }
    foreach ($folder in $folders) {
        $path = Join-Path $folder.Folder.FullName 'Summary.csv'
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        try {
            $header = Get-Content -LiteralPath $path -TotalCount 1 -ErrorAction Stop
            $delimiter = if ($header.Contains(';')) { ';' } elseif ($header.Contains("`t")) { "`t" } else { ',' }
            $summary = Import-Csv -LiteralPath $path -Delimiter $delimiter -ErrorAction Stop | Select-Object -First 1
            $matchField = if ($Kind -eq 'Files') { 'MatchedKeys' } else { 'MatchedPermissions' }
            if ($null -eq $summary -or -not $summary.PSObject.Properties[$matchField] -or
                -not $summary.PSObject.Properties['SourceUniqueKeys'] -or
                -not $summary.PSObject.Properties['TargetUniqueKeys']) { continue }
            $matched = [long]0; $source = [long]0; $target = [long]0
            if (-not [long]::TryParse([string]$summary.PSObject.Properties[$matchField].Value, [ref]$matched) -or
                -not [long]::TryParse([string]$summary.SourceUniqueKeys, [ref]$source) -or
                -not [long]::TryParse([string]$summary.TargetUniqueKeys, [ref]$target) -or
                $matched -lt 0 -or $source -lt 0 -or $target -lt 0 -or $matched -gt $source) { continue }
            return [pscustomobject]@{
                Path = $path
                Date = $folder.Date
                Summary = $summary
                Matched = $matched
                Source = $source
                Target = $target
            }
        }
        catch { continue }
    }
    return $null
}

function Get-SmartM365PortfolioRow {
    param(
        [Parameter(Mandatory)]$Migration,
        [Parameter(Mandatory)][AllowEmptyString()][string]$SourceScope,
        [Parameter(Mandatory)][AllowEmptyString()][string]$SourceTooltip,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TargetScope,
        [Parameter(Mandatory)][AllowEmptyString()][string]$TargetTooltip,
        [Parameter(Mandatory)][string]$SourceType,
        [Parameter(Mandatory)][string]$TargetType
    )
    $cfg = $Migration.Config
    $root = $Migration.Root
    $name = [string]$cfg.Name
    $sourceScan = Get-SmartM365LatestPortfolioScan `
        (Join-Path $root $cfg.Output.SourceFileScans) "$SourceType-FileInventory-$name-*.csv"
    $targetScan = Get-SmartM365LatestPortfolioScan `
        (Join-Path $root $cfg.Output.TargetFileScans) "$TargetType-FileInventory-$name-*.csv"
    $comparison = Get-SmartM365LatestPortfolioComparison `
        (Join-Path $root $cfg.Output.FileComparisons) $name
    $permissionComparison = Get-SmartM365LatestPortfolioComparison `
        (Join-Path $root $cfg.Output.PermissionComparisons) $name -Kind Permissions
    $sourcePermissions = Get-SmartM365LatestPortfolioScan `
        (Join-Path $root $cfg.Output.SourcePermissionScans) "$SourceType-PermissionInventory-$name-*.csv"
    $targetPermissions = Get-SmartM365LatestPortfolioScan `
        (Join-Path $root $cfg.Output.TargetPermissionScans) "$TargetType-PermissionInventory-$name-*.csv"

    $rate = $null
    $rateText = '—'
    $status = 'Scan needed'
    $detail = 'Source or target inventory is missing.'
    if ($sourceScan -and $targetScan) {
        $status = 'Compare needed'
        $detail = 'No valid file comparison is available.'
        if ($null -ne $sourceScan.Rows -and $sourceScan.Rows -eq 0) {
            $status = 'Review needed'
            $detail = 'The source file inventory contains zero rows.'
        }
        elseif ($null -ne $targetScan.Rows -and $targetScan.Rows -eq 0) {
            $detail = 'Target file inventory is empty; compare to record the source files missing before a copy.'
        }
        if ($comparison) {
            $summary = $comparison.Summary
            $sourceName = [IO.Path]::GetFileName(([string]$summary.SourceCsv).Replace('/', '\'))
            $targetName = [IO.Path]::GetFileName(([string]$summary.TargetCsv).Replace('/', '\'))
            $usesLatest = [string]::Equals($sourceName, $sourceScan.File.Name, [StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals($targetName, $targetScan.File.Name, [StringComparison]::OrdinalIgnoreCase)
            $targetEmptyVerified = $comparison.Target -eq 0 -and $null -ne $targetScan.Rows -and
                $targetScan.Rows -eq 0 -and $summary.PSObject.Properties['TargetEmptyVerified'] -and
                [string]$summary.TargetEmptyVerified -eq 'True'
            if ($comparison.Source -gt 0 -and ($comparison.Target -gt 0 -or $targetEmptyVerified)) {
                $rate = [double]$comparison.Matched / [double]$comparison.Source * 100
                $rateText = ('{0:N2} %' -f $rate)
            }
            if ($status -eq 'Review needed' -and $detail -like '*source file inventory contains zero rows*') {
                $detail += ' A comparison does not replace a complete scan.'
            }
            elseif (-not $usesLatest) {
                $detail = 'Newer scans are available than those used by the latest comparison.'
            }
            elseif ($null -eq $rate) {
                $status = 'Review needed'
                $detail = 'Empty inventory: comparison percentage is inconclusive.'
            }
            else {
                $validation = [string]$summary.ValidationStatus
                if (-not $validation) {
                    $differences = @('MissingInTarget','ExtraInTarget','ExtraFoldersInTarget',
                        'DifferentSize','ChangedModifiedDate','TargetOlderThanSource','ChangedVersion') |
                        Where-Object {
                            $property = $summary.PSObject.Properties[$_]
                            $property -and [string]$property.Value -match '^[1-9]\d*$'
                        }
                    $filtered = @('SourceFilteredRows','TargetFilteredRows','SourceExcludedRows','TargetExcludedRows') |
                        Where-Object {
                            $property = $summary.PSObject.Properties[$_]
                            $property -and [string]$property.Value -match '^[1-9]\d*$'
                        }
                    $validation = if ($differences.Count -gt 0 -or $rate -lt 100) { 'ReviewNeeded' }
                        elseif ($filtered.Count -gt 0) { 'ReviewScopeFilter' }
                        else { 'NoRelevantDifference' }
                }
                $status = if ($validation -eq 'NoRelevantDifference') { 'Up to date' } else { 'Review needed' }
                $detail = "Validation: $validation. Matches: $($comparison.Matched) / $($comparison.Source)."
                if ($targetEmptyVerified) {
                    $detail += ' Target scan verified empty; source files are listed as missing.'
                }
                $maxAge = 24.0; $maxGap = 12.0
                $configured = 0.0
                if ([double]::TryParse([string]$cfg.Comparison['MaxScanAgeHours'], [ref]$configured) -and $configured -gt 0) {
                    $maxAge = $configured
                }
                if ([double]::TryParse([string]$cfg.Comparison['MaxScanAgeDifferenceHours'], [ref]$configured) -and $configured -gt 0) {
                    $maxGap = $configured
                }
                $sourceAge = ((Get-Date) - $sourceScan.Date).TotalHours
                $targetAge = ((Get-Date) - $targetScan.Date).TotalHours
                $scanGap = [math]::Abs(($sourceScan.Date - $targetScan.Date).TotalHours)
                if ($sourceAge -gt $maxAge -or $targetAge -gt $maxAge -or $scanGap -gt $maxGap -or
                    $sourceAge -lt -0.25 -or $targetAge -lt -0.25 -or
                    $sourceScan.Provenance -like '*unverified*' -or $targetScan.Provenance -like '*unverified*' -or
                    $sourceScan.Provenance -like '*invalid receipt*' -or $targetScan.Provenance -like '*invalid receipt*') {
                    $status = 'Review needed'
                    $detail += ' Check scan freshness or provenance.'
                }
            }
        }
    }
    $scanSourceText = if ($sourceScan) { $sourceScan.Date.ToString('yyyy-MM-dd HH:mm') } else { '—' }
    $scanTargetText = if ($targetScan) { $targetScan.Date.ToString('yyyy-MM-dd HH:mm') } else { '—' }
    $sourcePermissionScanText = if ($sourcePermissions) { $sourcePermissions.Date.ToString('yyyy-MM-dd HH:mm') } else { '—' }
    $targetPermissionScanText = if ($targetPermissions) { $targetPermissions.Date.ToString('yyyy-MM-dd HH:mm') } else { '—' }
    $sourceScansSortDate = if ($sourceScan -and $sourcePermissions) {
        if ($sourceScan.Date -ge $sourcePermissions.Date) { $sourceScan.Date } else { $sourcePermissions.Date }
    } elseif ($sourceScan) { $sourceScan.Date } elseif ($sourcePermissions) { $sourcePermissions.Date } else { [datetime]::MinValue }
    $targetScansSortDate = if ($targetScan -and $targetPermissions) {
        if ($targetScan.Date -ge $targetPermissions.Date) { $targetScan.Date } else { $targetPermissions.Date }
    } elseif ($targetScan) { $targetScan.Date } elseif ($targetPermissions) { $targetPermissions.Date } else { [datetime]::MinValue }
    $scanGapDays = $null
    $scanGapText = '—'
    $scanGapTooltip = 'Gap unavailable: source or target scan is missing.'
    if ($sourceScan -and $targetScan) {
        $signedGap = ($targetScan.Date - $sourceScan.Date).TotalDays
        $scanGapDays = [math]::Abs($signedGap)
        $scanGapText = ('{0:N2} d' -f $scanGapDays)
        $newerSide = if ($signedGap -gt 0) { 'Target is newer' }
            elseif ($signedGap -lt 0) { 'Source is newer' }
            else { 'Scans were taken at the same time' }
        $scanGapTooltip = ('Absolute gap: {0:N2} days ({1:N1} hours). {2}.' -f
            $scanGapDays, ($scanGapDays * 24), $newerSide)
    }
    $permissionRate = $null
    $permissionRateText = '—'
    $permissionDetail = 'No valid permission comparison is available.'
    if ($permissionComparison) {
        if ($permissionComparison.Source -gt 0 -and $permissionComparison.Target -gt 0) {
            $permissionRate = [double]$permissionComparison.Matched / [double]$permissionComparison.Source * 100
            $permissionRateText = ('{0:N2} %' -f $permissionRate)
            $permissionDetail = "Matching permissions: $($permissionComparison.Matched) / $($permissionComparison.Source)."
        }
        else { $permissionDetail = 'Empty permission inventory: percentage is inconclusive.' }
        if (-not $sourcePermissions -or -not $targetPermissions) {
            $permissionDetail += ' Latest permission scan is unavailable.'
        }
        else {
            $permissionSourceName = [IO.Path]::GetFileName(([string]$permissionComparison.Summary.SourceCsv).Replace('/', '\'))
            $permissionTargetName = [IO.Path]::GetFileName(([string]$permissionComparison.Summary.TargetCsv).Replace('/', '\'))
            if (-not [string]::Equals($permissionSourceName, $sourcePermissions.File.Name, [StringComparison]::OrdinalIgnoreCase) -or
                -not [string]::Equals($permissionTargetName, $targetPermissions.File.Name, [StringComparison]::OrdinalIgnoreCase)) {
                $permissionDetail += ' Newer permission scans are available.'
            }
        }
        $scopeWarning = $permissionComparison.Summary.PSObject.Properties['ScopeWarning']
        if ($scopeWarning -and $scopeWarning.Value) {
            $permissionDetail += ' Scope warning in the report.'
        }
        $validationProperty = $permissionComparison.Summary.PSObject.Properties['ValidationStatus']
        $permissionValidation = if ($validationProperty) { [string]$validationProperty.Value } else { '' }
        if ($permissionValidation -and $permissionValidation -ne 'NoRelevantDifference') {
            $permissionDetail += " Validation: $permissionValidation."
        }
    }
    $fileStatus = $status
    $fileDetail = $detail
    $permissionStatus = 'Scan needed'
    if ($sourcePermissions -and $targetPermissions) {
        $permissionStatus = 'Compare needed'
        if (($null -ne $sourcePermissions.Rows -and $sourcePermissions.Rows -eq 0) -or
            ($null -ne $targetPermissions.Rows -and $targetPermissions.Rows -eq 0)) {
            $permissionStatus = 'Review needed'
            $permissionDetail += ' A published permission inventory contains zero rows.'
        }
        elseif ($permissionComparison) {
            $permissionSourceName = [IO.Path]::GetFileName(([string]$permissionComparison.Summary.SourceCsv).Replace('/', '\'))
            $permissionTargetName = [IO.Path]::GetFileName(([string]$permissionComparison.Summary.TargetCsv).Replace('/', '\'))
            $usesLatestPermissions = [string]::Equals($permissionSourceName, $sourcePermissions.File.Name, [StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals($permissionTargetName, $targetPermissions.File.Name, [StringComparison]::OrdinalIgnoreCase)
            if (-not $usesLatestPermissions) {
                $permissionDetail += ' Compare the latest permission scans.'
            }
            elseif ($null -eq $permissionRate) {
                $permissionStatus = 'Review needed'
            }
            else {
                $permissionStatus = if ($permissionValidation -eq 'NoRelevantDifference' -and
                    $permissionRate -ge 100 -and (-not $scopeWarning -or -not $scopeWarning.Value)) {
                    'Up to date'
                } else { 'Review needed' }
                $maxPermissionAge = 48.0; $maxPermissionGap = 24.0
                $configured = 0.0
                if ([double]::TryParse([string]$cfg.Comparison['PermissionMaxScanAgeHours'], [ref]$configured) -and $configured -gt 0) {
                    $maxPermissionAge = $configured
                }
                if ([double]::TryParse([string]$cfg.Comparison['PermissionMaxScanAgeDifferenceHours'], [ref]$configured) -and $configured -gt 0) {
                    $maxPermissionGap = $configured
                }
                $sourcePermissionAge = ((Get-Date) - $sourcePermissions.Date).TotalHours
                $targetPermissionAge = ((Get-Date) - $targetPermissions.Date).TotalHours
                $permissionGap = [math]::Abs(($sourcePermissions.Date - $targetPermissions.Date).TotalHours)
                if ($sourcePermissionAge -gt $maxPermissionAge -or $targetPermissionAge -gt $maxPermissionAge -or
                    $permissionGap -gt $maxPermissionGap -or
                    $sourcePermissionAge -lt -0.25 -or $targetPermissionAge -lt -0.25 -or
                    $sourcePermissions.Provenance -like '*unverified*' -or $targetPermissions.Provenance -like '*unverified*' -or
                    $sourcePermissions.Provenance -like '*invalid receipt*' -or $targetPermissions.Provenance -like '*invalid receipt*') {
                    $permissionStatus = 'Review needed'
                    $permissionDetail += ' Check permission scan freshness or provenance.'
                }
            }
        }
    }
    else {
        $permissionDetail += ' Source or target permission scan is missing.'
    }
    $status = @('Review needed', 'Scan needed', 'Compare needed', 'Up to date') |
        Where-Object { $_ -eq $fileStatus -or $_ -eq $permissionStatus } |
        Select-Object -First 1
    $statusTooltip = "Files: $fileStatus — $fileDetail`nPermissions: $permissionStatus — $permissionDetail"
    return [pscustomobject]@{
        Migration = $Migration.Name
        Source = $SourceScope
        SourceTooltip = $SourceTooltip
        SourceScan = $scanSourceText
        SourceScanTooltip = if ($sourceScan) { "$($sourceScan.File.Name)`nDate: $($sourceScan.Provenance)" } else { 'No source inventory.' }
        SourcePermissionScan = $sourcePermissionScanText
        SourceScansDisplay = "Files $scanSourceText`nPerms $sourcePermissionScanText"
        SourceScansSortDate = $sourceScansSortDate
        SourceScansTooltip = "Files: $(if ($sourceScan) { "$($sourceScan.File.Name) — $($sourceScan.Provenance)" } else { 'No source file inventory.' })`nPermissions: $(if ($sourcePermissions) { "$($sourcePermissions.File.Name) — $($sourcePermissions.Provenance)" } else { 'No source permission inventory.' })"
        Destination = $TargetScope
        DestinationTooltip = $TargetTooltip
        TargetScan = $scanTargetText
        TargetScanTooltip = if ($targetScan) { "$($targetScan.File.Name)`nDate: $($targetScan.Provenance)" } else { 'No target inventory.' }
        TargetPermissionScan = $targetPermissionScanText
        TargetScansDisplay = "Files $scanTargetText`nPerms $targetPermissionScanText"
        TargetScansSortDate = $targetScansSortDate
        TargetScansTooltip = "Files: $(if ($targetScan) { "$($targetScan.File.Name) — $($targetScan.Provenance)" } else { 'No target file inventory.' })`nPermissions: $(if ($targetPermissions) { "$($targetPermissions.File.Name) — $($targetPermissions.Provenance)" } else { 'No target permission inventory.' })"
        ScanGapDays = $scanGapDays
        ScanGapText = $scanGapText
        ScanGapTooltip = $scanGapTooltip
        ComparisonRate = $rate
        ComparisonPercent = $rateText
        ComparisonDate = if ($comparison) { $comparison.Date.ToString('yyyy-MM-dd HH:mm') } else { '—' }
        ComparisonDisplay = if ($comparison) { '{0} · {1}' -f $comparison.Date.ToString('yyyy-MM-dd HH:mm'), $rateText } else { '—' }
        ComparisonTooltip = if ($comparison) { "$($comparison.Path)`n$detail" } else { $detail }
        PermissionComparisonRate = $permissionRate
        PermissionComparisonPercent = $permissionRateText
        PermissionComparisonDate = if ($permissionComparison) { $permissionComparison.Date.ToString('yyyy-MM-dd HH:mm') } else { '—' }
        PermissionComparisonDisplay = if ($permissionComparison) { '{0} · {1}' -f $permissionComparison.Date.ToString('yyyy-MM-dd HH:mm'), $permissionRateText } else { '—' }
        PermissionComparisonTooltip = if ($permissionComparison) { "$($permissionComparison.Path)`n$permissionDetail" } else { $permissionDetail }
        Status = $status
        StatusTooltip = $statusTooltip
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAdaROqcqSlLxhu
# ZRltWHykXiHZroO4Ya7wTy3zuxn9VaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCzXN1D9XQ/U+l3V/LsSBBK
# O7EiKyzJcOswWjp3MKq76zANBgkqhkiG9w0BAQEFAASCAYCuLd0VFORhFzuNnoDB
# TwESbcFmr8DMUoc+bVucWj7EcdujjcyvhKkFgQEWmDNjWFdRogTvPzZtUWdq4YZ6
# Pqyr1kUb3pSpc4VlMykqjK1KuqOwxb4T+9ir7JOlhbYtz//zdQEiB0NZUM3Jmsey
# +Op7qv3RBTVYwyCA7Vb1plsM+vSBEbRlpKXRcSGBB3naHdxT2QzkcIcCTnbfZH/P
# nTkK0aKF+LDJQ2iF6tVKK868o3Q5ECVE6BADzT/GVZDUQmkAKeE844nze8BTD0FY
# wDNSDPMb1IbLtVKkttWSSohLVaacs4y3/wCvpvx1pzXcq+cNEZe6wbn9d5qBg/BB
# VTiVFWZxckihAlVYOerUN/KEwm9N+R0bkhsz0jz8B3fj2z9p8vJjRJOIOYhTH3HU
# 5314JJCNaJf13bBMiW3fhVFwvI0YOsQZv3M9uJQmCVzzz+fzDHRxYnWzkw/nzE0s
# lkEHoDr/NmpIl/xv5FdXSg5/Hh4o0bzvrp8ukrMRzHaMuEU=
# SIG # End signature block
