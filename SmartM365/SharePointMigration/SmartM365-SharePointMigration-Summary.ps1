<#
.SYNOPSIS
    Read-only portfolio summary for the SharePoint migration GUI.
.VERSION
    1.0.15
#>

. (Join-Path $PSScriptRoot 'Scripts\Launchers\SmartM365-SharePointMigration-LauncherCommon.ps1')

function Get-SmartM365PortfolioGapVisual {
    param($Days)
    # Use the unrounded gap so formatting to two decimals cannot hide a threshold crossing.
    $band = if ($null -eq $Days) { 'Unavailable' } elseif ($Days -gt 1.0) { 'Red' }
        elseif ($Days -gt 0.5) { 'Yellow' } else { 'Green' }
    $color = switch ($band) { 'Red' { '#B42335' } 'Yellow' { '#9A6700' } 'Green' { '#087F5B' } default { '#64748B' } }
    $background = switch ($band) { 'Red' { '#FFF0F2' } 'Yellow' { '#FFF8E8' } 'Green' { '#ECFAF4' } default { '#F0F4F8' } }
    [pscustomobject]@{ Band = $band; Color = $color; Background = $background }
}

function Get-SmartM365PortfolioRateVisual {
    param($Rate, $Date, [string]$Notice = '', [string]$Caption = '')
    $available = $null -ne $Rate
    $color = if (-not $available) { '#64748B' } elseif ($Rate -ge 100 -and -not $Notice) { '#087F5B' } else { '#9A6700' }
    $background = if (-not $available) { '#F0F4F8' } elseif ($Rate -ge 100 -and -not $Notice) { '#ECFAF4' } else { '#FFF8E8' }
    if (-not $Caption) {
        $Caption = if ($null -eq $Date) { 'No comparison' }
            elseif ($Date.Date -eq (Get-Date).Date) { $Date.ToString('HH:mm') + ' today' }
            elseif ($Date.Date -eq (Get-Date).Date.AddDays(-1)) { $Date.ToString('HH:mm') + ' yesterday' }
            else { $Date.ToString('yyyy-MM-dd HH:mm') }
    }
    [pscustomobject]@{
        Percent = if ($available) { '{0:N2} %' -f $Rate } else { '—' }
        Progress = if ($available) { [double]$Rate } else { 0.0 }
        ProgressVisibility = if ($available) { 'Visible' } else { 'Hidden' }
        Color = $color; Background = $background; Caption = $Caption; Notice = $Notice
        NoticeVisibility = if ($Notice) { 'Visible' } else { 'Collapsed' }
    }
}

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

function Test-SmartM365InventoryCsvComplete {
    param([System.IO.FileInfo]$File)
    if ($null -eq $File -or $File.Name -like '*-Errors.csv') { return $false }
    $errorPath = Join-Path -Path $File.DirectoryName -ChildPath ("{0}-Errors.csv" -f $File.BaseName)
    return -not (Test-Path -LiteralPath $errorPath -PathType Leaf)
}

function Get-SmartM365LatestPortfolioScan {
    param([string]$Directory, [string]$Filter)
    $scans = @(Get-SmartM365InventoryScans -Directory $Directory -Filter $Filter -Recurse)
    if ($scans.Count -eq 0) { return $null }
    return $scans[0]
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
        $path = Join-Path $folder.Folder.FullName ("{0}-Summary.csv" -f $MigrationName)
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            $path = Join-Path $folder.Folder.FullName 'Summary.csv'
        }
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
    $usesLatest = $false
    $usesLatestPermissions = $false
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
    $today = (Get-Date).Date
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
    $fileNotice = if ($comparison -and (-not $sourceScan -or -not $targetScan)) { 'Scan unavailable' }
        elseif ($comparison -and -not $usesLatest) { 'Recalculate' } else { '' }
    $permissionNotice = if ($permissionComparison -and (-not $sourcePermissions -or -not $targetPermissions)) { 'Scan unavailable' }
        elseif ($permissionComparison -and -not $usesLatestPermissions) { 'Recalculate' } else { '' }
    $globalRate = if ($null -ne $rate -and $null -ne $permissionRate) { ([double]$rate + [double]$permissionRate) / 2.0 } else { $null }
    $globalNotice = if ($null -eq $globalRate) { 'Both rates required' }
        elseif ($fileNotice -eq 'Scan unavailable' -or $permissionNotice -eq 'Scan unavailable') { 'Scan unavailable' }
        elseif ($fileNotice -or $permissionNotice) { 'Recalculate' } else { '' }
    $globalTooltip = "Equal-weight average: (file comparison % + permission comparison %) / 2. Each rate retains its own source denominator.`nFiles: $rateText; permissions: $permissionRateText."
    if ($globalNotice) { $globalTooltip += "`n$globalNotice." }
    $fileDate = if ($comparison) { $comparison.Date } else { $null }
    $permissionDate = if ($permissionComparison) { $permissionComparison.Date } else { $null }
    return [pscustomobject]@{
        Migration = $Migration.Name
        Source = $SourceScope
        SourceTooltip = $SourceTooltip
        SourceScan = $scanSourceText
        SourceFileScanFile = if ($sourceScan) { $sourceScan.File } else { $null }
        SourceFileScanDate = if ($sourceScan) { $sourceScan.Date } else { $null }
        SourceFileScanProvenance = if ($sourceScan) { $sourceScan.Provenance } else { '' }
        SourceInventoryDisplay = '—'
        SourceInventoryTooltip = 'No source file inventory.'
        SourceScanTooltip = if ($sourceScan) { "$($sourceScan.File.Name)`nDate: $($sourceScan.Provenance)" } else { 'No source inventory.' }
        SourcePermissionScan = $sourcePermissionScanText
        SourceFileScanIsToday = [bool]($sourceScan -and $sourceScan.Date.Date -eq $today)
        SourcePermissionScanIsToday = [bool]($sourcePermissions -and $sourcePermissions.Date.Date -eq $today)
        SourceScansDisplay = "Files $scanSourceText`nPerms $sourcePermissionScanText"
        SourceScansSortDate = $sourceScansSortDate
        SourceScansTooltip = "Files: $(if ($sourceScan) { "$($sourceScan.File.Name) — $($sourceScan.Provenance)" } else { 'No source file inventory.' })`nPermissions: $(if ($sourcePermissions) { "$($sourcePermissions.File.Name) — $($sourcePermissions.Provenance)" } else { 'No source permission inventory.' })"
        Destination = $TargetScope
        DestinationTooltip = $TargetTooltip
        TargetScan = $scanTargetText
        TargetFileScanFile = if ($targetScan) { $targetScan.File } else { $null }
        TargetFileScanDate = if ($targetScan) { $targetScan.Date } else { $null }
        TargetFileScanProvenance = if ($targetScan) { $targetScan.Provenance } else { '' }
        TargetInventoryDisplay = '—'
        TargetInventoryTooltip = 'No target file inventory.'
        TargetScanTooltip = if ($targetScan) { "$($targetScan.File.Name)`nDate: $($targetScan.Provenance)" } else { 'No target inventory.' }
        TargetPermissionScan = $targetPermissionScanText
        TargetFileScanIsToday = [bool]($targetScan -and $targetScan.Date.Date -eq $today)
        TargetPermissionScanIsToday = [bool]($targetPermissions -and $targetPermissions.Date.Date -eq $today)
        TargetScansDisplay = "Files $scanTargetText`nPerms $targetPermissionScanText"
        TargetScansSortDate = $targetScansSortDate
        TargetScansTooltip = "Files: $(if ($targetScan) { "$($targetScan.File.Name) — $($targetScan.Provenance)" } else { 'No target file inventory.' })`nPermissions: $(if ($targetPermissions) { "$($targetPermissions.File.Name) — $($targetPermissions.Provenance)" } else { 'No target permission inventory.' })"
        ScanGapDays = $scanGapDays
        ScanGapText = $scanGapText
        ScanGapTooltip = "$scanGapTooltip`nGreen: up to 12 hours. Yellow: over 12 hours, up to 24 hours. Red: over 24 hours."
        ScanGapVisual = Get-SmartM365PortfolioGapVisual -Days $scanGapDays
        ComparisonRate = $rate
        ComparisonPercent = $rateText
        ComparisonDate = if ($comparison) { $comparison.Date.ToString('yyyy-MM-dd HH:mm') } else { '—' }
        ComparisonDisplay = if ($comparison) { '{0} · {1}' -f $comparison.Date.ToString('yyyy-MM-dd HH:mm'), $rateText } else { '—' }
        ComparisonTooltip = if ($comparison) { "Date: $($comparison.Date.ToString('yyyy-MM-dd HH:mm'))`n$($comparison.Path)`n$detail" } else { $detail }
        ComparisonVisual = Get-SmartM365PortfolioRateVisual -Rate $rate -Date $fileDate -Notice $fileNotice
        PermissionComparisonRate = $permissionRate
        PermissionComparisonPercent = $permissionRateText
        PermissionComparisonDate = if ($permissionComparison) { $permissionComparison.Date.ToString('yyyy-MM-dd HH:mm') } else { '—' }
        PermissionComparisonDisplay = if ($permissionComparison) { '{0} · {1}' -f $permissionComparison.Date.ToString('yyyy-MM-dd HH:mm'), $permissionRateText } else { '—' }
        PermissionComparisonTooltip = if ($permissionComparison) { "Date: $($permissionComparison.Date.ToString('yyyy-MM-dd HH:mm'))`n$($permissionComparison.Path)`n$permissionDetail" } else { $permissionDetail }
        PermissionComparisonVisual = Get-SmartM365PortfolioRateVisual -Rate $permissionRate -Date $permissionDate -Notice $permissionNotice
        GlobalComparisonRate = $globalRate
        GlobalComparisonVisual = Get-SmartM365PortfolioRateVisual -Rate $globalRate -Notice $globalNotice -Caption 'Equal weighting'
        GlobalComparisonTooltip = $globalTooltip
        Status = $status
        StatusTooltip = $statusTooltip
        SourcePermissionScanFile = if ($sourcePermissions) { $sourcePermissions.File } else { $null }
        ComplexityScore = $null
        ComplexityText = '—'
        ComplexityCaption = ''
        ComplexityVisual = Get-SmartM365ComplexityVisual
        ComplexityTooltip = 'Complexity has not been evaluated.'
    }
}

function Get-SmartM365ComplexityVisual {
    param([string]$Color = '#64748B', [string]$Background = '#F0F4F8')
    [pscustomobject]@{ Color = $Color; Background = $Background }
}

function Get-SmartM365ComplexityEntraCache {
    param([Parameter(Mandatory)]$Migration, [Parameter(Mandatory)][string]$ProjectRoot)
    $comparison = $Migration.Config['Comparison']
    $configured = if ($comparison) { [string]$comparison['EntraUsersCachePath'] } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($configured)) {
        $path = if ([IO.Path]::IsPathRooted($configured)) { $configured } else { Join-Path $Migration.Root $configured }
        if (Test-Path -LiteralPath $path -PathType Leaf) { return Get-Item -LiteralPath $path }
        return $null
    }
    $directory = Join-Path $ProjectRoot '_Local\EntraUsersCache'
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { return $null }
    return Get-ChildItem -LiteralPath $directory -Filter '*.csv' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
}

function Get-SmartM365ComplexityRequest {
    <# Inputs, arguments and signature of the source-only complexity score for one migration. #>
    param(
        [Parameter(Mandatory)]$Migration,
        [Parameter(Mandatory)][string]$ProjectRoot,
        [System.IO.FileInfo]$FileScan,
        [System.IO.FileInfo]$PermissionScan
    )
    if (-not $FileScan) { return $null }
    $cfg = $Migration.Config
    $entraCache = if ($PermissionScan) { Get-SmartM365ComplexityEntraCache -Migration $Migration -ProjectRoot $ProjectRoot } else { $null }
    $mapping = $null
    $mappingName = if ($cfg['Comparison']) { [string]$cfg['Comparison']['PathMappingsFile'] } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($mappingName)) {
        $mappingPath = if ([IO.Path]::IsPathRooted($mappingName)) { $mappingName } else { Join-Path $Migration.Root $mappingName }
        if (Test-Path -LiteralPath $mappingPath -PathType Leaf) { $mapping = Get-Item -LiteralPath $mappingPath }
    }
    $targetSiteUrl = if ($cfg['Target']) { [string]$cfg['Target']['SiteUrl'] } else { '' }
    $configDirectory = Join-Path $ProjectRoot 'Config'
    $configPath = Join-Path $configDirectory 'migration-complexity.json.txt'
    $configHash = if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        (Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash
    } else { '' }
    $output = Join-Path $Migration.Root 'migration-complexity.json.txt'
    $arguments = @(
        (Join-Path $ProjectRoot 'Scripts\Compare\migration_complexity.py'),
        '--migration-name', $Migration.Name,
        '--file-inventory', $FileScan.FullName,
        '--config-dir', $configDirectory,
        '--output', $output
    )
    if ($PermissionScan) { $arguments += @('--permission-inventory', $PermissionScan.FullName) }
    if ($entraCache) { $arguments += @('--entra-users-csv', $entraCache.FullName) }
    if ($mapping) { $arguments += @('--path-mapping-file', $mapping.FullName) }
    if ($targetSiteUrl) { $arguments += @('--target-site-url', $targetSiteUrl) }
    $inputKeys = foreach ($item in @($FileScan, $PermissionScan, $entraCache, $mapping)) {
        if ($item) { '{0}:{1}' -f $item.Name, $item.Length } else { '' }
    }
    $signature = (@($Migration.Name, $configHash, $targetSiteUrl) + @($inputKeys)) -join '|'
    [pscustomobject]@{
        Migration = $Migration.Name; Output = $output; Arguments = $arguments; Signature = $signature
        FileScan = $FileScan; PermissionScan = $PermissionScan; EntraCache = $entraCache; Mapping = $mapping
        TargetSiteUrl = $targetSiteUrl; ConfigHash = $configHash
    }
}

function Test-SmartM365ComplexityInput {
    param($Recorded, [System.IO.FileInfo]$Expected)
    if (-not $Expected) { return $null -eq $Recorded }
    return $null -ne $Recorded -and [string]$Recorded.Name -ceq $Expected.Name -and [long]$Recorded.Length -eq $Expected.Length
}

function Get-SmartM365ComplexityPresentation {
    <# Reads the stored complexity result and states whether it matches the current inputs. #>
    param($Request, [string]$ErrorMessage = '', [switch]$ReadOnly)
    $unavailable = [pscustomobject]@{
        State = 'Unavailable'; Score = $null; Text = 'Scan needed'; Caption = 'No source scan'
        Visual = Get-SmartM365ComplexityVisual; Tooltip = 'Complexity needs a source file inventory scan.'
        Summary = 'Run a source file scan to evaluate complexity.'; Criteria = @(); Result = $null
    }
    if (-not $Request) { return $unavailable }
    $result = $null
    if (Test-Path -LiteralPath $Request.Output -PathType Leaf) {
        try { $result = Get-Content -LiteralPath $Request.Output -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
        catch { $result = $null }
    }
    $current = $false
    if ($result -and [int]$result.SchemaVersion -eq 1 -and $result.Inputs) {
        $inputs = $result.Inputs
        $current = (Test-SmartM365ComplexityInput $inputs.FileInventory $Request.FileScan) -and
            (Test-SmartM365ComplexityInput $inputs.PermissionInventory $Request.PermissionScan) -and
            (Test-SmartM365ComplexityInput $inputs.EntraUsers $Request.EntraCache) -and
            (Test-SmartM365ComplexityInput $inputs.PathMapping $Request.Mapping) -and
            [string]$inputs.TargetSiteUrl -ceq $Request.TargetSiteUrl -and
            $Request.ConfigHash -and [string]$inputs.ConfigSha256 -eq $Request.ConfigHash
    }
    if (-not $result -and $ReadOnly) {
        return [pscustomobject]@{
            State = 'Missing'; Score = $null; Text = '—'; Caption = 'Not evaluated'
            Visual = Get-SmartM365ComplexityVisual; Tooltip = 'No complexity result is stored for this migration.'
            Summary = 'No complexity result is stored for this migration.'; Criteria = @(); Result = $null
        }
    }
    if (-not $result) {
        $state = if ($ErrorMessage) { 'Error' } else { 'Missing' }
        return [pscustomobject]@{
            State = $state; Score = $null
            Text = if ($ErrorMessage) { 'Error' } else { 'Computing' }
            Caption = if ($ErrorMessage) { 'See tooltip' } else { 'Background job' }
            Visual = Get-SmartM365ComplexityVisual
            Tooltip = if ($ErrorMessage) { "Complexity calculation failed:`n$ErrorMessage" } else { 'Complexity is being calculated from the latest source scans.' }
            Summary = if ($ErrorMessage) { "Calculation failed: $ErrorMessage" } else { 'Calculating complexity from the latest source scans.' }
            Criteria = @(); Result = $null
        }
    }
    $names = @{}
    foreach ($criterion in @($result.Criteria)) { $names[[string]$criterion.Id] = [string]$criterion.Name }
    $drivers = @($result.TopDrivers | ForEach-Object { $names[[string]$_] }) -join ', '
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(('Complexity: {0} ({1}/100). Source scans only; independent from comparison progress.' -f $result.Level, $result.Score))
    $lines.Add('Top drivers: ' + $(if ($drivers) { $drivers } else { 'none' }))
    if (@($result.FloorApplied).Count) {
        $lines.Add(('Floor rule applied ({0}): score level {1} raised to {2}.' -f
            (@($result.FloorApplied | ForEach-Object { $names[[string]$_] }) -join ', '), $result.ScoreLevel, $result.Level))
    }
    if (@($result.MissingInputs).Count) {
        $lines.Add('Partial: not evaluated without ' + (@($result.MissingInputs) -join ', ') + '.')
    }
    $lines.Add(('Files: {0}' -f $result.Inputs.FileInventory.Name))
    $lines.Add(('Permissions: {0}' -f $(if ($result.Inputs.PermissionInventory) { $result.Inputs.PermissionInventory.Name } else { 'none' })))
    $computed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$result.ComputedAtUtc, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::RoundtripKind, [ref]$computed)) {
        $lines.Add('Computed: ' + $computed.ToLocalTime().ToString('yyyy-MM-dd HH:mm'))
    }
    $state = 'Current'
    $caption = '{0}/100' -f $result.Score
    if ([string]$result.Status -eq 'Partial') { $caption += ' · partial' }
    if (-not $current) {
        $state = if ($ErrorMessage) { 'Error' } else { 'Stale' }
        $caption = if ($ErrorMessage) { 'Update failed' } elseif ($ReadOnly) { 'Outdated' } else { 'Updating' }
        $lines.Insert(0, $(if ($ErrorMessage) { "Update failed; showing the previous result.`n$ErrorMessage" }
            elseif ($ReadOnly) { 'Source scans or settings changed after this result; read-only views do not recalculate.' }
            else { 'Newer source scans or settings: recalculating in the background; showing the previous result.' }))
    }
    $criteria = foreach ($criterion in @($result.Criteria)) {
        $bands = @($criterion.Bands)
        [pscustomobject]@{
            Axis = [string]$criterion.Axis; Criterion = [string]$criterion.Name
            Value = [string]$criterion.ValueDisplay; Band = [string]$criterion.Band
            Points = if ($criterion.Evaluated) { '{0} / 3' -f $criterion.Points } else { 'N/A' }
            Weighted = if ($criterion.Evaluated) { '{0} / {1}' -f $criterion.WeightedPoints, $criterion.MaxWeightedPoints } else { 'N/A' }
            Detail = [string]$criterion.Detail
            Thresholds = @(for ($index = 0; $index -lt $bands.Count; $index++) { '{0} pt: {1}' -f $index, $bands[$index] }) -join "`n"
        }
    }
    [pscustomobject]@{
        State = $state; Score = [int]$result.Score; Text = [string]$result.Level; Caption = $caption
        Visual = Get-SmartM365ComplexityVisual -Color ([string]$result.LevelColor) -Background ([string]$result.LevelBackground)
        Tooltip = $lines -join "`n"
        Summary = ('{0} · {1}/100 · top drivers: {2}' -f $result.Level, $result.Score, $(if ($drivers) { $drivers } else { 'none' }))
        Criteria = @($criteria); Result = $result
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAEEzNTAh1An0lu
# Q7KtlsgR/RHv+qJ+O1OaXCIFPkY47qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJdwYUbc5CL1e4PGFIT10cUO80YvmCS1gkvkWnOUZoCSMA0GCSqG
# SIb3DQEBAQUABIIBgJ9L1R/421jFPTLpz+qbRa+UqGRjsiYQb7ECBxVsOyOn/tUj
# FEyx/bmCX4LJM1iZOlKPT9Yock6FYhWx/FXBVeaJVyp9zDqo+f5fN/gpC0KfAtrV
# CjePnJG6CCRLXPG5kNnnn17AHvdj8eiBrnbLe6k+jWitAHLt8aw06HfPQ/aP4y7L
# 6mxUfjRjTFOfxrtUjyI/GNh4r5XM6KHoM8qjNsw5/Y1goBeBKt4reakhEHEfzhat
# 2a2oNhBz7FODQzzc2BH8+8V/+WXyIg/AJv2nwCZazZVk/wBUnNDxWIqNbAR4i/0H
# 3zUNukOfM7jd7H2NNJnbRzkavLeuSq8CWAblwI8YibEFChhpDj87QC8maiShRsC/
# b+y47klKC+NJ4pUAnXO/PUfSggyl/oRUPfT6liVm60+5gtfjc0iXlBEv92+k6KqG
# px8M83P3IW/+jUCIMJUOJ7IZ1+mQO//pntQ82p3emi09TgFWhnv8y/f7CF33AuYB
# aqSxzn3iPSAShmMSpaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTcwNzE0
# NTZaMC8GCSqGSIb3DQEJBDEiBCBKVVQAoArJ3SPoGRyiw7JWg7gpwfLcDZDCnaC9
# 1ucU+zANBgkqhkiG9w0BAQEFAASCAgATAnbCnRm80zTxKl2rlIhd839hTmADtG0/
# 7AHQGuhzuQQISBkKvUtgUNy+BV3Pg8cVpFYtjetjVyaPVzFPeKyjztUTsLqa3sFu
# j7cqHVHPpoRkKmIDycO/VRTcYYf9eJgCm/2OZVKnJPD+acYRCThi8Np2vALKaCOq
# 0RIXTbyaagWdd4rQL6P2/0Iy5SnCBfQMjlwv0aAe1O8+l3lI7FJFIRy+Z0sTLoiH
# RZx/OgOCa6Lsf8T3Im3v852o3WQB3r23BI9Wq5m7TwEUPzJ0SToEnmYoLwXVVZ07
# oLw3e8Fnitb6kG60b2ShD1ez09qn2AVhGQS16QzdwVxrHP+J2arH8leDJfTSBjZq
# /ks3Age1bFYpkYxayDLt6EDy2L4jmyj+3yQA+gZ/MjVl5DGdYkdOi1Y0J9t0lOLe
# rOOqwbksa8R4qoBTmet/QzKdb7lVPZ6IBRZFAWG7GSzyO54NpraaVWMI7J0bbcwK
# hwSi3ueBCkLXRihi1CJn4C/TsBuX+GD/5Bd//XE6MODKTjVL4p43pPVoqys0zFyi
# IuP2/QQhAslc3khYwK729c8V16AOTm7cJFcr5vsXDkR+7UmePukC9VMg1gS0zzvZ
# 4s0iUHuPxPFT8qThZSASpCOdpp5wCK3O4+cvAbXJELN/83WRUxzI0ymraFD8TIfv
# iNfr1zMrog==
# SIG # End signature block
