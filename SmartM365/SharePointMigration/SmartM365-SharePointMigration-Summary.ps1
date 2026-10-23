<#
.SYNOPSIS
    Read-only portfolio summary for the SharePoint migration GUI.
.VERSION
    1.0.19
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

# Get-SmartM365ComplexityRequest lives in LauncherCommon so source scans can score complexity too.
function Test-SmartM365ComplexityInput {
    param($Recorded, [System.IO.FileInfo]$Expected, [switch]$CompareTimestamp)
    if (-not $Expected) { return $null -eq $Recorded }
    if ($null -eq $Recorded -or [string]$Recorded.Name -cne $Expected.Name -or [long]$Recorded.Length -ne $Expected.Length) { return $false }
    if (-not $CompareTimestamp) { return $true }
    # Rewritten files keep their name (Entra cache, mapping); compare at second precision.
    $recordedTicks = [long]$Recorded.LastWriteTimeUtcTicks
    $expectedTicks = $Expected.LastWriteTimeUtc.Ticks
    $second = [TimeSpan]::TicksPerSecond
    return ($recordedTicks - ($recordedTicks % $second)) -eq ($expectedTicks - ($expectedTicks % $second))
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
            (Test-SmartM365ComplexityInput $inputs.EntraUsers $Request.EntraCache -CompareTimestamp) -and
            (Test-SmartM365ComplexityInput $inputs.PathMapping $Request.Mapping -CompareTimestamp) -and
            [string]$inputs.TargetSiteUrl -ceq $Request.TargetSiteUrl -and
            $Request.ConfigHash -and [string]$inputs.ConfigSha256 -eq $Request.ConfigHash -and
            $inputs.PSObject.Properties['SourceAssessmentSha256'] -and
            [string]$inputs.SourceAssessmentSha256 -eq [string]$Request.AssessmentHash -and
            $inputs.PSObject.Properties['AutomaticAssessmentHashes'] -and
            [string]$inputs.AutomaticAssessmentHashes.File -eq [string]$Request.AutomaticAssessmentHashes.File -and
            [string]$inputs.AutomaticAssessmentHashes.Permission -eq [string]$Request.AutomaticAssessmentHashes.Permission -and
            $inputs.AutomaticAssessmentHashes.PSObject.Properties['FileDetails'] -and
            [string]$inputs.AutomaticAssessmentHashes.FileDetails -eq [string]$Request.AutomaticAssessmentHashes.FileDetails -and
            [string]$inputs.AutomaticAssessmentHashes.PermissionDetails -eq [string]$Request.AutomaticAssessmentHashes.PermissionDetails
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
    $coverage = if ([double]$result.TotalWeight -gt 0) { [math]::Round(100 * [double]$result.EvaluatedWeight / [double]$result.TotalWeight, 1) } else { 0 }
    $coverageText = 'Coverage: {0}% of enabled weights evaluated ({1}/{2}).' -f $coverage, $result.EvaluatedWeight, $result.TotalWeight
    $lines.Add($coverageText)
    if ($result.PSObject.Properties['FunctionalEstimate'] -and $result.FunctionalEstimate) {
        $lines.Add('Functional effort is an automatic technical estimate, not functional acceptance.')
    }
    foreach ($criterion in @($result.Criteria)) {
        if ($criterion.PSObject.Properties['Assessment']) {
            $assessment = $criterion.Assessment
            $lines.Add(('{0}: collection={1}; qualification={2}.' -f $criterion.Name,$assessment.CollectionComplete,$assessment.QualificationComplete))
            if ($assessment.PSObject.Properties['KnownRisks'] -and @($assessment.KnownRisks).Count) {
                $lines.Add('Known functional risks: ' + (@($assessment.KnownRisks | ForEach-Object { '{0} (effort {1})' -f $_.Technology,$_.Effort }) -join ', '))
            }
        }
    }
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
    $caption += ' · {0}% coverage' -f $coverage
    if ($result.PSObject.Properties['FunctionalEstimate'] -and $result.FunctionalEstimate) { $caption += ' · estimate' }
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
            Thresholds = if ($criterion.PSObject.Properties['Assessment']) {
                [string]$criterion.Detail + "`n" + ($criterion.Assessment | ConvertTo-Json -Depth 12)
            } elseif ($criterion.PSObject.Properties['Components'] -and @($criterion.Components).Count) {
                [string]$criterion.Detail + "`nMAX; weight counted once. " + (@($criterion.Components | ForEach-Object {
                    '{0}: {1}' -f $_.Name, (@(for ($index = 0; $index -lt @($_.Bands).Count; $index++) { '{0} pt: {1}' -f $index, $_.Bands[$index] }) -join '; ')
                }) -join "`n")
            } else { @(for ($index = 0; $index -lt $bands.Count; $index++) { '{0} pt: {1}' -f $index, $bands[$index] }) -join "`n" }
        }
    }
    [pscustomobject]@{
        State = $state; Score = [int]$result.Score; Text = [string]$result.Level; Caption = $caption
        Visual = Get-SmartM365ComplexityVisual -Color ([string]$result.LevelColor) -Background ([string]$result.LevelBackground)
        Tooltip = $lines -join "`n"
        Summary = ('{0} · {1}/100 · {2} · top drivers: {3}' -f $result.Level, $result.Score, $coverageText, $(if ($drivers) { $drivers } else { 'none' }))
        Criteria = @($criteria); Result = $result
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDx+h/kpowtNBVx
# JAp1E5ouA7RZza4SkT6UlxZrf/kVdqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAA1fHnuwHMSc+E5mpUNM2N
# DJRtz8RoHUoXpqaraBDXhzANBgkqhkiG9w0BAQEFAASCAYB8Tg6XnY5mu/aI5DDR
# Ph6AnV57ufOX6aKii79sObOYff24rKqK1HfSydFJQWROEISNYqKEuQAKlfUkX8AQ
# BC0c3Q3Yu2zBdQZNM+LisnesrXVa0EHCiWjxXirLB7610j76FBzNL7cZuyzxe8BN
# 0L6iUMshrhFet0KtFApczV8en+baIv0y1Bb2K7Pf4ZNxpaZ5uEQEa0V2w2QiAPER
# Wir7c6tEPVaMcoBSZAd1ePMbSjf2aY/1fOjUidJABCqAQ98vBcgCxVFZ0+/QheLp
# BDnbO59uyM9vIRbabw6iAjuxDkUk3PIZW2Yk4rt/PJJqiVOyd8fTyb+02bu0UTzb
# 1TOM4iRYKdAQ7wup9s6Hxy9DU3ut6ubageqUpgRU7jpz08MIW3Hq+rPOsvXGHIlT
# AAaqdoLPAEgwV71DSMyXMf6nRmsVUBG3LeI9iZ9LDZ5NiqvsMk2TbNkdEvVEbExU
# A8GmfsfMAXniJ9M7dL375PvwOorJ3ybZ0mgjUL+pbf3JVBs=
# SIG # End signature block
