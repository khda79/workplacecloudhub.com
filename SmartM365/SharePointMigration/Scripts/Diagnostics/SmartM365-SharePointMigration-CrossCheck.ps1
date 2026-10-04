<#
.SYNOPSIS
    Read-only cross-check of ShareGate diagnostics and existing inventory comparisons.
.VERSION
    1.0.1
#>

function Get-SmartM365CrossCheckCsvRows {
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @() }
    $header = Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop
    $delimiter = if ($header.Contains(';')) { ';' } elseif ($header.Contains("`t")) { "`t" } else { ',' }
    return @(Import-Csv -LiteralPath $Path -Delimiter $delimiter -ErrorAction Stop)
}

function Get-SmartM365CrossCheckNumber {
    param($Row, [string]$Name)
    if (-not $Row) { return [long]0 }
    $property = $Row.PSObject.Properties[$Name]
    $number = [long]0
    if ($property -and [long]::TryParse([string]$property.Value, [ref]$number) -and $number -ge 0) { return $number }
    return [long]0
}

function Get-SmartM365CrossCheckScopeKey {
    param([string]$SiteUrl, [string]$ListTitle, [string]$Side = 'source')
    if ([string]::IsNullOrWhiteSpace($SiteUrl) -or [string]::IsNullOrWhiteSpace($ListTitle)) { return '' }
    $uri = $null
    if (-not [uri]::TryCreate($SiteUrl.Trim(), [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -notin @('http','https')) { return '' }
    $site = $uri.GetLeftPart([UriPartial]::Path).TrimEnd('/').ToLowerInvariant()
    $title = [regex]::Replace($ListTitle.Trim(), '\s+', ' ').ToLowerInvariant()
    return $Side + '|' + $site + '|' + $title
}

function Get-SmartM365CrossCheckDateText {
    param($Value)
    if ($Value -is [datetimeoffset]) { return $Value.LocalDateTime.ToString('yyyy-MM-dd HH:mm') }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
    $parsed = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.LocalDateTime.ToString('yyyy-MM-dd HH:mm')
    }
    return [string]$Value
}

function Get-SmartM365CrossCheckPermissionTargetWeb {
    param($Row, [string]$TargetSiteUrl)
    if ($Row.PSObject.Properties['TargetWebUrl'] -and $Row.TargetWebUrl) { return [string]$Row.TargetWebUrl }
    $sourceUri = $null; $targetUri = $null
    if (-not [uri]::TryCreate([string]$Row.SourceWebUrl, [UriKind]::Absolute, [ref]$sourceUri) -or
        -not [uri]::TryCreate($TargetSiteUrl, [UriKind]::Absolute, [ref]$targetUri)) { return '' }
    $sourcePath = $sourceUri.AbsolutePath.TrimEnd('/')
    $objectPath = if ($Row.PSObject.Properties['ComparisonObjectPath']) { [string]$Row.ComparisonObjectPath } else { '' }
    if (-not $objectPath.StartsWith('/')) { return '' }
    if (-not $sourcePath) { return $TargetSiteUrl }
    $match = [regex]::Match($objectPath, '(?i)^(?<web>.*?' + [regex]::Escape($sourcePath) + ')(?:/|$)')
    if (-not $match.Success) { return '' }
    $webPath = $match.Groups['web'].Value
    $targetPath = $targetUri.AbsolutePath.TrimEnd('/')
    if (-not ($webPath.Equals($targetPath, [StringComparison]::OrdinalIgnoreCase) -or
        $webPath.StartsWith($targetPath + '/', [StringComparison]::OrdinalIgnoreCase))) { return '' }
    return $targetUri.GetLeftPart([UriPartial]::Authority) + $webPath
}

function Get-SmartM365CrossCheckFileName {
    param([string]$Path)
    return [IO.Path]::GetFileName($Path.Replace('/', '\'))
}

function Get-SmartM365CrossCheckComparison {
    param($Migration, $Status, [ValidateSet('Files','Permissions')][string]$Kind)
    $outputKey = if ($Kind -eq 'Files') { 'FileComparisons' } else { 'PermissionComparisons' }
    $comparison = Get-SmartM365LatestPortfolioComparison `
        (Join-Path $Migration.Root $Migration.Config.Output[$outputKey]) $Migration.Name -Kind $Kind
    if (-not $comparison) {
        return [pscustomobject]@{ Kind=$Kind; Comparison=$null; State='Comparison missing'; Current=$false;
            DetailAvailable=$false; ReportPath=''; SummaryPath=''; DetailPath='' }
    }
    $summary = $comparison.Summary
    $sourceScan = if ($Kind -eq 'Files') { $Status.SourceFileCsv } else { $Status.SourcePermCsv }
    $targetScan = if ($Kind -eq 'Files') { $Status.TargetFileCsv } else { $Status.TargetPermCsv }
    $inputsExist = $sourceScan -and $targetScan
    $current = $inputsExist -and
        [string]::Equals((Get-SmartM365CrossCheckFileName ([string]$summary.SourceCsv)), $sourceScan.Name, [StringComparison]::OrdinalIgnoreCase) -and
        [string]::Equals((Get-SmartM365CrossCheckFileName ([string]$summary.TargetCsv)), $targetScan.Name, [StringComparison]::OrdinalIgnoreCase)
    $scanEvidence = [string]$summary.ScanEvidenceStatus
    $state = if (-not $inputsExist) { 'Scans unavailable' }
        elseif (-not $current) { 'Newer scans available' }
        elseif ($scanEvidence -eq 'Stale') { 'Stale scan evidence' }
        elseif ($scanEvidence -eq 'Unverified') { 'Unverified scan evidence' }
        elseif ($scanEvidence -eq 'LegacyFilename') { 'Legacy scan dates' }
        elseif ($scanEvidence -eq 'Verified') { 'Recorded scan evidence verified' }
        else { 'Scan evidence unknown' }
    $detailName = if ($Kind -eq 'Files') { 'LibrarySummary.csv' } else { 'PermissionSummary.csv' }
    $detailPath = Join-Path (Split-Path -Path $comparison.Path -Parent) $detailName
    $detailAvailable = Test-Path -LiteralPath $detailPath -PathType Leaf
    if (-not $detailAvailable) { $state += '; detail missing' }
    $report = [string]$summary.HtmlSummary
    if (-not $report -or -not (Test-Path -LiteralPath $report -PathType Leaf)) { $report = '' }
    return [pscustomobject]@{ Kind=$Kind; Comparison=$comparison; State=$state; Current=[bool]$current;
        DetailAvailable=[bool]$detailAvailable; ReportPath=$report; SummaryPath=$comparison.Path; DetailPath=$detailPath }
}

function Get-SmartM365DiagnosticCrossCheck {
    param(
        [Parameter(Mandatory)]$Migration,
        [Parameter(Mandatory)]$Status,
        $DiagnosticSummary,
        [object[]]$DiagnosticRows = @(),
        [bool]$DiagnosticVerified = $false
    )
    $files = Get-SmartM365CrossCheckComparison -Migration $Migration -Status $Status -Kind Files
    $permissions = Get-SmartM365CrossCheckComparison -Migration $Migration -Status $Status -Kind Permissions
    $evidence = [System.Collections.Generic.List[object]]::new()
    $hasDiagnostics = $null -ne $DiagnosticSummary
    if ($hasDiagnostics) {
        $toFixItems = [long]$DiagnosticSummary.IssueItemState['To fix']
        $toFixLines = [long]$DiagnosticSummary.IssueLineState['To fix']
        $dateText = Get-SmartM365CrossCheckDateText $DiagnosticSummary.GeneratedAtUtc
        $evidence.Add([pscustomobject]@{ Evidence='ShareGate'; Date=$dateText;
            Rate='—'; Coverage="$($DiagnosticSummary.DistinctItems) keyed items";
            Differences="$toFixItems items / $toFixLines lines to fix; $($DiagnosticSummary.UnkeyedRows) unkeyed lines";
            State=if ($DiagnosticVerified) { 'Latest report SHA256 verified' } else { 'Legacy or unverified analysis' };
            Path=[string]$DiagnosticSummary.ReportPath })
    }
    else {
        $evidence.Add([pscustomobject]@{ Evidence='ShareGate'; Date='—'; Rate='—'; Coverage='—';
            Differences='Analyze the latest ShareGate report'; State='Analysis missing'; Path='' })
    }
    foreach ($entry in @($files,$permissions)) {
        $comparison = $entry.Comparison
        if (-not $comparison) {
            $evidence.Add([pscustomobject]@{ Evidence=$entry.Kind; Date='—'; Rate='—'; Coverage='—';
                Differences='Run source and target scans, then compare'; State=$entry.State; Path='' })
            continue
        }
        $summary = $comparison.Summary
        $rate = if ($comparison.Source -gt 0) { '{0:N2} %' -f (100 * $comparison.Matched / $comparison.Source) } else { '—' }
        $differences = if ($entry.Kind -eq 'Files') {
            'Missing {0}; extra {1}; size {2}; modified {3}; older target {4}; version {5}' -f `
                (Get-SmartM365CrossCheckNumber $summary 'MissingInTarget'),
                (Get-SmartM365CrossCheckNumber $summary 'ExtraInTarget'),
                (Get-SmartM365CrossCheckNumber $summary 'DifferentSize'),
                (Get-SmartM365CrossCheckNumber $summary 'ChangedModifiedDate'),
                (Get-SmartM365CrossCheckNumber $summary 'TargetOlderThanSource'),
                (Get-SmartM365CrossCheckNumber $summary 'ChangedVersion')
        }
        else {
            'Missing {0}; disabled missing {1}; extra {2}; more {3}; less {4}; level changed {5}' -f `
                (Get-SmartM365CrossCheckNumber $summary 'MissingInSPO'),
                (Get-SmartM365CrossCheckNumber $summary 'DisabledEntraUsersNotInSPO'),
                (Get-SmartM365CrossCheckNumber $summary 'ExtraInSPO'),
                (Get-SmartM365CrossCheckNumber $summary 'TargetHasMorePermissions'),
                (Get-SmartM365CrossCheckNumber $summary 'TargetHasLessPermissions'),
                (Get-SmartM365CrossCheckNumber $summary 'PermissionLevelDifferent')
        }
        $evidence.Add([pscustomobject]@{ Evidence=$entry.Kind; Date=$comparison.Date.ToString('yyyy-MM-dd HH:mm');
            Rate=$rate; Coverage="$($comparison.Matched) / $($comparison.Source) source keys";
            Differences=$differences; State=$entry.State; Path=$entry.ReportPath })
    }

    $scopes = @{}
    $unmatched = [ordered]@{ ShareGate=0; Files=0; Permissions=0 }
    function Ensure-Scope {
        param([string]$Key, [string]$Site, [string]$List)
        if (-not $scopes.ContainsKey($Key)) {
            $scopes[$Key] = [pscustomobject]@{ Site=$Site; List=$List; ShareGateItems=[System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase);
                ShareGateLines=0; SourceListIds=[System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase);
                LibraryKeys=[System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase);
                FilesMissing=[long]0; FilesExtra=[long]0; FileChangeFlags=[long]0;
                PermsMissing=[long]0; PermsDisabled=[long]0; PermsExtra=[long]0; PermsChanged=[long]0 }
        }
        return $scopes[$Key]
    }
    foreach ($row in $DiagnosticRows) {
        if ($row.Status -notin @('Error','Warning') -or $row.State -ne 'To fix') { continue }
        $hasTarget = $row.PSObject.Properties['DestinationUrl'] -and $row.DestinationUrl
        $site = if ($hasTarget) { [string]$row.DestinationUrl } else { [string]$row.SourceUrl }
        $side = if ($hasTarget) { 'target' } else { 'source' }
        $key = Get-SmartM365CrossCheckScopeKey $site ([string]$row.SourceList) $side
        if (-not $key) { $unmatched.ShareGate++; continue }
        $scope = Ensure-Scope $key $site ([string]$row.SourceList)
        $scope.ShareGateLines++
        if ($row.ItemKey) { [void]$scope.ShareGateItems.Add([string]$row.ItemKey) }
        if ($row.SourceListId) { [void]$scope.SourceListIds.Add([string]$row.SourceListId) }
    }
    if ($files.Comparison -and $files.DetailAvailable) {
        foreach ($row in @(Get-SmartM365CrossCheckCsvRows $files.DetailPath)) {
            $missing = Get-SmartM365CrossCheckNumber $row 'MissingInTarget'
            $extra = Get-SmartM365CrossCheckNumber $row 'ExtraInTarget'
            $changeFlags = (Get-SmartM365CrossCheckNumber $row 'DifferentSize') +
                (Get-SmartM365CrossCheckNumber $row 'ChangedModifiedDate') +
                (Get-SmartM365CrossCheckNumber $row 'TargetOlderThanSource') +
                (Get-SmartM365CrossCheckNumber $row 'ChangedVersion')
            if (-not ($missing -or $extra -or $changeFlags)) { continue }
            $hasTarget = $row.PSObject.Properties['TargetWebUrl'] -and $row.TargetWebUrl
            $site = if ($hasTarget) { [string]$row.TargetWebUrl } else { [string]$row.SourceWebUrl }
            $side = if ($hasTarget) { 'target' } else { 'source' }
            $list = if ($row.PSObject.Properties['TargetLibraryTitle'] -and $row.TargetLibraryTitle) { [string]$row.TargetLibraryTitle } else { [string]$row.SourceLibraryTitle }
            $key = Get-SmartM365CrossCheckScopeKey $site $list $side
            if (-not $key) { $unmatched.Files++; continue }
            $scope = Ensure-Scope $key $site $list
            $scope.FilesMissing += $missing
            $scope.FilesExtra += $extra
            $scope.FileChangeFlags += $changeFlags
            if ($row.LibraryKey) { [void]$scope.LibraryKeys.Add([string]$row.LibraryKey) }
        }
    }
    if ($permissions.Comparison -and $permissions.DetailAvailable) {
        foreach ($row in @(Get-SmartM365CrossCheckCsvRows $permissions.DetailPath)) {
            $missing = Get-SmartM365CrossCheckNumber $row 'MissingInSPO'
            $disabled = Get-SmartM365CrossCheckNumber $row 'DisabledEntraUsersNotInSPO'
            $extra = Get-SmartM365CrossCheckNumber $row 'ExtraInSPO'
            $changed = (Get-SmartM365CrossCheckNumber $row 'TargetHasMorePermissions') +
                (Get-SmartM365CrossCheckNumber $row 'TargetHasLessPermissions') +
                (Get-SmartM365CrossCheckNumber $row 'PermissionLevelDifferent')
            if (-not ($missing -or $disabled -or $extra -or $changed)) { continue }
            $targetSiteUrl = if ($Migration.Config.ContainsKey('Target')) { [string]$Migration.Config.Target.SiteUrl } else { '' }
            $derivedTarget = Get-SmartM365CrossCheckPermissionTargetWeb $row $targetSiteUrl
            $site = if ($derivedTarget) { $derivedTarget } else { [string]$row.SourceWebUrl }
            $side = if ($derivedTarget) { 'target' } else { 'source' }
            $list = [string]$row.ListTitle
            $key = Get-SmartM365CrossCheckScopeKey $site $list $side
            if (-not $key) { $unmatched.Permissions++; continue }
            $scope = Ensure-Scope $key $site $list
            $scope.PermsMissing += $missing
            $scope.PermsDisabled += $disabled
            $scope.PermsExtra += $extra
            $scope.PermsChanged += $changed
        }
    }
    $scopeRows = [System.Collections.Generic.List[object]]::new()
    $ambiguous = 0
    foreach ($scope in $scopes.Values) {
        $isAmbiguous = $scope.SourceListIds.Count -gt 1 -or $scope.LibraryKeys.Count -gt 1
        if ($isAmbiguous) { $ambiguous++ }
        $sgCount = $scope.ShareGateItems.Count
        $gapCount = $scope.FilesMissing + $scope.FilesExtra + $scope.FileChangeFlags +
            $scope.PermsMissing + $scope.PermsDisabled + $scope.PermsExtra + $scope.PermsChanged
        $relationship = if ($scope.ShareGateLines -gt 0 -and $gapCount -gt 0) { 'Both report issues' }
            elseif ($gapCount -gt 0) { 'Comparison difference only' }
            else { 'ShareGate issue only' }
        $assessment = if ($isAmbiguous) { 'Ambiguous scope' }
            elseif (-not $hasDiagnostics) { 'ShareGate analysis missing' }
            elseif (-not $DiagnosticVerified -or $files.State -ne 'Recorded scan evidence verified' -or
                $permissions.State -ne 'Recorded scan evidence verified') { "$relationship; review evidence" }
            else { $relationship }
        $scopeRows.Add([pscustomobject]@{ Site=$scope.Site; List=$scope.List; ShareGateToFix="$sgCount items / $($scope.ShareGateLines) lines";
            FilesMissing=$scope.FilesMissing; FilesExtra=$scope.FilesExtra; FileChangeFlags=$scope.FileChangeFlags;
            PermsMissing=$scope.PermsMissing; PermsDisabled=$scope.PermsDisabled; PermsExtra=$scope.PermsExtra;
            PermsChanged=$scope.PermsChanged; Assessment=$assessment; SortWeight=if ($isAmbiguous) { 0 } else { $gapCount + $scope.ShareGateLines } })
    }
    return [pscustomobject]@{ Evidence=$evidence.ToArray(); Scopes=@($scopeRows | Sort-Object -Property @{ Expression='SortWeight'; Descending=$true },Site,List);
        Unmatched=$unmatched; Ambiguous=$ambiguous; FilesReport=$files.ReportPath; PermissionsReport=$permissions.ReportPath;
        FilesSummary=$files.SummaryPath; PermissionsSummary=$permissions.SummaryPath }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDChodBwGknKGFE
# CGi+zHFDQVD6BVsBuZD5re8EHc/opKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCUE8uRWyySQgDFNxSzjdBh
# v12JEKQIW5m0TUREA1IY/jANBgkqhkiG9w0BAQEFAASCAYAeq40krfrKGnb1pqT+
# W2uYwaSsoVmCgofzPxpzJiypXrgIH9Y0Wi7J2JNscG8azz4XgvASj2tzXAK2rlk2
# 0fcgZaQ5eh3ws8a0DImWGLw/2kGeVLvGwsxVROJDcpfajBrA488PQ1S8lSW6D2Gj
# vsw9AMFB49Kz0Jyq+8KrBA21a830RdPd2UOy1hIxwfHym+g/Z0qp/C6Bi4y9wEkj
# m2bnh/LAetKQDysGSOvTa0dHMSL2UDL5zLCyRSs3NwopnCZbbGcNdIGEZFlXpkbV
# Mv0G5eRKLhgbT8IDVAk/bWMC/jgVIGIqsSw3znpSHlau8+X5fhYt7ZVo4tw6i5x4
# jer42OLP4xrUEEuOoY3CnpkM7IBvZl9uDjHXNmlfFmZj52ZaXKvesauGv1F3DJnx
# EaR+jC7Yf1zVkqdkq4Gs5Ol1JpxZUhNMjxyR4NIXnmiqcmoznabb+8a6aljTh+p9
# 0HjB1yreMWBXwKg1vO00ks0JGLibqOWaUdqx10SAs2xde0Q=
# SIG # End signature block
