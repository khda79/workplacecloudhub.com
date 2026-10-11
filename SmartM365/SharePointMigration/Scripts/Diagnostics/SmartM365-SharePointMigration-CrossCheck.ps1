<#
.SYNOPSIS
    Read-only cross-check of ShareGate diagnostics and existing inventory comparisons.
.VERSION
    1.0.2
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
    $summaryName = [IO.Path]::GetFileName([string]$comparison.Path)
    $prefix = if ($summaryName -match '^(.*-)Summary\.csv$') { $Matches[1] } else { '' }
    $detailPath = Join-Path (Split-Path -Path $comparison.Path -Parent) ($prefix + $detailName)
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
        $scopeRows.Add([pscustomobject]@{ Site=$scope.Site; List=$scope.List; ShareGateToFix=$(if ($hasDiagnostics) { "$sgCount items / $($scope.ShareGateLines) lines" } else { '—' });
            FilesMissing=$scope.FilesMissing; FilesExtra=$scope.FilesExtra; FileChangeFlags=$scope.FileChangeFlags;
            PermsMissing=$scope.PermsMissing; PermsDisabled=$scope.PermsDisabled; PermsExtra=$scope.PermsExtra;
            PermsChanged=$scope.PermsChanged; Assessment=$assessment; SortWeight=if ($isAmbiguous) { 0 } else { $gapCount + $scope.ShareGateLines } })
    }
    return [pscustomobject]@{ Evidence=$evidence.ToArray(); Scopes=@($scopeRows | Sort-Object -Property @{ Expression='SortWeight'; Descending=$true },Site,List);
        Unmatched=$unmatched; Ambiguous=$ambiguous; FilesReport=$files.ReportPath; PermissionsReport=$permissions.ReportPath;
        FilesSummary=$files.SummaryPath; PermissionsSummary=$permissions.SummaryPath }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCjZJHxYGHfp2nb
# YtNv5fETf6M1mz5u2zvYOJW1jvALVaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIhfSOjfF/OOPq+rdRr36xx529i3X0kriH9b9BKP4HqpMA0GCSqG
# SIb3DQEBAQUABIIBgGXTpuV97VAa9cxEwdYP+miC6+OQizefg0/zhIBKEKa1mFOv
# SElkUis8ikmXvymZXtJQhlWeUF9z+EUW16R65i6lwioTTuYKat3cBjZf/xDfWayV
# bLXATsannwq/gSpTNMq9dvkc2C01U6RVPaBfcRvzH3vQAFqjipTzgD/siji64CGf
# sjC1DSfSn9PgU/j9pL168NB4ad6Zjhzor+BzNyteun0/tuUvo2S7mLMd1t88IQaY
# Q+pQ4U9UQ15Xq3XkyxVxkW73CAiNJSFZW/bb6L2EWiVVrK23H1REb53TMHe6Wl65
# jTX0m+BMzUGASHWJwqf9owN1J+EqypTwv89mJQ4Fd0+Pewox89BIiN9wYkHbJdIp
# DYZdX+PqsJPn5GBtYZtWAQTlORcOrJmurCjoNcmMwUL4Y4aYRERpJfZEvxI5BCk1
# XF0cgpr6AHwAhk903R8Q9otbv4UpHvQQAtwbArUKcynLroDDyl0Kwfgx9bpxjS1g
# SHz4M2yjSwNX89oiZKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTEwOTM1
# NDZaMC8GCSqGSIb3DQEJBDEiBCBM5/Bgv4Zi53aArWUmaohKp3DAxRMQ5gnMcYZX
# 9V72MzANBgkqhkiG9w0BAQEFAASCAgAumwXJ9FAPOAzpA3oUvZnQWqODvof6XczZ
# XuBrZNqy1dcFZ8jhnzFP6VHsQFQMd0SdCJSv57ZxsmBBNKt7AeC3nQOs7KXw1h9G
# Ht8bV7n7iJIFs+dpzBBrAzdErVAjHReU8Vj4l7mYyQog3mCbycrrPB2cp4oyOygE
# TYVvyXh0C2FUTimoe2/Tn5OoUq7mbdeJVkFfTa/03AfzSZ421BIvDDjdw6uPEMB3
# /ll4XnxTAL6z9/BtkIusQiMZIspS28JXb6cEuMQ284OtOO/jtNdTd0Q1egJqP97C
# LAWJxulvQCwubwTfM5n7RjAuTvLSBdlFHPNNe6ZPqbkLVw1YmjqaRUwYoZZdPoqA
# JxSTEGKg3PkahFCQLaZWjphfB5etDZ9mIuEgzGO9+KeZmzHtKJChqQoU1WNbEUU9
# GMTm1rOpgmvpDonwfXN7s5ibUy+L+1VjeFokGHdzSlBaNddyAiJjxWNRNizuVJeK
# VawBz2rbdg9zlxfzkTES/6gje1JKI11ensigEgza5Yoyb3bQ+HIjAPRk1ccIzV+c
# Ur9QLOE1L3OWeayMTj3En9aaeWyMeMBoHXuQcH7zBNgm+bfANOQqL8aGtlkqJ+v+
# tKEISvuxuN5cGjWOBUr1kjGGqOnYg/H3hv1sorfvtL+AZOW6vAQo4RVeuDoQHXYs
# 05hCUYwg2w==
# SIG # End signature block
